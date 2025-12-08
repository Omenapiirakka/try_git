import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import javax.xml.parsers.DocumentBuilderFactory;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.NodeList;

import java.io.*;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.concurrent.*;
import java.util.stream.*;

/**
 * Excel to CSV Column Extractor - Java 25.0.1
 * Extracts a specific column from Excel files and saves as CSV with semicolon separators.
 * Supports: .xlsx (OOXML), .xls (BIFF), .xml (Office 2003 SpreadsheetML)
 * Uses virtual threads for concurrent file processing.
 */
public class ExcelToCsvExtractor {

    private static final String CSV_SEPARATOR = ";";
    private static final String CSV_FOLDER = "CSV";
    private static final String LOGS_FOLDER = "logs";

    public sealed interface ExtractionResult permits ExtractionSuccess, ExtractionFailure {}

    public record ExtractionSuccess(Path sourceFile, Path csvFile, int rowCount, List<String> values) implements ExtractionResult {}

    public record ExtractionFailure(Path sourceFile, String errorMessage) implements ExtractionResult {}

    public record ColumnData(int index, String name) {}

    public record AppConfig(String columnName, Path folderPath, boolean mergeOutput) {}

    public static void main(String[] args) {
        // Launch GUI if no arguments provided
        if (args.length == 0) {
            ExcelToCsvGui.launch();
            return;
        }

        var config = parseArguments(args);
        if (config == null) {
            System.exit(1);
        }

        var results = processExcelFiles(config.folderPath(), config.columnName(), config.mergeOutput());
        handleResults(results, config.folderPath(), config.mergeOutput());
    }

    /**
     * Handles extraction results - prints summary and writes logs.
     * Public method for GUI access.
     */
    public static void handleResults(List<ExtractionResult> results, Path folder, boolean mergeOutput) {
        printResults(results, folder, mergeOutput);
    }

    private static AppConfig parseArguments(String[] args) {
        if (args.length < 2) {
            printUsage();
            return null;
        }

        boolean mergeOutput = false;
        String columnName = null;
        String folderPathStr = null;

        for (int i = 0; i < args.length; i++) {
            if (args[i].equals("--merge") || args[i].equals("-m")) {
                mergeOutput = true;
            } else if (columnName == null) {
                columnName = args[i];
            } else if (folderPathStr == null) {
                folderPathStr = args[i];
            }
        }

        if (columnName == null || folderPathStr == null) {
            printUsage();
            return null;
        }

        var folderPath = Path.of(folderPathStr);

        if (!Files.isDirectory(folderPath)) {
            System.err.println("Error: " + folderPath + " is not a valid directory");
            return null;
        }

        return new AppConfig(columnName, folderPath, mergeOutput);
    }

    private static void printUsage() {
        System.out.println("""
            Usage: java -jar excel-to-csv.jar [options] <column-name> <folder-path>

            Arguments:
              column-name   The column header to extract (case-insensitive)
              folder-path   Path to folder containing Excel files

            Options:
              --merge, -m   Generate a single merged CSV file containing all data

            Supported formats:
              .xlsx         Excel 2007+ (OOXML)
              .xls          Excel 97-2003 (BIFF)
              .xml          Office 2003 XML (SpreadsheetML)

            Output:
              CSV files are written to a 'CSV' subfolder
              Logs are written to 'CSV/logs' subfolder

            Example:
              java -jar excel-to-csv.jar email ./data
              java -jar excel-to-csv.jar --merge email ./data
            """);
    }

    /**
     * Process all Excel files in the specified folder.
     * Public method for GUI access.
     */
    public static List<ExtractionResult> processExcelFiles(Path folder, String columnName) {
        return processExcelFiles(folder, columnName, false);
    }

    /**
     * Process all Excel files in the specified folder.
     * When mergeOutput is true, individual CSV files are not written.
     */
    public static List<ExtractionResult> processExcelFiles(Path folder, String columnName, boolean mergeOutput) {
        try (var executor = Executors.newVirtualThreadPerTaskExecutor()) {
            var excelFiles = findExcelFiles(folder);

            if (excelFiles.isEmpty()) {
                System.out.println("No Excel files found in " + folder);
                return List.of();
            }

            // Create CSV output folder
            var csvFolder = folder.resolve(CSV_FOLDER);
            Files.createDirectories(csvFolder);

            var futures = excelFiles.stream()
                .map(file -> executor.submit(() -> processFile(file, columnName, csvFolder, mergeOutput)))
                .toList();

            return futures.stream()
                .map(ExcelToCsvExtractor::getFutureResult)
                .toList();

        } catch (Exception e) {
            System.err.println("Error processing files: " + e.getMessage());
            return List.of();
        }
    }

    private static List<Path> findExcelFiles(Path folder) {
        try (var stream = Files.list(folder)) {
            return stream
                .filter(Files::isRegularFile)
                .filter(path -> {
                    var name = path.getFileName().toString().toLowerCase();
                    return name.endsWith(".xlsx") || name.endsWith(".xls") || name.endsWith(".xml");
                })
                .toList();
        } catch (IOException e) {
            System.err.println("Error listing directory: " + e.getMessage());
            return List.of();
        }
    }

    private static ExtractionResult getFutureResult(Future<ExtractionResult> future) {
        try {
            return future.get();
        } catch (InterruptedException | ExecutionException e) {
            return new ExtractionFailure(Path.of("unknown"), e.getMessage());
        }
    }

    private static ExtractionResult processFile(Path file, String columnName, Path csvFolder, boolean mergeOutput) {
        System.out.println("Processing: " + file.getFileName());

        var fileName = file.getFileName().toString().toLowerCase();

        try {
            // Check file extension first, then verify content for xlsx/xls files
            if (fileName.endsWith(".xml") || isRawXmlFile(file)) {
                return processXmlSpreadsheet(file, columnName, csvFolder, mergeOutput);
            } else {
                return processExcelFile(file, columnName, csvFolder, mergeOutput);
            }
        } catch (Exception e) {
            return new ExtractionFailure(file, e.getMessage());
        }
    }

    /**
     * Detects if a file is a raw XML file by checking its content.
     * Office 2003 XML files with .xlsx or .xls extensions would fail in Apache POI,
     * so we detect them here and route to the XML parser instead.
     */
    private static boolean isRawXmlFile(Path file) {
        try (var is = Files.newInputStream(file)) {
            byte[] header = new byte[100];
            int bytesRead = is.read(header);
            if (bytesRead < 5) {
                return false;
            }

            var content = new String(header, 0, bytesRead, java.nio.charset.StandardCharsets.UTF_8).trim();

            // Check for XML declaration or root element
            if (content.startsWith("<?xml") || content.startsWith("<")) {
                // Additional check: verify it's not a valid ZIP file (OOXML .xlsx files are ZIP)
                // ZIP files start with PK (0x50 0x4B)
                if (header[0] == 0x50 && header[1] == 0x4B) {
                    return false; // It's a ZIP/OOXML file
                }
                return true;
            }
            return false;
        } catch (IOException e) {
            return false;
        }
    }

    private static ExtractionResult processExcelFile(Path file, String columnName, Path csvFolder, boolean mergeOutput) {
        try (var is = Files.newInputStream(file);
             var workbook = createWorkbook(file, is)) {

            var sheet = workbook.getSheetAt(0);
            var headerRow = sheet.getRow(0);

            if (headerRow == null) {
                return new ExtractionFailure(file, "No header row found");
            }

            var columnData = findColumn(headerRow, columnName);

            return switch (columnData) {
                case null -> new ExtractionFailure(file, "Column '" + columnName + "' not found");
                case ColumnData(var index, _) -> {
                    var values = extractColumnValues(sheet, index);
                    // Only write individual CSV if not merging
                    Path csvPath = mergeOutput ? null : writeCsv(file, values, csvFolder);
                    yield new ExtractionSuccess(file, csvPath, values.size(), values);
                }
            };

        } catch (IOException e) {
            return new ExtractionFailure(file, e.getMessage());
        }
    }

    private static ExtractionResult processXmlSpreadsheet(Path file, String columnName, Path csvFolder, boolean mergeOutput) {
        try {
            var factory = DocumentBuilderFactory.newInstance();
            factory.setNamespaceAware(true);
            var builder = factory.newDocumentBuilder();
            var doc = builder.parse(Files.newInputStream(file));

            // Find all Row elements
            var rows = findXmlRows(doc);

            if (rows.isEmpty()) {
                return new ExtractionFailure(file, "No rows found in XML spreadsheet");
            }

            // First row is header
            var headerRow = rows.get(0);
            var headerCells = getXmlRowCells(headerRow);

            var columnIndex = findXmlColumnIndex(headerCells, columnName);

            if (columnIndex < 0) {
                return new ExtractionFailure(file, "Column '" + columnName + "' not found");
            }

            // Extract values from data rows
            var values = new ArrayList<String>();
            for (int i = 1; i < rows.size(); i++) {
                var cells = getXmlRowCells(rows.get(i));
                if (columnIndex < cells.size()) {
                    values.add(cells.get(columnIndex));
                } else {
                    values.add("");
                }
            }

            // Only write individual CSV if not merging
            Path csvPath = mergeOutput ? null : writeCsv(file, values, csvFolder);
            return new ExtractionSuccess(file, csvPath, values.size(), values);

        } catch (Exception e) {
            return new ExtractionFailure(file, "XML parsing error: " + e.getMessage());
        }
    }

    private static List<Element> findXmlRows(Document doc) {
        var rows = new ArrayList<Element>();

        // Office 2003 SpreadsheetML uses ss:Row elements
        var nodeList = doc.getElementsByTagNameNS("urn:schemas-microsoft-com:office:spreadsheet", "Row");
        if (nodeList.getLength() == 0) {
            // Try without namespace
            nodeList = doc.getElementsByTagName("Row");
        }

        for (int i = 0; i < nodeList.getLength(); i++) {
            if (nodeList.item(i) instanceof Element elem) {
                rows.add(elem);
            }
        }

        return rows;
    }

    private static List<String> getXmlRowCells(Element rowElement) {
        var cells = new ArrayList<String>();

        // Office 2003 SpreadsheetML uses ss:Cell and ss:Data elements
        var cellNodes = rowElement.getElementsByTagNameNS("urn:schemas-microsoft-com:office:spreadsheet", "Cell");
        if (cellNodes.getLength() == 0) {
            cellNodes = rowElement.getElementsByTagName("Cell");
        }

        for (int i = 0; i < cellNodes.getLength(); i++) {
            if (cellNodes.item(i) instanceof Element cellElem) {
                // Check for ss:Index attribute (for sparse cells)
                var indexAttr = cellElem.getAttributeNS("urn:schemas-microsoft-com:office:spreadsheet", "Index");
                if (indexAttr != null && !indexAttr.isEmpty()) {
                    int targetIndex = Integer.parseInt(indexAttr) - 1; // 1-based to 0-based
                    while (cells.size() < targetIndex) {
                        cells.add("");
                    }
                }

                var dataNodes = cellElem.getElementsByTagNameNS("urn:schemas-microsoft-com:office:spreadsheet", "Data");
                if (dataNodes.getLength() == 0) {
                    dataNodes = cellElem.getElementsByTagName("Data");
                }

                if (dataNodes.getLength() > 0) {
                    cells.add(dataNodes.item(0).getTextContent().trim());
                } else {
                    cells.add("");
                }
            }
        }

        return cells;
    }

    private static int findXmlColumnIndex(List<String> headerCells, String columnName) {
        var variants = List.of(
            columnName.toLowerCase(),
            columnName.toUpperCase(),
            capitalize(columnName)
        );

        for (int i = 0; i < headerCells.size(); i++) {
            var header = headerCells.get(i).trim();
            for (var variant : variants) {
                if (header.equalsIgnoreCase(variant)) {
                    return i;
                }
            }
        }
        return -1;
    }

    private static Workbook createWorkbook(Path file, InputStream is) throws IOException {
        var fileName = file.getFileName().toString();
        return fileName.endsWith(".xlsx")
            ? new XSSFWorkbook(is)
            : new HSSFWorkbook(is);
    }

    private static ColumnData findColumn(Row headerRow, String columnName) {
        var variants = List.of(
            columnName.toLowerCase(),
            columnName.toUpperCase(),
            capitalize(columnName)
        );

        for (var cell : headerRow) {
            var header = getCellValue(cell).trim();
            for (var variant : variants) {
                if (header.equalsIgnoreCase(variant)) {
                    return new ColumnData(cell.getColumnIndex(), header);
                }
            }
        }
        return null;
    }

    private static String capitalize(String str) {
        if (str == null || str.isEmpty()) return str;
        return str.substring(0, 1).toUpperCase() + str.substring(1).toLowerCase();
    }

    private static List<String> extractColumnValues(Sheet sheet, int columnIndex) {
        return IntStream.rangeClosed(1, sheet.getLastRowNum())
            .mapToObj(sheet::getRow)
            .filter(Objects::nonNull)
            .map(row -> row.getCell(columnIndex))
            .map(ExcelToCsvExtractor::getCellValue)
            .toList();
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) return "";
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf((long) cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> cell.getCellFormula();
            default -> "";
        };
    }

    private static Path writeCsv(Path excelFile, List<String> values, Path csvFolder) throws IOException {
        var csvName = excelFile.getFileName().toString()
            .replaceAll("\\.(xlsx|xls|xml)$", ".csv");
        var csvPath = csvFolder.resolve(csvName);

        // Filter out empty values and write each value on a new line
        var filteredValues = values.stream()
            .filter(value -> value != null && !value.trim().isEmpty())
            .toList();

        // Build CSV content with each entry on a new line
        var content = new StringBuilder();
        for (var value : filteredValues) {
            content.append(escapeCsvValue(value))
                   .append(CSV_SEPARATOR)
                   .append(System.lineSeparator());
        }

        Files.writeString(csvPath, content.toString());

        // Validate the output CSV
        validateCsv(csvPath);

        return csvPath;
    }

    /**
     * Escapes a value for CSV format.
     * If the value contains the separator, quotes, or newlines, it must be quoted.
     */
    private static String escapeCsvValue(String value) {
        if (value == null) {
            return "";
        }

        // Check if the value needs to be quoted
        boolean needsQuoting = value.contains(CSV_SEPARATOR)
            || value.contains("\"")
            || value.contains("\n")
            || value.contains("\r");

        if (needsQuoting) {
            // Escape double quotes by doubling them and wrap in quotes
            return "\"" + value.replace("\"", "\"\"") + "\"";
        }

        return value;
    }

    /**
     * Validates that a CSV file is properly formatted.
     * Throws IOException if the CSV is invalid.
     */
    private static void validateCsv(Path csvPath) throws IOException {
        var content = Files.readString(csvPath);

        // Check that file is not empty (allow empty files for empty input)
        if (content.isEmpty()) {
            return; // Empty CSV is valid
        }

        var lines = content.split(System.lineSeparator(), -1);

        for (int lineNum = 0; lineNum < lines.length; lineNum++) {
            var line = lines[lineNum];

            // Skip empty lines at the end
            if (line.isEmpty() && lineNum == lines.length - 1) {
                continue;
            }

            // Validate quoted fields are properly closed
            if (!isValidCsvLine(line)) {
                throw new IOException("Invalid CSV at line " + (lineNum + 1) + ": unbalanced quotes");
            }
        }
    }

    /**
     * Checks if a CSV line has balanced quotes.
     */
    private static boolean isValidCsvLine(String line) {
        boolean inQuotes = false;

        for (int i = 0; i < line.length(); i++) {
            char c = line.charAt(i);

            if (c == '"') {
                // Check for escaped quote (doubled)
                if (inQuotes && i + 1 < line.length() && line.charAt(i + 1) == '"') {
                    i++; // Skip the escaped quote
                } else {
                    inQuotes = !inQuotes;
                }
            }
        }

        // Quotes should be balanced (not inside a quoted field at end of line)
        return !inQuotes;
    }

    private static void printResults(List<ExtractionResult> results, Path folder, boolean mergeOutput) {
        var successes = new ArrayList<ExtractionSuccess>();
        var failures = new ArrayList<ExtractionFailure>();

        for (var result : results) {
            switch (result) {
                case ExtractionSuccess s -> {
                    successes.add(s);
                    // Only print individual CSV message if not merging
                    if (s.csvFile() != null) {
                        System.out.println("Created CSV: " + s.csvFile().getFileName() + " (" + s.rowCount() + " rows)");
                    } else {
                        System.out.println("Extracted from: " + s.sourceFile().getFileName() + " (" + s.rowCount() + " rows)");
                    }
                }
                case ExtractionFailure f -> {
                    failures.add(f);
                    System.err.println("Error - " + f.sourceFile().getFileName() + ": " + f.errorMessage());
                }
            }
        }

        // Write merged CSV if requested
        if (mergeOutput && !successes.isEmpty()) {
            writeMergedCsv(folder, successes);
        }

        if (!failures.isEmpty()) {
            writeErrorLog(folder, failures);
        }

        System.out.println("""

            Summary:
              Successful: %d files
              Failed: %d files
            """.formatted(successes.size(), failures.size()));
    }

    private static void writeMergedCsv(Path folder, List<ExtractionSuccess> successes) {
        var csvFolder = folder.resolve(CSV_FOLDER);
        var timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        var mergedPath = csvFolder.resolve("merged_" + timestamp + ".csv");

        try {
            // Filter out empty values
            var allValues = successes.stream()
                .flatMap(s -> s.values().stream())
                .filter(value -> value != null && !value.trim().isEmpty())
                .toList();

            // Build CSV content with each entry on a new line
            var content = new StringBuilder();
            for (var value : allValues) {
                content.append(escapeCsvValue(value))
                       .append(CSV_SEPARATOR)
                       .append(System.lineSeparator());
            }

            Files.writeString(mergedPath, content.toString());

            // Validate the output CSV
            validateCsv(mergedPath);

            System.out.println("Created merged CSV: " + mergedPath.getFileName() + " (" + allValues.size() + " values)");
        } catch (IOException e) {
            System.err.println("Failed to write merged CSV: " + e.getMessage());
        }
    }

    private static void writeErrorLog(Path folder, List<ExtractionFailure> failures) {
        var logsFolder = folder.resolve(CSV_FOLDER).resolve(LOGS_FOLDER);

        try {
            Files.createDirectories(logsFolder);
        } catch (IOException e) {
            System.err.println("Failed to create logs folder: " + e.getMessage());
            return;
        }

        var timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        var logPath = logsFolder.resolve("error_log_" + timestamp + ".txt");

        var logContent = """
            Excel to CSV Extraction Error Log
            Generated: %s
            -----------------------------------
            %s
            """.formatted(
                LocalDateTime.now(),
                failures.stream()
                    .map(f -> f.sourceFile().getFileName() + ": " + f.errorMessage())
                    .collect(Collectors.joining(System.lineSeparator()))
            );

        try {
            Files.writeString(logPath, logContent);
            System.out.println("Error log written to: " + CSV_FOLDER + "/" + LOGS_FOLDER + "/" + logPath.getFileName());
        } catch (IOException e) {
            System.err.println("Failed to write error log: " + e.getMessage());
        }
    }
}
