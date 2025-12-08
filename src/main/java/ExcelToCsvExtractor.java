import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;

/**
 * Excel to CSV Column Extractor v25.0.1
 * Extracts a specified column from Excel files and saves to CSV format.
 */
public class ExcelToCsvExtractor {

    private static final String VERSION = "25.0.1";
    private static final String CSV_DELIMITER = ";";

    public static void main(String[] args) {
        if (args.length < 2) {
            printUsage();
            System.exit(1);
        }

        String columnName = args[0];
        Path folderPath = Path.of(args[1]);

        if (!Files.isDirectory(folderPath)) {
            System.err.println("Error: " + folderPath + " is not a valid directory");
            System.exit(1);
        }

        processFolder(columnName, folderPath);
    }

    private static void printUsage() {
        System.out.println("Excel to CSV Extractor v" + VERSION);
        System.out.println("Usage: java -jar excel-to-csv.jar <column-name> <folder-path>");
        System.out.println();
        System.out.println("Arguments:");
        System.out.println("  column-name  Name of the column to extract (case-insensitive)");
        System.out.println("  folder-path  Path to folder containing Excel files (.xlsx, .xls)");
    }

    private static void processFolder(String columnName, Path folderPath) {
        List<Path> excelFiles = findExcelFiles(folderPath);

        if (excelFiles.isEmpty()) {
            System.out.println("No Excel files found in " + folderPath);
            return;
        }

        System.out.println("Found " + excelFiles.size() + " Excel file(s) to process");
        List<String> errors = new ArrayList<>();

        for (Path file : excelFiles) {
            processFile(file, columnName, errors);
        }

        if (!errors.isEmpty()) {
            writeErrorLog(folderPath, errors);
        }

        System.out.println("Processing complete. " + (excelFiles.size() - errors.size()) + " file(s) converted successfully.");
    }

    private static List<Path> findExcelFiles(Path folderPath) {
        try (var stream = Files.list(folderPath)) {
            return stream
                .filter(Files::isRegularFile)
                .filter(p -> {
                    String name = p.getFileName().toString().toLowerCase();
                    return name.endsWith(".xlsx") || name.endsWith(".xls");
                })
                .sorted()
                .collect(Collectors.toList());
        } catch (IOException e) {
            System.err.println("Error reading directory: " + e.getMessage());
            return Collections.emptyList();
        }
    }

    private static void processFile(Path file, String columnName, List<String> errors) {
        String fileName = file.getFileName().toString();
        System.out.println("Processing: " + fileName);

        try {
            List<String> values = extractColumn(file, columnName);
            if (values == null) {
                String error = fileName + ": Column '" + columnName + "' not found";
                errors.add(error);
                System.err.println("  " + error);
            } else {
                writeCsv(file, values);
                System.out.println("  Created CSV with " + values.size() + " row(s)");
            }
        } catch (Exception e) {
            String error = fileName + ": " + e.getMessage();
            errors.add(error);
            System.err.println("  Error: " + e.getMessage());
        }
    }

    private static List<String> extractColumn(Path file, String columnName) throws IOException {
        try (InputStream is = Files.newInputStream(file);
             Workbook workbook = createWorkbook(file, is)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                return null;
            }

            int columnIndex = findColumnIndex(headerRow, columnName);
            if (columnIndex == -1) {
                return null;
            }

            List<String> values = new ArrayList<>();
            int lastRowNum = sheet.getLastRowNum();

            for (int i = 1; i <= lastRowNum; i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(columnIndex);
                    String value = getCellValue(cell);
                    if (!value.isEmpty()) {
                        values.add(value);
                    }
                }
            }
            return values;
        }
    }

    private static Workbook createWorkbook(Path file, InputStream is) throws IOException {
        String fileName = file.getFileName().toString().toLowerCase();
        if (fileName.endsWith(".xlsx")) {
            return new XSSFWorkbook(is);
        }
        return new HSSFWorkbook(is);
    }

    private static int findColumnIndex(Row headerRow, String columnName) {
        return StreamSupport.stream(headerRow.spliterator(), false)
            .filter(cell -> getCellValue(cell).trim().equalsIgnoreCase(columnName))
            .findFirst()
            .map(Cell::getColumnIndex)
            .orElse(-1);
    }

    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }
        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> {
                double value = cell.getNumericCellValue();
                if (value == Math.floor(value) && !Double.isInfinite(value)) {
                    yield String.valueOf((long) value);
                }
                yield String.valueOf(value);
            }
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            case FORMULA -> evaluateFormula(cell);
            case BLANK -> "";
            default -> "";
        };
    }

    private static String evaluateFormula(Cell cell) {
        try {
            FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
            CellValue cellValue = evaluator.evaluate(cell);
            return switch (cellValue.getCellType()) {
                case STRING -> cellValue.getStringValue();
                case NUMERIC -> String.valueOf((long) cellValue.getNumberValue());
                case BOOLEAN -> String.valueOf(cellValue.getBooleanValue());
                default -> cell.getCellFormula();
            };
        } catch (Exception e) {
            return cell.getCellFormula();
        }
    }

    private static void writeCsv(Path excelFile, List<String> values) throws IOException {
        String csvName = excelFile.getFileName().toString().replaceAll("\\.(xlsx|xls)$", ".csv");
        Path csvFile = excelFile.getParent().resolve(csvName);

        try (BufferedWriter writer = Files.newBufferedWriter(csvFile)) {
            for (String value : values) {
                writer.write(escapeCsvValue(value) + CSV_DELIMITER);
                writer.newLine();
            }
        }
    }

    private static String escapeCsvValue(String value) {
        if (value.contains(CSV_DELIMITER) || value.contains("\"") || value.contains("\n")) {
            return "\"" + value.replace("\"", "\"\"") + "\"";
        }
        return value;
    }

    private static void writeErrorLog(Path folder, List<String> errors) {
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        Path logFile = folder.resolve("error_log_" + timestamp + ".txt");

        try (BufferedWriter writer = Files.newBufferedWriter(logFile)) {
            writer.write("Excel to CSV Extractor v" + VERSION + " - Error Log");
            writer.newLine();
            writer.write("Generated: " + LocalDateTime.now());
            writer.newLine();
            writer.write("=".repeat(50));
            writer.newLine();
            writer.newLine();

            for (String error : errors) {
                writer.write("- " + error);
                writer.newLine();
            }

            System.out.println("Error log written to: " + logFile.getFileName());
        } catch (IOException e) {
            System.err.println("Failed to write error log: " + e.getMessage());
        }
    }
}
