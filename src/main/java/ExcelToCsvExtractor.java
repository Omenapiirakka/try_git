import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

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
 * Uses virtual threads for concurrent file processing.
 */
public class ExcelToCsvExtractor {

    private static final String CSV_SEPARATOR = ";";

    public sealed interface ExtractionResult permits ExtractionSuccess, ExtractionFailure {}

    public record ExtractionSuccess(Path sourceFile, Path csvFile, int rowCount) implements ExtractionResult {}

    public record ExtractionFailure(Path sourceFile, String errorMessage) implements ExtractionResult {}

    public record ColumnData(int index, String name) {}

    public static void main(String[] args) {
        if (args.length < 2) {
            System.out.println("""
                Usage: java -jar excel-to-csv.jar <column-name> <folder-path>

                Arguments:
                  column-name   The column header to extract (case-insensitive)
                  folder-path   Path to folder containing Excel files

                Example:
                  java -jar excel-to-csv.jar email ./data
                """);
            System.exit(1);
        }

        var columnName = args[0];
        var folderPath = Path.of(args[1]);

        if (!Files.isDirectory(folderPath)) {
            System.err.println("Error: " + folderPath + " is not a valid directory");
            System.exit(1);
        }

        var results = processExcelFiles(folderPath, columnName);
        printResults(results, folderPath);
    }

    private static List<ExtractionResult> processExcelFiles(Path folder, String columnName) {
        try (var executor = Executors.newVirtualThreadPerTaskExecutor()) {
            var excelFiles = findExcelFiles(folder);

            if (excelFiles.isEmpty()) {
                System.out.println("No Excel files found in " + folder);
                return List.of();
            }

            var futures = excelFiles.stream()
                .map(file -> executor.submit(() -> processFile(file, columnName)))
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
                    return name.endsWith(".xlsx") || name.endsWith(".xls");
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

    private static ExtractionResult processFile(Path file, String columnName) {
        System.out.println("Processing: " + file.getFileName());

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
                    var csvPath = writeCsv(file, values);
                    yield new ExtractionSuccess(file, csvPath, values.size());
                }
            };

        } catch (IOException e) {
            return new ExtractionFailure(file, e.getMessage());
        }
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

    private static Path writeCsv(Path excelFile, List<String> values) throws IOException {
        var csvName = excelFile.getFileName().toString()
            .replaceAll("\\.(xlsx|xls)$", ".csv");
        var csvPath = excelFile.getParent().resolve(csvName);

        var content = values.stream()
            .map(value -> value + CSV_SEPARATOR)
            .collect(Collectors.joining(System.lineSeparator()));

        Files.writeString(csvPath, content + System.lineSeparator());
        return csvPath;
    }

    private static void printResults(List<ExtractionResult> results, Path folder) {
        var successes = new ArrayList<ExtractionSuccess>();
        var failures = new ArrayList<ExtractionFailure>();

        for (var result : results) {
            switch (result) {
                case ExtractionSuccess s -> {
                    successes.add(s);
                    System.out.println("Created CSV: " + s.csvFile().getFileName() + " (" + s.rowCount() + " rows)");
                }
                case ExtractionFailure f -> {
                    failures.add(f);
                    System.err.println("Error - " + f.sourceFile().getFileName() + ": " + f.errorMessage());
                }
            }
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

    private static void writeErrorLog(Path folder, List<ExtractionFailure> failures) {
        var timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        var logPath = folder.resolve("error_log_" + timestamp + ".txt");

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
            System.out.println("Error log written to: " + logPath.getFileName());
        } catch (IOException e) {
            System.err.println("Failed to write error log: " + e.getMessage());
        }
    }
}
