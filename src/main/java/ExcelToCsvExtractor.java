import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.*;
import java.nio.file.*;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class ExcelToCsvExtractor {

    public static void main(String[] args) {
        if (args.length < 2) {
            System.out.println("Usage: java -jar excel-to-csv.jar <column-name> <folder-path>");
            System.exit(1);
        }

        String columnName = args[0];
        String folderPath = args[1];

        File folder = new File(folderPath);
        if (!folder.isDirectory()) {
            System.err.println("Error: " + folderPath + " is not a valid directory");
            System.exit(1);
        }

        File[] excelFiles = folder.listFiles((dir, name) ->
            name.endsWith(".xlsx") || name.endsWith(".xls"));

        if (excelFiles == null || excelFiles.length == 0) {
            System.out.println("No Excel files found in " + folderPath);
            return;
        }

        List<String> errors = new ArrayList<>();

        for (File file : excelFiles) {
            System.out.println("Processing: " + file.getName());
            try {
                List<String> values = extractColumn(file, columnName);
                if (values == null) {
                    String error = file.getName() + ": Column '" + columnName + "' not found";
                    errors.add(error);
                    System.err.println(error);
                } else {
                    writeCsv(file, values);
                    System.out.println("Created CSV for: " + file.getName());
                }
            } catch (Exception e) {
                String error = file.getName() + ": " + e.getMessage();
                errors.add(error);
                System.err.println(error);
            }
        }

        if (!errors.isEmpty()) {
            writeErrorLog(folder, errors);
        }
    }

    private static List<String> extractColumn(File file, String columnName) throws IOException {
        try (InputStream is = new FileInputStream(file);
             Workbook workbook = createWorkbook(file, is)) {

            Sheet sheet = workbook.getSheetAt(0);
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) return null;

            int columnIndex = findColumnIndex(headerRow, columnName);
            if (columnIndex == -1) return null;

            List<String> values = new ArrayList<>();
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                if (row != null) {
                    Cell cell = row.getCell(columnIndex);
                    values.add(getCellValue(cell));
                }
            }
            return values;
        }
    }

    private static Workbook createWorkbook(File file, InputStream is) throws IOException {
        if (file.getName().endsWith(".xlsx")) {
            return new XSSFWorkbook(is);
        }
        return new HSSFWorkbook(is);
    }

    private static int findColumnIndex(Row headerRow, String columnName) {
        String[] variants = {
            columnName.toLowerCase(),
            columnName.toUpperCase(),
            capitalize(columnName)
        };

        for (Cell cell : headerRow) {
            String header = getCellValue(cell).trim();
            for (String variant : variants) {
                if (header.equalsIgnoreCase(variant)) {
                    return cell.getColumnIndex();
                }
            }
        }
        return -1;
    }

    private static String capitalize(String str) {
        if (str == null || str.isEmpty()) return str;
        return str.substring(0, 1).toUpperCase() + str.substring(1).toLowerCase();
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

    private static void writeCsv(File excelFile, List<String> values) throws IOException {
        String csvName = excelFile.getName().replaceAll("\\.(xlsx|xls)$", ".csv");
        File csvFile = new File(excelFile.getParent(), csvName);

        try (PrintWriter writer = new PrintWriter(csvFile)) {
            for (String value : values) {
                writer.println(value + ";");
            }
        }
    }

    private static void writeErrorLog(File folder, List<String> errors) {
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyyMMdd_HHmmss"));
        File logFile = new File(folder, "error_log_" + timestamp + ".txt");

        try (PrintWriter writer = new PrintWriter(logFile)) {
            writer.println("Excel to CSV Extraction Error Log");
            writer.println("Generated: " + LocalDateTime.now());
            writer.println("-----------------------------------");
            for (String error : errors) {
                writer.println(error);
            }
            System.out.println("Error log written to: " + logFile.getName());
        } catch (IOException e) {
            System.err.println("Failed to write error log: " + e.getMessage());
        }
    }
}
