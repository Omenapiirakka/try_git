import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.nio.file.*;

/**
 * Test data generator for Excel to CSV Extractor v25.0.1
 * Creates sample Excel files for testing the extraction functionality.
 */
public class GenerateTestExcel {

    public static void main(String[] args) throws IOException {
        Path testDir = args.length > 0 ? Path.of(args[0]) : Path.of(".");

        System.out.println("Generating test Excel files in: " + testDir.toAbsolutePath());

        createExcel(testDir.resolve("test_emails.xlsx"), "Email",
            new String[]{"john@example.com", "jane@example.com", "bob@example.com"});

        createExcel(testDir.resolve("test_emails_upper.xlsx"), "EMAIL",
            new String[]{"alice@test.org", "charlie@test.org"});

        createExcel(testDir.resolve("test_no_email.xlsx"), "Name",
            new String[]{"John Doe", "Jane Smith"});

        createExcelWithNumbers(testDir.resolve("test_with_numbers.xlsx"), "ID",
            new long[]{1001, 1002, 1003, 1004});

        System.out.println("Test Excel files created successfully.");
    }

    private static void createExcel(Path path, String columnName, String[] values) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");

            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue(columnName);

            for (int i = 0; i < values.length; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(values[i]);
            }

            try (OutputStream out = Files.newOutputStream(path)) {
                workbook.write(out);
            }
        }
        System.out.println("  Created: " + path.getFileName());
    }

    private static void createExcelWithNumbers(Path path, String columnName, long[] values) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");

            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue(columnName);

            for (int i = 0; i < values.length; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(values[i]);
            }

            try (OutputStream out = Files.newOutputStream(path)) {
                workbook.write(out);
            }
        }
        System.out.println("  Created: " + path.getFileName());
    }
}
