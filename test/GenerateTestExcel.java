import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;

public class GenerateTestExcel {

    public static void main(String[] args) throws IOException {
        String testDir = args.length > 0 ? args[0] : ".";

        // Create test file with "Email" column (capitalized)
        createExcel(testDir + "/test_emails.xlsx", "Email",
            new String[]{"john@example.com", "jane@example.com", "bob@example.com"});

        // Create test file with "EMAIL" column (uppercase)
        createExcel(testDir + "/test_emails_upper.xlsx", "EMAIL",
            new String[]{"alice@test.org", "charlie@test.org"});

        // Create test file without email column (to test error logging)
        createExcel(testDir + "/test_no_email.xlsx", "Name",
            new String[]{"John Doe", "Jane Smith"});

        System.out.println("Test Excel files created in: " + testDir);
    }

    private static void createExcel(String path, String columnName, String[] values) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Data");

            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue(columnName);

            for (int i = 0; i < values.length; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(values[i]);
            }

            try (FileOutputStream out = new FileOutputStream(path)) {
                workbook.write(out);
            }
        }
    }
}
