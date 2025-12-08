import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.nio.file.*;
import java.util.*;

/**
 * Test Excel Generator - Java 25.0.1
 * Creates test Excel files for validation of the ExcelToCsvExtractor.
 */
public class GenerateTestExcel {

    public record TestFile(String filename, String columnName, List<String> values) {}

    public static void main(String[] args) throws IOException {
        var testDir = Path.of(args.length > 0 ? args[0] : ".");

        var testFiles = List.of(
            new TestFile("test_emails.xlsx", "Email",
                List.of("john@example.com", "jane@example.com", "bob@example.com")),
            new TestFile("test_emails_upper.xlsx", "EMAIL",
                List.of("alice@test.org", "charlie@test.org")),
            new TestFile("test_no_email.xlsx", "Name",
                List.of("John Doe", "Jane Smith"))
        );

        for (var testFile : testFiles) {
            createExcel(testDir.resolve(testFile.filename()), testFile.columnName(), testFile.values());
        }

        System.out.println("""
            Test Excel files created in: %s

            Files generated:
              - test_emails.xlsx     (Email column - capitalized)
              - test_emails_upper.xlsx (EMAIL column - uppercase)
              - test_no_email.xlsx   (Name column - no email, tests error logging)
            """.formatted(testDir));
    }

    private static void createExcel(Path path, String columnName, List<String> values) throws IOException {
        try (var workbook = new XSSFWorkbook()) {
            var sheet = workbook.createSheet("Data");

            var header = sheet.createRow(0);
            header.createCell(0).setCellValue(columnName);

            for (var i = 0; i < values.size(); i++) {
                var row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue(values.get(i));
            }

            try (var out = Files.newOutputStream(path)) {
                workbook.write(out);
            }
        }
    }
}
