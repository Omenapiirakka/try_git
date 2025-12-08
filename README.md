# Excel to CSV Column Extractor

Extracts a specific column from Excel files (.xlsx, .xls) and saves as CSV with semicolon separators.

## Build

```bash
mvn clean package
```

## Usage

```bash
java -jar target/excel-to-csv-1.0.0.jar <column-name> <folder-path>
```

### Parameters

- `column-name`: The column header to extract (case-insensitive)
- `folder-path`: Path to folder containing Excel files

### Example

```bash
java -jar target/excel-to-csv-1.0.0.jar email ./data
```

## Features

- Case-insensitive column matching (tries: lowercase, UPPERCASE, Capitalized)
- Processes all .xlsx and .xls files in the folder
- Creates CSV files with semicolon separators
- Logs errors to `error_log_<timestamp>.txt` when columns are not found

## Testing

A test Excel generator is included in the `test/` folder. After building:

```bash
# Copy dependencies for test compilation
mvn dependency:copy-dependencies

# Compile and run test generator
javac -cp "target/dependency/*" test/GenerateTestExcel.java
java -cp "test:target/dependency/*" GenerateTestExcel ./test

# Run extraction
java -jar target/excel-to-csv-1.0.0.jar email ./test
```

This creates 3 test files: one with "Email" column, one with "EMAIL" column, and one without an email column (to test error logging).
