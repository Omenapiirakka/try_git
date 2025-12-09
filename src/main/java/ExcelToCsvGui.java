import javax.swing.*;
import javax.swing.border.EmptyBorder;
import javax.swing.border.TitledBorder;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.io.PrintStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

/**
 * Swing GUI for Excel to CSV Column Extractor.
 * Provides a user-friendly interface for extracting columns from Excel files.
 */
public class ExcelToCsvGui extends JFrame {

    private JTextField folderField;
    private JTextField columnField;
    private JComboBox<ExcelToCsvExtractor.CsvDelimiter> delimiterComboBox;
    private JComboBox<ExcelToCsvExtractor.CsvEncoding> encodingComboBox;
    private JCheckBox mergeCheckbox;
    private JCheckBox scrambleCheckbox;
    private JButton browseButton;
    private JButton extractButton;
    private JTextArea outputArea;
    private JProgressBar progressBar;
    private JLabel statusLabel;

    public ExcelToCsvGui() {
        initializeUI();
    }

    private void initializeUI() {
        setTitle("Excel to CSV Column Extractor");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setMinimumSize(new Dimension(600, 500));

        // Main panel with padding
        var mainPanel = new JPanel(new BorderLayout(10, 10));
        mainPanel.setBorder(new EmptyBorder(15, 15, 15, 15));

        // Input panel at the top
        var inputPanel = createInputPanel();
        mainPanel.add(inputPanel, BorderLayout.NORTH);

        // Output area in the center
        var outputPanel = createOutputPanel();
        mainPanel.add(outputPanel, BorderLayout.CENTER);

        // Status bar at the bottom
        var statusPanel = createStatusPanel();
        mainPanel.add(statusPanel, BorderLayout.SOUTH);

        setContentPane(mainPanel);
        pack();
        setLocationRelativeTo(null); // Center on screen
    }

    private JPanel createInputPanel() {
        var panel = new JPanel(new GridBagLayout());
        panel.setBorder(new TitledBorder("Configuration"));
        var gbc = new GridBagConstraints();
        gbc.insets = new Insets(5, 5, 5, 5);
        gbc.fill = GridBagConstraints.HORIZONTAL;

        // Folder selection row
        gbc.gridx = 0;
        gbc.gridy = 0;
        gbc.weightx = 0;
        panel.add(new JLabel("Folder:"), gbc);

        gbc.gridx = 1;
        gbc.weightx = 1;
        folderField = new JTextField(30);
        folderField.setToolTipText("Path to folder containing Excel files (.xlsx, .xls, .xml)");
        panel.add(folderField, gbc);

        gbc.gridx = 2;
        gbc.weightx = 0;
        browseButton = new JButton("Browse...");
        browseButton.addActionListener(this::browseFolder);
        panel.add(browseButton, gbc);

        // Column name row
        gbc.gridx = 0;
        gbc.gridy = 1;
        gbc.weightx = 0;
        panel.add(new JLabel("Column Name:"), gbc);

        gbc.gridx = 1;
        gbc.gridwidth = 2;
        gbc.weightx = 1;
        columnField = new JTextField(20);
        columnField.setToolTipText("Column header to extract (case-insensitive)");
        panel.add(columnField, gbc);

        // Delimiter row
        gbc.gridx = 0;
        gbc.gridy = 2;
        gbc.gridwidth = 1;
        gbc.weightx = 0;
        panel.add(new JLabel("Delimiter:"), gbc);

        gbc.gridx = 1;
        gbc.gridwidth = 2;
        gbc.weightx = 1;
        delimiterComboBox = new JComboBox<>(ExcelToCsvExtractor.CsvDelimiter.values());
        delimiterComboBox.setSelectedItem(ExcelToCsvExtractor.CsvDelimiter.SEMICOLON);
        delimiterComboBox.setToolTipText("CSV field separator character");
        panel.add(delimiterComboBox, gbc);

        // Encoding row
        gbc.gridx = 0;
        gbc.gridy = 3;
        gbc.gridwidth = 1;
        gbc.weightx = 0;
        panel.add(new JLabel("Encoding:"), gbc);

        gbc.gridx = 1;
        gbc.gridwidth = 2;
        gbc.weightx = 1;
        encodingComboBox = new JComboBox<>(ExcelToCsvExtractor.CsvEncoding.values());
        encodingComboBox.setSelectedItem(ExcelToCsvExtractor.CsvEncoding.UTF_8);
        encodingComboBox.setToolTipText("CSV file encoding (UTF-8 with BOM recommended for Excel compatibility)");
        panel.add(encodingComboBox, gbc);

        // Options row
        gbc.gridx = 0;
        gbc.gridy = 4;
        gbc.gridwidth = 1;
        gbc.weightx = 0;
        panel.add(new JLabel("Options:"), gbc);

        gbc.gridx = 1;
        gbc.gridwidth = 2;
        mergeCheckbox = new JCheckBox("Merge all results into a single CSV file");
        mergeCheckbox.setToolTipText("Generate one merged CSV containing data from all processed files");
        panel.add(mergeCheckbox, gbc);

        // Debug scramble option row
        gbc.gridx = 1;
        gbc.gridy = 5;
        gbc.gridwidth = 2;
        scrambleCheckbox = new JCheckBox("Scramble output text (Debug)");
        scrambleCheckbox.setToolTipText("For debug purposes only: scrambles/randomizes text in CSV output to anonymize data");
        panel.add(scrambleCheckbox, gbc);

        // Extract button row
        gbc.gridx = 0;
        gbc.gridy = 6;
        gbc.gridwidth = 3;
        gbc.anchor = GridBagConstraints.CENTER;
        gbc.fill = GridBagConstraints.NONE;
        extractButton = new JButton("Extract Column");
        extractButton.setFont(extractButton.getFont().deriveFont(Font.BOLD, 14f));
        extractButton.addActionListener(this::startExtraction);
        panel.add(extractButton, gbc);

        return panel;
    }

    private JPanel createOutputPanel() {
        var panel = new JPanel(new BorderLayout());
        panel.setBorder(new TitledBorder("Output"));

        outputArea = new JTextArea();
        outputArea.setEditable(false);
        outputArea.setFont(new Font(Font.MONOSPACED, Font.PLAIN, 12));
        outputArea.setMargin(new Insets(5, 5, 5, 5));

        var scrollPane = new JScrollPane(outputArea);
        scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS);
        panel.add(scrollPane, BorderLayout.CENTER);

        return panel;
    }

    private JPanel createStatusPanel() {
        var panel = new JPanel(new BorderLayout(10, 5));

        progressBar = new JProgressBar();
        progressBar.setStringPainted(true);
        progressBar.setString("Ready");
        panel.add(progressBar, BorderLayout.CENTER);

        statusLabel = new JLabel("Select a folder and enter column name to begin");
        statusLabel.setFont(statusLabel.getFont().deriveFont(Font.ITALIC));
        panel.add(statusLabel, BorderLayout.SOUTH);

        return panel;
    }

    private void browseFolder(ActionEvent e) {
        var chooser = new JFileChooser();
        chooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        chooser.setDialogTitle("Select Folder Containing Excel Files");

        // Start from current folder field value or user home
        var currentPath = folderField.getText().trim();
        if (!currentPath.isEmpty() && Files.isDirectory(Path.of(currentPath))) {
            chooser.setCurrentDirectory(Path.of(currentPath).toFile());
        }

        if (chooser.showOpenDialog(this) == JFileChooser.APPROVE_OPTION) {
            folderField.setText(chooser.getSelectedFile().getAbsolutePath());
        }
    }

    private void startExtraction(ActionEvent e) {
        var folderPath = folderField.getText().trim();
        var columnName = columnField.getText().trim();

        // Validate inputs
        if (folderPath.isEmpty()) {
            showError("Please select a folder containing Excel files.");
            return;
        }

        if (columnName.isEmpty()) {
            showError("Please enter a column name to extract.");
            return;
        }

        var folder = Path.of(folderPath);
        if (!Files.isDirectory(folder)) {
            showError("The specified path is not a valid directory.");
            return;
        }

        // Disable UI during extraction
        setInputsEnabled(false);
        outputArea.setText("");
        progressBar.setIndeterminate(true);
        progressBar.setString("Processing...");
        statusLabel.setText("Extracting column '" + columnName + "' from Excel files...");

        // Run extraction in background thread
        var delimiter = (ExcelToCsvExtractor.CsvDelimiter) delimiterComboBox.getSelectedItem();
        var encoding = (ExcelToCsvExtractor.CsvEncoding) encodingComboBox.getSelectedItem();
        var worker = new ExtractionWorker(folder, columnName, mergeCheckbox.isSelected(), scrambleCheckbox.isSelected(), delimiter, encoding);
        worker.execute();
    }

    private void setInputsEnabled(boolean enabled) {
        folderField.setEnabled(enabled);
        columnField.setEnabled(enabled);
        delimiterComboBox.setEnabled(enabled);
        encodingComboBox.setEnabled(enabled);
        mergeCheckbox.setEnabled(enabled);
        scrambleCheckbox.setEnabled(enabled);
        browseButton.setEnabled(enabled);
        extractButton.setEnabled(enabled);
    }

    private void showError(String message) {
        JOptionPane.showMessageDialog(this, message, "Input Error", JOptionPane.ERROR_MESSAGE);
    }

    private void appendOutput(String text) {
        SwingUtilities.invokeLater(() -> {
            outputArea.append(text + "\n");
            outputArea.setCaretPosition(outputArea.getDocument().getLength());
        });
    }

    /**
     * Background worker for extraction process.
     */
    private class ExtractionWorker extends SwingWorker<ExtractionSummary, String> {

        private final Path folder;
        private final String columnName;
        private final boolean mergeOutput;
        private final boolean scrambleOutput;
        private final ExcelToCsvExtractor.CsvDelimiter delimiter;
        private final ExcelToCsvExtractor.CsvEncoding encoding;

        ExtractionWorker(Path folder, String columnName, boolean mergeOutput, boolean scrambleOutput, ExcelToCsvExtractor.CsvDelimiter delimiter, ExcelToCsvExtractor.CsvEncoding encoding) {
            this.folder = folder;
            this.columnName = columnName;
            this.mergeOutput = mergeOutput;
            this.scrambleOutput = scrambleOutput;
            this.delimiter = delimiter;
            this.encoding = encoding;
        }

        @Override
        protected ExtractionSummary doInBackground() {
            // Redirect System.out and System.err to capture output
            var originalOut = System.out;
            var originalErr = System.err;

            try {
                // Create custom print streams that publish to the worker
                var guiOut = new PrintStream(new GuiOutputStream(this::publish, false));
                var guiErr = new PrintStream(new GuiOutputStream(this::publish, true));
                System.setOut(guiOut);
                System.setErr(guiErr);

                // Run extraction
                var results = ExcelToCsvExtractor.processExcelFiles(folder, columnName, mergeOutput, scrambleOutput, delimiter, encoding);
                ExcelToCsvExtractor.handleResults(results, folder, mergeOutput, scrambleOutput, delimiter, encoding);

                // Count successes and failures
                int successCount = 0;
                int failureCount = 0;
                for (var result : results) {
                    if (result instanceof ExcelToCsvExtractor.ExtractionSuccess) {
                        successCount++;
                    } else {
                        failureCount++;
                    }
                }

                return new ExtractionSummary(successCount, failureCount);

            } finally {
                System.setOut(originalOut);
                System.setErr(originalErr);
            }
        }

        @Override
        protected void process(List<String> chunks) {
            for (var chunk : chunks) {
                outputArea.append(chunk);
                outputArea.setCaretPosition(outputArea.getDocument().getLength());
            }
        }

        @Override
        protected void done() {
            setInputsEnabled(true);
            progressBar.setIndeterminate(false);

            try {
                var summary = get();
                progressBar.setValue(100);
                progressBar.setString("Complete");

                if (summary.successCount() == 0 && summary.failureCount() == 0) {
                    statusLabel.setText("No Excel files found in the selected folder.");
                } else {
                    statusLabel.setText(String.format("Completed: %d successful, %d failed",
                        summary.successCount(), summary.failureCount()));
                }

                if (summary.successCount() > 0) {
                    appendOutput("\nCSV files saved to: " + folder.resolve("CSV"));
                }

            } catch (Exception ex) {
                progressBar.setString("Error");
                statusLabel.setText("Error: " + ex.getMessage());
                appendOutput("\nError: " + ex.getMessage());
            }
        }
    }

    /**
     * Summary of extraction results.
     */
    private record ExtractionSummary(int successCount, int failureCount) {}

    /**
     * Custom OutputStream that publishes text to SwingWorker.
     */
    private static class GuiOutputStream extends java.io.OutputStream {
        private final java.util.function.Consumer<String> publisher;
        private final boolean isError;
        private final StringBuilder buffer = new StringBuilder();

        GuiOutputStream(java.util.function.Consumer<String> publisher, boolean isError) {
            this.publisher = publisher;
            this.isError = isError;
        }

        @Override
        public void write(int b) {
            char c = (char) b;
            buffer.append(c);
            if (c == '\n') {
                flush();
            }
        }

        @Override
        public void flush() {
            if (!buffer.isEmpty()) {
                var text = buffer.toString();
                publisher.accept(isError ? "[ERROR] " + text : text);
                buffer.setLength(0);
            }
        }
    }

    /**
     * Launch the GUI application.
     */
    public static void launch() {
        // Set look and feel to system default
        try {
            UIManager.setLookAndFeel(UIManager.getSystemLookAndFeelClassName());
        } catch (Exception e) {
            // Fall back to default look and feel
        }

        SwingUtilities.invokeLater(() -> {
            var gui = new ExcelToCsvGui();
            gui.setVisible(true);
        });
    }
}
