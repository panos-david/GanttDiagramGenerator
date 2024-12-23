package app;

import dom.gantt.TaskAbstract;
import dom.gantt.TaskConcrete;
import service.IMainController;
import util.FileTypes;
import util.ProjectInfo;

import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableRowSorter;
import java.awt.*;
import java.awt.event.*;
import java.io.File;
import java.util.List;
import java.util.*;
import org.apache.poi.ss.usermodel.IndexedColors;

public class MainAppSwing extends JFrame {

    private IMainController appController;
    private JTable table;
    private DefaultTableModel tableModel;
    private JFileChooser fileChooser;

    private String originalFileName;
    private StyleSettings headerStyleSettings = new StyleSettings();
    private StyleSettings topBarStyleSettings = new StyleSettings();
    private StyleSettings topDataStyleSettings = new StyleSettings();
    private StyleSettings nonTopBarStyleSettings = new StyleSettings();
    private StyleSettings nonTopDataStyleSettings = new StyleSettings();
    private StyleSettings normalStyleSettings = new StyleSettings();

    public MainAppSwing() {
        appController = new ApplicationController();
        fileChooser = new JFileChooser();
        initUI();
    }

    private void initUI() {
        setTitle("Task Manager");
        setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        setSize(900, 600);
        setLocationRelativeTo(null);

        JMenuBar menuBar = new JMenuBar();

        JMenu menuFile = new JMenu("File");
        JMenuItem menuItemLoad = new JMenuItem("Load");
        JMenuItem menuItemSave = new JMenuItem("Save");
        menuFile.add(menuItemLoad);
        menuFile.add(menuItemSave);
        menuBar.add(menuFile);

        JMenu menuSettings = new JMenu("Settings");
        JMenuItem menuItemStyles = new JMenuItem("Styles");
        menuSettings.add(menuItemStyles);
        menuBar.add(menuSettings);

        setJMenuBar(menuBar);

        String[] columnNames = {"ID", "Name", "Container ID", "Start Day", "End Day", "Cost", "Effort"};

        tableModel = new DefaultTableModel(columnNames, 0) {
            @Override
            public boolean isCellEditable(int row, int column) {
                return true;
            }

            @Override
            public Class<?> getColumnClass(int columnIndex) {
                switch (columnIndex) {
                    case 0: // ID
                    case 2: // Container ID
                    case 3: // Start Day
                    case 4: // End Day
                        return Integer.class;
                    case 5: // Cost
                    case 6: // Effort
                        return Double.class;
                    default:
                        return String.class;
                }
            }
        };

        table = new JTable(tableModel);

        TableRowSorter<DefaultTableModel> sorter = new TableRowSorter<>(tableModel);
        table.setRowSorter(sorter);

        JScrollPane scrollPane = new JScrollPane(table);
        add(scrollPane, BorderLayout.CENTER);

        
        JPanel controlsPanel = new JPanel();
        controlsPanel.setLayout(new FlowLayout(FlowLayout.LEFT));

        JButton btnGetAllTasks = new JButton("Show All Tasks");
        JButton btnGetTopLevelTasks = new JButton("Show Top-Level Tasks");
        JButton btnGetTasksInRange = new JButton("Show Tasks In Range");
        JLabel lblFirstId = new JLabel("First ID:");
        JTextField txtFirstId = new JTextField(5);
        JLabel lblLastId = new JLabel("Last ID:");
        JTextField txtLastId = new JTextField(5);
        JButton btnExport = new JButton("Export to Excel");

        controlsPanel.add(btnGetAllTasks);
        controlsPanel.add(btnGetTopLevelTasks);
        controlsPanel.add(lblFirstId);
        controlsPanel.add(txtFirstId);
        controlsPanel.add(lblLastId);
        controlsPanel.add(txtLastId);
        controlsPanel.add(btnGetTasksInRange);
        controlsPanel.add(btnExport);

        add(controlsPanel, BorderLayout.SOUTH);


        menuItemLoad.addActionListener(e -> {
            int result = fileChooser.showOpenDialog(MainAppSwing.this);
            if (result == JFileChooser.APPROVE_OPTION) {
                File selectedFile = fileChooser.getSelectedFile();
                originalFileName = selectedFile.getName(); 
                FileTypes fileType = getFileType(selectedFile);
                if (fileType == null) {
                    JOptionPane.showMessageDialog(MainAppSwing.this, "Unsupported file type.", "Error", JOptionPane.ERROR_MESSAGE);
                    return;
                }
                List<String> loadedTasks = appController.load(selectedFile.getAbsolutePath(), fileType);
                if (loadedTasks != null) {
                    updateTable(appController.getAllTasks());
                } else {
                    JOptionPane.showMessageDialog(MainAppSwing.this, "Failed to load file.", "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
        });

        btnGetAllTasks.addActionListener(e -> updateTable(appController.getAllTasks()));

        btnGetTopLevelTasks.addActionListener(e -> updateTable(appController.getTopLevelTasksOnly()));

        btnGetTasksInRange.addActionListener(e -> {
            try {
                int firstId = Integer.parseInt(txtFirstId.getText());
                int lastId = Integer.parseInt(txtLastId.getText());
                updateTable(appController.getTasksInRange(firstId, lastId));
            } catch (NumberFormatException ex) {
                JOptionPane.showMessageDialog(MainAppSwing.this, "Please enter valid integer values for IDs.", "Error", JOptionPane.ERROR_MESSAGE);
            }
        });

        menuItemSave.addActionListener(e -> {
            List<TaskAbstract> tasks = appController.getAllTasks();
            if (tasks == null || tasks.isEmpty()) {
                JOptionPane.showMessageDialog(MainAppSwing.this, "No tasks to save.", "Information", JOptionPane.INFORMATION_MESSAGE);
                return;
            }
            FileNameExtensionFilter excelFilter = new FileNameExtensionFilter("Excel Files", "xlsx");
            fileChooser.setFileFilter(excelFilter);
            if (originalFileName != null) {
                String defaultFileName = getBaseName(originalFileName) + ".xlsx";
                fileChooser.setSelectedFile(new File(defaultFileName));
            }
            int result = fileChooser.showSaveDialog(MainAppSwing.this);
            if (result == JFileChooser.APPROVE_OPTION) {
                File saveFile = fileChooser.getSelectedFile();
                if (!saveFile.getName().toLowerCase().endsWith(".xlsx")) {
                    saveFile = new File(saveFile.getAbsolutePath() + ".xlsx");
                }
                ProjectInfo projectInfo = appController.prepareTargetWorkbook(FileTypes.XLSX, saveFile.getAbsolutePath());
                if (projectInfo != null) {
                    createStyles(); 
                    boolean success = appController.createNewSheet("Tasks", tasks, "HeaderStyle", "TopBarStyle",
                            "TopDataStyle", "NonTopBarStyle", "NonTopDataStyle", "NormalStyle");
                    if (success) {
                        JOptionPane.showMessageDialog(MainAppSwing.this, "Tasks saved successfully at " + saveFile.getAbsolutePath(), "Success", JOptionPane.INFORMATION_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(MainAppSwing.this, "Failed to save tasks.", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                } else {
                    JOptionPane.showMessageDialog(MainAppSwing.this, "Failed to prepare target workbook.", "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
            fileChooser.setFileFilter(null);
            fileChooser.setSelectedFile(null);
        });

        btnExport.addActionListener(e -> {
            List<TaskAbstract> tasks = getCurrentTasksFromTable();
            if (tasks == null || tasks.isEmpty()) {
                JOptionPane.showMessageDialog(MainAppSwing.this, "No tasks to export.", "Information", JOptionPane.INFORMATION_MESSAGE);
                return;
            }
            FileNameExtensionFilter excelFilter = new FileNameExtensionFilter("Excel Files", "xlsx");
            fileChooser.setFileFilter(excelFilter);
            if (originalFileName != null) {
                String defaultFileName = getBaseName(originalFileName) + "_Export.xlsx";
                fileChooser.setSelectedFile(new File(defaultFileName));
            }
            int result = fileChooser.showSaveDialog(MainAppSwing.this);
            if (result == JFileChooser.APPROVE_OPTION) {
                File saveFile = fileChooser.getSelectedFile();
                if (!saveFile.getName().toLowerCase().endsWith(".xlsx")) {
                    saveFile = new File(saveFile.getAbsolutePath() + ".xlsx");
                }
                ProjectInfo projectInfo = appController.prepareTargetWorkbook(FileTypes.XLSX, saveFile.getAbsolutePath());
                if (projectInfo != null) {
                    createStyles(); 
                    boolean success = appController.createNewSheet("Tasks", tasks, "HeaderStyle", "TopBarStyle",
                            "TopDataStyle", "NonTopBarStyle", "NonTopDataStyle", "NormalStyle");
                    if (success) {
                        JOptionPane.showMessageDialog(MainAppSwing.this, "Tasks exported successfully at " + saveFile.getAbsolutePath(), "Success", JOptionPane.INFORMATION_MESSAGE);
                    } else {
                        JOptionPane.showMessageDialog(MainAppSwing.this, "Failed to export tasks.", "Error", JOptionPane.ERROR_MESSAGE);
                    }
                } else {
                    JOptionPane.showMessageDialog(MainAppSwing.this, "Failed to prepare target workbook.", "Error", JOptionPane.ERROR_MESSAGE);
                }
            }
            fileChooser.setFileFilter(null);
            fileChooser.setSelectedFile(null);
        });

        menuItemStyles.addActionListener(e -> openStyleSettingsDialog());
    }

    private void updateTable(List<TaskAbstract> tasks) {
        tableModel.setRowCount(0);
        if (tasks != null) {
            for (TaskAbstract task : tasks) {
                Object[] rowData = {
                        task.getId(),
                        task.getName(),
                        task.getContainerId(),
                        task.getStartDay(),
                        task.getEndDay(),
                        task.getCost(),
                        task.getEffort()
                };
                tableModel.addRow(rowData);
            }
        }
    }

    private List<TaskAbstract> getCurrentTasksFromTable() {
        List<TaskAbstract> tasks = new ArrayList<>();
        int rowCount = tableModel.getRowCount();
        for (int i = 0; i < rowCount; i++) {
            int id = (Integer) tableModel.getValueAt(i, 0);
            String name = (String) tableModel.getValueAt(i, 1);
            int containerId = (Integer) tableModel.getValueAt(i, 2);
            Integer startDay = (Integer) tableModel.getValueAt(i, 3);
            Integer endDay = (Integer) tableModel.getValueAt(i, 4);
            double cost = (Double) tableModel.getValueAt(i, 5);
            double effort = (Double) tableModel.getValueAt(i, 6);
            TaskAbstract task = new TaskConcrete(id, name, containerId, startDay, endDay, cost, effort);
            tasks.add(task);
        }
        return tasks;
    }

    private FileTypes getFileType(File file) {
        String fileName = file.getName().toLowerCase();
        if (fileName.endsWith(".xlsx")) {
            return FileTypes.XLSX;
        } else if (fileName.endsWith(".xls")) {
            return FileTypes.XLS;
        } else if (fileName.endsWith(".csv")) {
            return FileTypes.CSV;
        } else if (fileName.endsWith(".tsv")) {
            return FileTypes.TSV;
        }
        return null;
    }

    private String getBaseName(String fileName) {
        int index = fileName.lastIndexOf('.');
        if (index > 0 && index <= fileName.length() - 2) {
            return fileName.substring(0, index);
        }
        return fileName;
    }

    private void openStyleSettingsDialog() {
        JDialog dialog = new JDialog(this, "Style Settings", true);
        dialog.setSize(600, 500);
        dialog.setLocationRelativeTo(this);

        JTabbedPane tabbedPane = new JTabbedPane();

        tabbedPane.addTab("Header Style", createStyleSettingsPanel(headerStyleSettings));
        tabbedPane.addTab("Top-Level Bar Style", createStyleSettingsPanel(topBarStyleSettings));
        tabbedPane.addTab("Top-Level Data Style", createStyleSettingsPanel(topDataStyleSettings));
        tabbedPane.addTab("Non-Top-Level Bar Style", createStyleSettingsPanel(nonTopBarStyleSettings));
        tabbedPane.addTab("Non-Top-Level Data Style", createStyleSettingsPanel(nonTopDataStyleSettings));
        tabbedPane.addTab("Normal Style", createStyleSettingsPanel(normalStyleSettings));

        JButton btnOk = new JButton("OK");
        btnOk.addActionListener(e -> {
            dialog.dispose();
        });

        JButton btnCancel = new JButton("Cancel");
        btnCancel.addActionListener(e -> {
            dialog.dispose();
        });

        JPanel buttonsPanel = new JPanel();
        buttonsPanel.add(btnOk);
        buttonsPanel.add(btnCancel);

        dialog.add(tabbedPane, BorderLayout.CENTER);
        dialog.add(buttonsPanel, BorderLayout.SOUTH);

        dialog.setVisible(true);
    }

    private JPanel createStyleSettingsPanel(StyleSettings styleSettings) {
        JPanel panel = new JPanel(new GridLayout(9, 2));

        panel.add(new JLabel("Font Name:"));
        String[] fontNames = GraphicsEnvironment.getLocalGraphicsEnvironment().getAvailableFontFamilyNames();
        JComboBox<String> fontNameCombo = new JComboBox<>(fontNames);
        fontNameCombo.setSelectedItem(styleSettings.fontName);
        panel.add(fontNameCombo);

        panel.add(new JLabel("Font Size:"));
        Integer[] fontSizes = new Integer[30];
        for (int i = 0; i < fontSizes.length; i++) {
            fontSizes[i] = i + 8;
        }
        JComboBox<Integer> fontSizeCombo = new JComboBox<>(fontSizes);
        fontSizeCombo.setSelectedItem(styleSettings.fontSize);
        panel.add(fontSizeCombo);

        
        panel.add(new JLabel("Bold:"));
        JCheckBox boldCheckBox = new JCheckBox();
        boldCheckBox.setSelected(styleSettings.bold);
        panel.add(boldCheckBox);

        panel.add(new JLabel("Italic:"));
        JCheckBox italicCheckBox = new JCheckBox();
        italicCheckBox.setSelected(styleSettings.italic);
        panel.add(italicCheckBox);

        panel.add(new JLabel("Wrap Text:"));
        JCheckBox wrapTextCheckBox = new JCheckBox();
        wrapTextCheckBox.setSelected(styleSettings.wrapText);
        panel.add(wrapTextCheckBox);

        panel.add(new JLabel("Font Color:"));
        JButton fontColorButton = new JButton("Select Color");
        fontColorButton.setBackground(styleSettings.fontColor);
        fontColorButton.addActionListener(e -> {
            Color color = JColorChooser.showDialog(this, "Choose Font Color", styleSettings.fontColor);
            if (color != null) {
                styleSettings.fontColor = color;
                fontColorButton.setBackground(color);
            }
        });
        panel.add(fontColorButton);

        panel.add(new JLabel("Fill Color:"));
        JButton fillColorButton = new JButton("Select Color");
        fillColorButton.setBackground(styleSettings.fillColor);
        fillColorButton.addActionListener(e -> {
            Color color = JColorChooser.showDialog(this, "Choose Fill Color", styleSettings.fillColor);
            if (color != null) {
                styleSettings.fillColor = color;
                fillColorButton.setBackground(color);
            }
        });
        panel.add(fillColorButton);

        panel.add(new JLabel("Fill Pattern:"));
        String[] fillPatterns = {"NO_FILL", "SOLID_FOREGROUND", "FINE_DOTS", "ALT_BARS", "SPARSE_DOTS", "THICK_HORZ_BANDS",
                "THICK_VERT_BANDS", "THICK_BACKWARD_DIAG", "THICK_FORWARD_DIAG", "BIG_SPOTS", "BRICKS", "THIN_HORZ_BANDS",
                "THIN_VERT_BANDS", "THIN_BACKWARD_DIAG", "THIN_FORWARD_DIAG", "SQUARES", "DIAMONDS", "LESS_DOTS", "LEAST_DOTS"};
        JComboBox<String> fillPatternCombo = new JComboBox<>(fillPatterns);
        fillPatternCombo.setSelectedItem(styleSettings.fillPattern);
        panel.add(fillPatternCombo);

        panel.add(new JLabel("Alignment:"));
        String[] alignments = {"LEFT", "CENTER", "RIGHT", "JUSTIFY"};
        JComboBox<String> alignmentCombo = new JComboBox<>(alignments);
        alignmentCombo.setSelectedItem(styleSettings.alignment);
        panel.add(alignmentCombo);

        fontNameCombo.addActionListener(e -> styleSettings.fontName = (String) fontNameCombo.getSelectedItem());
        fontSizeCombo.addActionListener(e -> styleSettings.fontSize = (Integer) fontSizeCombo.getSelectedItem());
        boldCheckBox.addActionListener(e -> styleSettings.bold = boldCheckBox.isSelected());
        italicCheckBox.addActionListener(e -> styleSettings.italic = italicCheckBox.isSelected());
        wrapTextCheckBox.addActionListener(e -> styleSettings.wrapText = wrapTextCheckBox.isSelected());
        fillPatternCombo.addActionListener(e -> styleSettings.fillPattern = (String) fillPatternCombo.getSelectedItem());
        alignmentCombo.addActionListener(e -> styleSettings.alignment = (String) alignmentCombo.getSelectedItem());

        return panel;
    }

    private void createStyles() {
        appController.createDefaultStyles();

        appController.addFontedStyle("HeaderStyle",
                getClosestIndexedColor(headerStyleSettings.fontColor).getIndex(),
                (short) headerStyleSettings.fontSize,
                headerStyleSettings.fontName,
                headerStyleSettings.bold,
                headerStyleSettings.italic,
                false,
                getClosestIndexedColor(headerStyleSettings.fillColor).getIndex(),
                headerStyleSettings.fillPattern,
                headerStyleSettings.alignment,
                headerStyleSettings.wrapText);

        appController.addFontedStyle("TopBarStyle",
                getClosestIndexedColor(topBarStyleSettings.fontColor).getIndex(),
                (short) topBarStyleSettings.fontSize,
                topBarStyleSettings.fontName,
                topBarStyleSettings.bold,
                topBarStyleSettings.italic,
                false,
                getClosestIndexedColor(topBarStyleSettings.fillColor).getIndex(),
                topBarStyleSettings.fillPattern,
                topBarStyleSettings.alignment,
                topBarStyleSettings.wrapText);

        appController.addFontedStyle("TopDataStyle",
                getClosestIndexedColor(topDataStyleSettings.fontColor).getIndex(),
                (short) topDataStyleSettings.fontSize,
                topDataStyleSettings.fontName,
                topDataStyleSettings.bold,
                topDataStyleSettings.italic,
                false,
                getClosestIndexedColor(topDataStyleSettings.fillColor).getIndex(),
                topDataStyleSettings.fillPattern,
                topDataStyleSettings.alignment,
                topDataStyleSettings.wrapText);

        appController.addFontedStyle("NonTopBarStyle",
                getClosestIndexedColor(nonTopBarStyleSettings.fontColor).getIndex(),
                (short) nonTopBarStyleSettings.fontSize,
                nonTopBarStyleSettings.fontName,
                nonTopBarStyleSettings.bold,
                nonTopBarStyleSettings.italic,
                false,
                getClosestIndexedColor(nonTopBarStyleSettings.fillColor).getIndex(),
                nonTopBarStyleSettings.fillPattern,
                nonTopBarStyleSettings.alignment,
                nonTopBarStyleSettings.wrapText);

        appController.addFontedStyle("NonTopDataStyle",
                getClosestIndexedColor(nonTopDataStyleSettings.fontColor).getIndex(),
                (short) nonTopDataStyleSettings.fontSize,
                nonTopDataStyleSettings.fontName,
                nonTopDataStyleSettings.bold,
                nonTopDataStyleSettings.italic,
                false,
                getClosestIndexedColor(nonTopDataStyleSettings.fillColor).getIndex(),
                nonTopDataStyleSettings.fillPattern,
                nonTopDataStyleSettings.alignment,
                nonTopDataStyleSettings.wrapText);

        appController.addFontedStyle("NormalStyle",
                getClosestIndexedColor(normalStyleSettings.fontColor).getIndex(),
                (short) normalStyleSettings.fontSize,
                normalStyleSettings.fontName,
                normalStyleSettings.bold,
                normalStyleSettings.italic,
                false,
                getClosestIndexedColor(normalStyleSettings.fillColor).getIndex(),
                normalStyleSettings.fillPattern,
                normalStyleSettings.alignment,
                normalStyleSettings.wrapText);
    }

    private IndexedColors getClosestIndexedColor(Color color) {
        if (color == null) {
            return IndexedColors.AUTOMATIC;
        }
        Map<Color, IndexedColors> colorMap = new HashMap<>();
        colorMap.put(Color.BLACK, IndexedColors.BLACK);
        colorMap.put(Color.WHITE, IndexedColors.WHITE);
        colorMap.put(Color.RED, IndexedColors.RED);
        colorMap.put(Color.GREEN, IndexedColors.BRIGHT_GREEN);
        colorMap.put(Color.BLUE, IndexedColors.BLUE);
        colorMap.put(Color.YELLOW, IndexedColors.YELLOW);
        colorMap.put(new Color(128, 0, 128), IndexedColors.VIOLET); // Purple
        colorMap.put(Color.ORANGE, IndexedColors.ORANGE);
        colorMap.put(Color.GRAY, IndexedColors.GREY_50_PERCENT);
        colorMap.put(Color.PINK, IndexedColors.PINK);
        colorMap.put(Color.CYAN, IndexedColors.TURQUOISE);
        colorMap.put(Color.MAGENTA, IndexedColors.ROSE);

        for (Map.Entry<Color, IndexedColors> entry : colorMap.entrySet()) {
            if (entry.getKey().equals(color)) {
                return entry.getValue();
            }
        }

        return IndexedColors.AUTOMATIC;
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(() -> {
            MainAppSwing app = new MainAppSwing();
            app.setVisible(true);
        });
    }
}
