package app.naive;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


import dom.gantt.TaskAbstract;
import util.FileTypes;
import util.ProjectInfo;

public class MainControllerImpl implements service.IMainController {

    private ProjectInfo projectInfo;
    private Workbook targetWorkbook;
    private String targetPath;
    private Map<String, CellStyle> stylesMap = new HashMap<>();

    @Override
    public List<String> load(String sourcePath, FileTypes filetype) {
        List<String> taskDescriptions = new ArrayList<>();


        try {
            switch (filetype) {
                case XLS:
                case XLSX:
                    taskDescriptions = loadFromExcel(sourcePath);
                    break;
                case CSV:
                case TSV:
                    taskDescriptions = loadFromCSV(sourcePath, filetype);
                    break;
                default:
                    System.err.println("Unsupported file type: " + filetype);
                    return null;
            }


            projectInfo = new ProjectInfo();
            
            return taskDescriptions;

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private List<String> loadFromExcel(String sourcePath) throws IOException {
        List<String> tasks = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(sourcePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            for (Row row : sheet) {
                StringBuilder taskBuilder = new StringBuilder();
                for (Cell cell : row) {
                    taskBuilder.append(cell.toString()).append(" | ");
                }
                tasks.add(taskBuilder.toString());
            }
        }
        return tasks;
    }

    private List<String> loadFromCSV(String sourcePath, FileTypes filetype) throws IOException {
        List<String> tasks = new ArrayList<>();
        String delimiter = filetype == FileTypes.TSV ? "\t" : ",";

        try (BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(sourcePath), "UTF-8"))) {
            String line;
            while ((line = br.readLine()) != null) {
                String[] cells = line.split(delimiter);
                StringBuilder taskBuilder = new StringBuilder();
                for (String cell : cells) {
                    taskBuilder.append(cell).append(" | ");
                }
                tasks.add(taskBuilder.toString());
            }
        }
        return tasks;
    }

    @Override
    public ProjectInfo prepareTargetWorkbook(FileTypes fileType, String targetPath) {
        this.targetPath = targetPath;
        try {
            if (fileType == FileTypes.XLSX) {
                targetWorkbook = new XSSFWorkbook();
            } else if (fileType == FileTypes.XLS) {
                
                targetWorkbook = new XSSFWorkbook();
                System.err.println("XLS format is deprecated. Using XLSX instead.");
            } else {
                System.err.println("Unsupported target file type: " + fileType);
                return null;
            }

  

            ProjectInfo info = new ProjectInfo();
            info.setTargetPath(targetPath);
            info.setFileType(fileType);

            return info;

        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    @Override
    public List<TaskAbstract> getAllTasks() {
        if (projectInfo == null) {
            return null;
        }
        return projectInfo.getAllTasks();
    }

    @Override
    public List<TaskAbstract> getTopLevelTasksOnly() {
        if (projectInfo == null) {
            return null;
        }
        return projectInfo.getTopLevelTasks();
    }

    @Override
    public List<TaskAbstract> getTasksInRange(int firstIncluded, int lastIncluded) {
        if (projectInfo == null) {
            return null;
        }
        return projectInfo.getTasksInRange(firstIncluded, lastIncluded);
    }

    @Override
    public boolean rawWriteToExcelFile(List<TaskAbstract> tasks) {
        if (targetWorkbook == null) {
            System.err.println("Target workbook is not prepared.");
            return false;
        }

        String sheetName = new SimpleDateFormat("dd-MM-yyyy HH_mm_ss").format(new Date());
        Sheet sheet = targetWorkbook.createSheet(sheetName);

        int rowNum = 0;
        for (TaskAbstract task : tasks) {
            Row row = sheet.createRow(rowNum++);
            List<String> taskData = task.toStringList();
            int cellNum = 0;
            for (String data : taskData) {
                Cell cell = row.createCell(cellNum++);
                cell.setCellValue(data);
            }
        }

        try (FileOutputStream fos = new FileOutputStream(targetPath)) {
            targetWorkbook.write(fos);
            return true;
        } catch (IOException e) {
            e.printStackTrace();
            return false;
        }
    }

    @Override
    public String addFontedStyle(String styleName, short styleFontColor, short styleFontHeightInPoints, String styleFontName,
                                 boolean styleFontBold, boolean styleFontItalic, boolean styleFontStrikeout, short styleFillForegroundColor,
                                 String styleFillPatternString, String horizontalAlignmentString, boolean styleWrapText) {
        if (targetWorkbook == null) {
            System.err.println("Target workbook is not prepared.");
            return "Normal";
        }

        try {
            CellStyle style = targetWorkbook.createCellStyle();
            Font font = targetWorkbook.createFont();
            font.setColor(styleFontColor);
            font.setFontHeightInPoints(styleFontHeightInPoints);
            font.setFontName(styleFontName);
            font.setBold(styleFontBold);
            font.setItalic(styleFontItalic);
            font.setStrikeout(styleFontStrikeout);
            style.setFont(font);

            style.setFillForegroundColor(styleFillForegroundColor);
            FillPatternType fillPattern = FillPatternType.valueOf(styleFillPatternString.toUpperCase());
            style.setFillPattern(fillPattern);

            HorizontalAlignment alignment = HorizontalAlignment.valueOf(horizontalAlignmentString.toUpperCase());
            style.setAlignment(alignment);

            style.setWrapText(styleWrapText);

            stylesMap.put(styleName, style);

            return styleName;

        } catch (Exception e) {
            e.printStackTrace();
            return "Normal";
        }
    }

    @Override
    public boolean createNewSheet(String sheetName, List<TaskAbstract> tasks, String headerStyleName, String topBarStyleName,
                                  String topDataStyleName, String nonTopBarStyleName, String nonTopDataStyleName, String normalStyleName) {
        if (targetWorkbook == null) {
            System.err.println("Target workbook is not prepared.");
            return false;
        }

        try {
            Sheet sheet = targetWorkbook.createSheet(sheetName);
            Row headerRow = sheet.createRow(0);
            CellStyle headerStyle = getStyleByName(headerStyleName);
            String[] headers = {"ID", "Name", "Start Date", "End Date", "Duration", "Dependencies"};
            for (int i = 0; i < headers.length; i++) {
                Cell cell = headerRow.createCell(i);
                cell.setCellValue(headers[i]);
                if (headerStyle != null) {
                    cell.setCellStyle(headerStyle);
                }
            }

            int rowNum = 1;
            for (TaskAbstract task : tasks) {
                Row row = sheet.createRow(rowNum++);
                CellStyle barStyle = task.isTopLevel() ? getStyleByName(topBarStyleName) : getStyleByName(nonTopBarStyleName);
                CellStyle dataStyle = task.isTopLevel() ? getStyleByName(topDataStyleName) : getStyleByName(nonTopDataStyleName);

                List<String> taskData = task.toStringList();
                for (int i = 0; i < taskData.size(); i++) {
                    Cell cell = row.createCell(i);
                    cell.setCellValue(taskData.get(i));
                    if (i == 0 || i == 1) { 
                        if (barStyle != null) {
                            cell.setCellStyle(barStyle);
                        }
                    } else { 
                        if (dataStyle != null) {
                            cell.setCellStyle(dataStyle);
                        }
                    }
                }
            }

            try (FileOutputStream fos = new FileOutputStream(targetPath)) {
                targetWorkbook.write(fos);
            }

            return true;

        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }

    @Override
    public void createDefaultStyles() {
    }


    private CellStyle getStyleByName(String styleName) {
        return stylesMap.getOrDefault(styleName, targetWorkbook.createCellStyle());
    }
}
