package app;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.InputStreamReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.Date;
import java.util.List;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import dom.gantt.TaskAbstract;
import dom.gantt.TaskConcrete;
import service.IMainController;
import util.FileTypes;
import util.ProjectInfo;

public class ApplicationController implements IMainController {

    private ProjectInfo projectInfo;
    private Workbook targetWorkbook;
    private String targetPath;
    private Map<String, CellStyle> stylesMap = new HashMap<>();

    @Override
    public List<String> load(String sourcePath, FileTypes filetype) {
        List<String> taskDescriptions = new ArrayList<>();
        try {
            List<TaskAbstract> tasks;
            switch (filetype) {
                case XLS:
                case XLSX:
                    tasks = loadAndParseFromExcel(sourcePath);
                    break;
                case CSV:
                case TSV:
                    tasks = loadAndParseFromCSV(sourcePath, filetype);
                    break;
                default:
                    System.err.println("Unsupported file type: " + filetype);
                    return null;
            }

            tasks.sort(Comparator.comparingInt(TaskAbstract::getId));
            projectInfo = new ProjectInfo();
            projectInfo.setTasks(tasks);

            for (TaskAbstract t : tasks) {
                taskDescriptions.add(t.toString());
            }

            return taskDescriptions;
        } catch (Exception e) {
            e.printStackTrace();
            return null;
        }
    }

    private List<TaskAbstract> loadAndParseFromExcel(String sourcePath) throws IOException {
        List<TaskAbstract> tasks = new ArrayList<>();
        try (FileInputStream fis = new FileInputStream(sourcePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);
            boolean header = true;
            for (Row row : sheet) {
                if (header) {
                    header = false;
                    continue;
                }
                if (row == null) continue;
                int id = (int) getNumericCellValue(row.getCell(0));
                String name = getStringCellValue(row.getCell(1));
                int containerId = (int) getNumericCellValue(row.getCell(2));
                Integer startDay = null;
                Integer endDay = null;
                Double cost = 0.0;
                Double effort = 0.0;

                if (containerId != 0) {
                    startDay = (int) getNumericCellValue(row.getCell(3));
                    endDay = (int) getNumericCellValue(row.getCell(4));
                    cost = getNumericCellValue(row.getCell(5));
                    effort = getNumericCellValue(row.getCell(6));
                }

                TaskAbstract task = new TaskConcrete(id, name, containerId, startDay, endDay, cost, effort);
                tasks.add(task);
            }
        }
        return tasks;
    }

    private List<TaskAbstract> loadAndParseFromCSV(String sourcePath, FileTypes filetype) throws IOException {
        List<TaskAbstract> tasks = new ArrayList<>();
        String delimiter = filetype == FileTypes.TSV ? "\t" : ",";
        try (BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(sourcePath), "UTF-8"))) {
            String line;
            boolean header = true;
            while ((line = br.readLine()) != null) {
                if (header) {
                    header = false;
                    continue;
                }
                String[] parts = line.split(delimiter, -1);
                if (parts.length < 7) {
                    continue;
                }
                try {
                    int id = parts[0].trim().isEmpty() ? 0 : Integer.parseInt(parts[0].trim());
                    String name = parts[1].trim();
                    int containerId = parts[2].trim().isEmpty() ? 0 : Integer.parseInt(parts[2].trim());
                    Integer startDay = null;
                    Integer endDay = null;
                    Double cost = 0.0;
                    Double effort = 0.0;

                    if (containerId != 0) {
                        startDay = parts[3].trim().isEmpty() ? null : Integer.parseInt(parts[3].trim());
                        endDay = parts[4].trim().isEmpty() ? null : Integer.parseInt(parts[4].trim());
                        cost = parts[5].trim().isEmpty() ? 0.0 : Double.parseDouble(parts[5].trim());
                        effort = parts[6].trim().isEmpty() ? 0.0 : Double.parseDouble(parts[6].trim());
                    }

                    TaskAbstract task = new TaskConcrete(id, name, containerId, startDay, endDay, cost, effort);
                    tasks.add(task);
                } catch (NumberFormatException e) {
                    
                }
            }
        }
        return tasks;
    }

    private double getNumericCellValue(Cell cell) {
        if (cell == null) return 0;
        if (cell.getCellType() == CellType.NUMERIC) {
            return cell.getNumericCellValue();
        } else if (cell.getCellType() == CellType.STRING) {
            try {
                return Double.parseDouble(cell.getStringCellValue().trim());
            } catch (NumberFormatException e) {
                return 0;
            }
        } else {
            return 0;
        }
    }

    private String getStringCellValue(Cell cell) {
        if (cell == null) return "";
        if (cell.getCellType() == CellType.STRING) {
            return cell.getStringCellValue().trim();
        } else if (cell.getCellType() == CellType.NUMERIC) {
            if (DateUtil.isCellDateFormatted(cell)) {
                SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
                return dateFormat.format(cell.getDateCellValue());
            } else {
                return String.valueOf((int) cell.getNumericCellValue());
            }
        } else if (cell.getCellType() == CellType.BOOLEAN) {
            return String.valueOf(cell.getBooleanCellValue());
        } else {
            return "";
        }
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

            FillPatternType fillPattern = FillPatternType.valueOf(styleFillPatternString.toUpperCase());
            style.setFillForegroundColor(styleFillForegroundColor);
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
            if (tasks == null || tasks.isEmpty()) {
                return false;
            }

            tasks.sort((t1, t2) -> Integer.compare(t1.getId(), t2.getId()));
            List<String[]> intermediateData = createIntermediateRepresentation(tasks);

            Sheet sheet = targetWorkbook.createSheet(sheetName);
            writeSheetWithX(sheet, intermediateData, tasks, headerStyleName, topBarStyleName, topDataStyleName,
                    nonTopBarStyleName, nonTopDataStyleName, normalStyleName);

            Sheet colorSheet = targetWorkbook.createSheet(sheetName + "_Colorized");
            writeSheetWithColorOnly(colorSheet, intermediateData, tasks);

            try (FileOutputStream fos = new FileOutputStream(targetPath)) {
                targetWorkbook.write(fos);
            }

            return true;

        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }

    private void writeSheetWithX(Sheet sheet, List<String[]> intermediateData, List<TaskAbstract> tasks,
                                 String headerStyleName, String topBarStyleName, String topDataStyleName,
                                 String nonTopBarStyleName, String nonTopDataStyleName, String normalStyleName) {

        CellStyle headerStyle = getStyleByName(headerStyleName);
        CellStyle topBar = getStyleByName(topBarStyleName);
        CellStyle topData = getStyleByName(topDataStyleName);
        CellStyle nonTopBar = getStyleByName(nonTopBarStyleName);
        CellStyle nonTopData = getStyleByName(nonTopDataStyleName);
        CellStyle normal = getStyleByName(normalStyleName);

        for (int rowNum = 0; rowNum < intermediateData.size(); rowNum++) {
            Row row = sheet.createRow(rowNum);
            String[] rowData = intermediateData.get(rowNum);

            for (int cellNum = 0; cellNum < rowData.length; cellNum++) {
                Cell cell = row.createCell(cellNum);
                String cellValue = rowData[cellNum];

                if (rowNum == 0) {
                    cell.setCellValue(cellValue);
                    if (headerStyle != null) cell.setCellStyle(headerStyle);
                } else {
                    TaskAbstract task = tasks.get(rowNum - 1);
                    boolean isTop = task.isTopLevel();
                    if (cellNum < 5) { 
                        cell.setCellValue(cellValue);
                        if (cellNum == 0) {
                            cell.setCellStyle(isTop ? topData : nonTopData);
                        } else if (cellNum < 4) {
                            cell.setCellStyle(isTop ? topData : nonTopData);
                        } else {
                            cell.setCellStyle(isTop ? topData : nonTopData);
                        }
                    } else {
                        cell.setCellValue(cellValue);
                        if (cellValue.equalsIgnoreCase("x")) {
                            cell.setCellStyle(isTop ? topBar : nonTopBar);
                        } else {
                            cell.setCellStyle(normal);
                        }
                    }
                }
            }
        }

        for (int i = 0; i < intermediateData.get(0).length; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void writeSheetWithColorOnly(Sheet sheet, List<String[]> intermediateData, List<TaskAbstract> tasks) {
        for (int rowNum = 0; rowNum < intermediateData.size(); rowNum++) {
            Row row = sheet.createRow(rowNum);
            String[] rowData = intermediateData.get(rowNum);
            for (int cellNum = 0; cellNum < rowData.length; cellNum++) {
                Cell cell = row.createCell(cellNum);
                String cellValue = rowData[cellNum];

                if (rowNum == 0) {
                    cell.setCellValue(cellValue);
                } else {
                    TaskAbstract task = tasks.get(rowNum - 1);
                    boolean isTop = task.isTopLevel();
                    if (cellNum < 5) {
                        cell.setCellValue(cellValue);
                    } else {
                        if ("x".equalsIgnoreCase(cellValue)) {
                            cell.setCellValue("");
                            if (isTop) {
                                changeCellBackgroundColor(cell, IndexedColors.LIGHT_BLUE.getIndex());
                            } else {
                                changeCellBackgroundColor(cell, IndexedColors.GREY_40_PERCENT.getIndex());
                            }
                        } else {
                            cell.setCellValue("");
                        }
                    }
                }
            }
        }

        for (int i = 0; i < intermediateData.get(0).length; i++) {
            sheet.autoSizeColumn(i);
        }
    }

    private void changeCellBackgroundColor(Cell cell, short color) {
        CellStyle cellStyle = cell.getCellStyle();
        if (cellStyle == null) {
            cellStyle = cell.getSheet().getWorkbook().createCellStyle();
        } else {
            CellStyle newStyle = cell.getSheet().getWorkbook().createCellStyle();
            newStyle.cloneStyleFrom(cellStyle);
            cellStyle = newStyle;
        }
        cellStyle.setFillForegroundColor(color);
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(cellStyle);
    }

    private CellStyle getStyleByName(String styleName) {
        return stylesMap.getOrDefault(styleName, null);
    }

    public List<String[]> createIntermediateRepresentation(List<TaskAbstract> tasks) {
        Integer earliestStartDay = null;
        Integer latestEndDay = null;

        for (TaskAbstract task : tasks) {
            Integer startDay = task.getStartDay();
            Integer endDay = task.getEndDay();
            if (startDay != null && (earliestStartDay == null || startDay < earliestStartDay)) {
                earliestStartDay = startDay;
            }
            if (endDay != null && (latestEndDay == null || endDay > latestEndDay)) {
                latestEndDay = endDay;
            }
        }

        if (earliestStartDay == null) earliestStartDay = 0;
        if (latestEndDay == null) latestEndDay = 0;

        List<String[]> intermediateData = new ArrayList<>();
        List<String> headerRow = new ArrayList<>();
        headerRow.add("Level");
        headerRow.add("ID");
        headerRow.add("Description");
        headerRow.add("Cost");
        headerRow.add("Effort");

        List<Integer> dayList = new ArrayList<>();
        for (int day = earliestStartDay; day <= latestEndDay; day++) {
            headerRow.add(String.valueOf(day));
            dayList.add(day);
        }

        intermediateData.add(headerRow.toArray(new String[0]));

        for (TaskAbstract task : tasks) {
            List<String> row = new ArrayList<>();
            row.add(task.isTopLevel() ? "top" : "");
            row.add(String.valueOf(task.getId()));
            row.add(task.getName());
            row.add(String.valueOf(task.getCost()));
            row.add(String.valueOf(task.getEffort()));

            for (int i = 0; i < dayList.size(); i++) {
                row.add("");
            }

            Integer taskStartDay = task.getStartDay();
            Integer taskEndDay = task.getEndDay();

            if (taskStartDay != null && taskEndDay != null) {
                for (int i = 0; i < dayList.size(); i++) {
                    int currentDay = dayList.get(i);
                    if (currentDay >= taskStartDay && currentDay <= taskEndDay) {
                        row.set(5 + i, "x");
                    }
                }
            }

            intermediateData.add(row.toArray(new String[0]));
        }

        return intermediateData;
    }

    @Override
    public void createDefaultStyles() {
        addFontedStyle("Normal", IndexedColors.BLACK.getIndex(), (short)11, "Calibri",
                false, false, false, IndexedColors.WHITE.getIndex(), "NO_FILL", "LEFT", false);

        addFontedStyle("DefaultHeaderStyle", IndexedColors.WHITE.getIndex(), (short)12, "Arial",
                true, false, false, IndexedColors.GREY_80_PERCENT.getIndex(), "SOLID_FOREGROUND", "CENTER", true);

        addFontedStyle("TopTask_bar_style", IndexedColors.BLACK.getIndex(), (short)11, "Calibri",
                true, false, false, IndexedColors.BLUE.getIndex(), "SOLID_FOREGROUND", "CENTER", false);

        addFontedStyle("TopTask_data_style", IndexedColors.BLUE.getIndex(), (short)11, "Calibri",
                true, false, false, IndexedColors.WHITE.getIndex(), "NO_FILL", "LEFT", false);

        addFontedStyle("NonTopTask_bar_style", IndexedColors.BLACK.getIndex(), (short)11, "Calibri",
                false, false, false, IndexedColors.GREY_50_PERCENT.getIndex(), "SOLID_FOREGROUND", "CENTER", false);

        addFontedStyle("NonTopTask_data_style", IndexedColors.BLACK.getIndex(), (short)11, "Calibri",
                false, false, false, IndexedColors.WHITE.getIndex(), "NO_FILL", "LEFT", false);
    }
}
