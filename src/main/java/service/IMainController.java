package service;

import dom.gantt.TaskAbstract;
import util.FileTypes;
import util.ProjectInfo;

import java.util.List;

public interface IMainController {

    List<String> load(String sourcePath, FileTypes filetype);

    ProjectInfo prepareTargetWorkbook(FileTypes fileType, String targetPath);

    List<TaskAbstract> getAllTasks();

    List<TaskAbstract> getTopLevelTasksOnly();

    List<TaskAbstract> getTasksInRange(int firstIncluded, int lastIncluded);

    boolean rawWriteToExcelFile(List<TaskAbstract> tasks);

    String addFontedStyle(String styleName, short styleFontColor, short styleFontHeightInPoints, String styleFontName,
                          boolean styleFontBold, boolean styleFontItalic, boolean styleFontStrikeout, short styleFillForegroundColor,
                          String styleFillPatternString, String HorizontalAlignmentString, boolean styleWrapText);

    boolean createNewSheet(String sheetName, List<TaskAbstract> tasks, String headerStyleName, String topBarStyleName,
                           String topDataStyleName, String nonTopBarStyleName, String nonTopDataStyleName, String normalStyleName);

    void createDefaultStyles();
}
