package app.naive;

import app.ApplicationController;
import dom.gantt.TaskAbstract;
import org.apache.poi.ss.usermodel.IndexedColors;
import util.FileTypes;
import util.ProjectInfo;

import java.util.List;

public class NaiveClient {


    public static void main(String[] args) {

        ApplicationController appController = new ApplicationController();


        List<String> loadedStr = appController.load("src/test/resources/input/EggsScrambled.xlsx", FileTypes.XLSX);
        if (loadedStr == null) {
            System.err.println("Failed to load tasks.");
            return;
        }
        System.out.println();
        System.out.println("----------");
        for (String s : loadedStr)
            System.out.println(s);


        ProjectInfo prjInfo = appController.prepareTargetWorkbook(FileTypes.XLSX, "src/test/resources/output/EggsScrambled_Output.xlsx");
        if (prjInfo == null) {
            System.err.println("Failed to prepare target workbook.");
            return;
        }
        System.out.println("----------");
        System.out.println("\n" + prjInfo);
        System.out.println("----------");

        boolean success = appController.rawWriteToExcelFile(appController.getAllTasks());
        if (!success) {
            System.err.println("Failed to write tasks to Excel.");
            return;
        }

        // String styleName,
        // short styleFontColor, short styleFontHeightInPoints, String styleFontName,
        // boolean styleFontBold, boolean styleFontItalic, boolean styleFontStrikeout,
        // short styleFillForegroundColor, String styleFillPatternString, String HorizontalAlignmentString, boolean styleWrapText

        String greenStyleName = appController.addFontedStyle("myTealThing",
                IndexedColors.TEAL.getIndex(), (short) 10, "Times New Roman",
                false, false, false,
                IndexedColors.WHITE.getIndex(), "Solid_FOREGROUND", "Left", false);

        String orangeStyleName = appController.addFontedStyle("myOrangeThing",
                IndexedColors.RED.getIndex(), (short) 10, "Times New Roman",
                false, false, false,
                IndexedColors.ORANGE.getIndex(), "Solid_FOREGROUND", "Left", false);

        appController.createNewSheet("ALL_Styled", appController.getAllTasks(),
                "DefaultHeaderStyle", "TopTask_bar_style", "TopTask_data_style", "NonTopTask_bar_style", "NonTopTask_data_style", "Normal");

        appController.createNewSheet("Top_Level", appController.getTopLevelTasksOnly(),
                "DefaultHeaderStyle", "TopTask_bar_style", "TopTask_data_style", "NonTopTask_bar_style", "NonTopTask_data_style", "Normal");

        appController.createNewSheet("Range", appController.getTasksInRange(103, 202),
                "DefaultHeaderStyle", "TopTask_bar_style", "TopTask_data_style", orangeStyleName, greenStyleName, "Normal");

        System.out.println("End of naive xlsx client");
    }
}
