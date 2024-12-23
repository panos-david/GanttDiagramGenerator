package util;

import dom.gantt.TaskAbstract;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;
import java.util.stream.Collectors;

public class ProjectInfo {
    private String targetPath;
    private FileTypes fileType;
    private List<TaskAbstract> tasks;

    public ProjectInfo() {
        this.tasks = new ArrayList<>();
    }

    public void setTargetPath(String targetPath) {
        this.targetPath = targetPath;
    }

    public void setFileType(FileTypes fileType) {
        this.fileType = fileType;
    }

    public String getTargetPath() {
        return targetPath;
    }

    public FileTypes getFileType() {
        return fileType;
    }

    public void setTasks(List<TaskAbstract> tasks) {
        this.tasks = tasks;
    }

    public List<TaskAbstract> getAllTasks() {
        return tasks;
    }

    public List<TaskAbstract> getTopLevelTasks() {
        return tasks.stream()
                .filter(TaskAbstract::isTopLevel)
                .collect(Collectors.toList());
    }

    public List<TaskAbstract> getTasksInRange(int firstIncluded, int lastIncluded) {
        return tasks.stream()
                .filter(task -> task.getId() >= firstIncluded && task.getId() <= lastIncluded)
                .collect(Collectors.toList());
    }

    public void sortTasksById() {
        tasks.sort(Comparator.comparingInt(TaskAbstract::getId));
    }

    public void sortTasksByName() {
        tasks.sort(Comparator.comparing(TaskAbstract::getName));
    }
}
