package dom.gantt;

import java.util.ArrayList;
import java.util.List;

public class TaskConcrete extends TaskAbstract {

    public TaskConcrete(int id, String name, int containerId, Integer startDay, Integer endDay, double cost, double effort) {
        super(id, name, containerId, startDay, endDay, cost, effort);
    }

    @Override
    public String toString() {
        return "TaskConcrete{" +
                "id=" + getId() +
                ", name='" + getName() + '\'' +
                ", containerId=" + getContainerId() +
                ", startDay=" + getStartDay() +
                ", endDay=" + getEndDay() +
                ", cost=" + cost +
                ", effort=" + effort +
                '}';
    }

    @Override
    public List<String> toStringList() {
        List<String> list = new ArrayList<>();
        list.add(String.valueOf(getId()));
        list.add(getName());
        list.add(String.valueOf(getContainerId()));
        list.add(getStartDay() != null ? String.valueOf(getStartDay()) : "");
        list.add(getEndDay() != null ? String.valueOf(getEndDay()) : "");
        list.add(String.valueOf(cost));
        list.add(String.valueOf(effort));
        return list;
    }
}
