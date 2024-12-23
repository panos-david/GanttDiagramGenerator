package dom.gantt;

import java.util.List;

public abstract class TaskAbstract {

    protected int id;
    protected String name;
    protected int containerId; 
    protected Integer startDay; 
    protected Integer endDay;
    protected double cost;
    protected double effort;


    public TaskAbstract(int id, String name, int containerId, Integer startDay, Integer endDay, double cost, double effort) {
        this.id = id;
        this.name = name;
        this.containerId = containerId;
        this.startDay = startDay;
        this.endDay = endDay;
        this.cost = cost;
        this.effort = effort;
    }


    public int getId() { return id; }
    public String getName() { return name; }
    public int getContainerId() { return containerId; }
    public Integer getStartDay() { return startDay; }
    public Integer getEndDay() { return endDay; }
    public double getCost() { return cost; }
    public double getEffort() { return effort; }
    public boolean isTopLevel() { return containerId == 0; }

    public abstract List<String> toStringList();
}
