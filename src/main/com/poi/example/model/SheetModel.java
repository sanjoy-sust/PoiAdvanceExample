package main.com.poi.example.model;

/**
 * Created by Lenovo on 16/11/2018.
 */
public class SheetModel {
    public long numberOfRows;
    public long numberOfColumns;
    public long timeToProcess;

    public long getNumberOfRows() {
        return numberOfRows;
    }

    public void setNumberOfRows(long numberOfRows) {
        this.numberOfRows = numberOfRows;
    }

    public long getNumberOfColumns() {
        return numberOfColumns;
    }

    public void setNumberOfColumns(long numberOfColumns) {
        this.numberOfColumns = numberOfColumns;
    }

    public long getTimeToProcess() {
        return timeToProcess;
    }

    public void setTimeToProcess(long timeToProcess) {
        this.timeToProcess = timeToProcess;
    }
}
