package main.com.poi.example.model;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

/**
 * Created by Lenovo on 16/11/2018.
 */
public class SheetModel {
    public long numberOfRows;
    public long numberOfColumns;
    public long timeToProcess;
    public SXSSFWorkbook workbook;
    public Sheet curOutSheet;
    public String curSheetName;

    public Sheet getCurOutSheet() {
        return curOutSheet;
    }

    public void setCurOutSheet(Sheet curOutSheet) {
        this.curOutSheet = curOutSheet;
    }



    public String getCurSheetName() {
        return curSheetName;
    }

    public void setCurSheetName(String curSheetName) {
        this.curSheetName = curSheetName;
    }



    public SXSSFWorkbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(SXSSFWorkbook workbook) {
        this.workbook = workbook;
    }



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
