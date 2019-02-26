// Java Program to demonstrate adjacency list 
// representation of graphs 

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;

public class TestProblem

{
    public static void main(String[] args) throws IOException {

/*
        InputStream inp = null;
        inp = new FileInputStream("E:\\Projects\\PoiAdvanceExample\\stackProblem.xlsx");

        Workbook wb = WorkbookFactory.create(inp);
        Sheet sheet = wb.getSheetAt(0);

        int rowsCount = sheet.getLastRowNum();
        int columnCount = sheet.getRow(0).getLastCellNum();
        String[][] inputData = new String[rowsCount+1][columnCount];

        for (int i = 0; i <= rowsCount; i++) {
            Row row = sheet.getRow(i);
            int colCounts = row.getLastCellNum();
            for (int j = 0; j < colCounts; j++) {
                Cell cell = row.getCell(j);
                if(cell.getCellType() == CellType.NUMERIC) {
                    inputData[i][j] = Double.toString(cell.getNumericCellValue());
                }
                if(cell.getCellType() == CellType.FORMULA) {
                    inputData[i][j] = cell.getCellFormula();
                }
                if(cell.getCellType() == CellType.STRING) {
                    inputData[i][j] = cell.getStringCellValue();
                }
            }
        }*/

       writeData();

    }

    private static void writeData() throws IOException {

        Workbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = (XSSFSheet) workbook.createSheet();

        int r = 0;
        for (int i=0;i<2;i++) {
            Row row = sheet.createRow(r++);
            int column = 0;
            for (int j =0;j<2;j++) {
                XSSFCell cell = (XSSFCell) row.createCell(column++);
                if (r == 1 || column == 1) cell.setCellValue(i);

                else if (column == 2) {
                    cell.setCellFormula("OFFSET(IU220,0,1)");
                    XSSFFormulaEvaluator evaluator =
                            (XSSFFormulaEvaluator) workbook.getCreationHelper().createFormulaEvaluator();
                    System.out.println(evaluator.evaluateInCell(cell).getCellType());
                }
            }
        }


        FileOutputStream fileOut = new FileOutputStream("stackProblem.xlsx");
        workbook.write(fileOut);
        workbook.close();
    }
}
//XSSFCell cell = sheet.getRow(1).createCell(1); cell.setCellFormula("OFFSET(IV220,0,1)");