import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import org.apache.poi.ss.util.*;

class TestProblem {

    public static void main(String[] args) throws Exception{

        Workbook workbook = new XSSFWorkbook();

        //hidden sheet for list values
        Sheet sheet = workbook.createSheet("ListSheet");
        sheet.createRow(0).createCell(0).setCellValue("SourceList");
        int r = 1;
        for (int i = 1; i < 5; i++) {
            sheet.createRow(r++).createCell(0).setCellValue(i);
        }
        //unselect that sheet because we will hide it later
        sheet.setSelected(false);

        //visible data sheet
        sheet = workbook.createSheet("Sheet1");

        //names for the list constraints
        Name namedCell = workbook.createName();
        namedCell.setNameName("List1To4");
        String reference = "ListSheet!$A$2:$A$5"; //List 1 to 4
        namedCell.setRefersToFormula(reference);

   /*     namedCell = workbook.createName();
        namedCell.setNameName("ListLeftCellTo4");
        reference = "INDEX(List1To4,INDEX(Sheet1!$1:$1000,ROW(),COLUMN()-1)):INDEX(List1To4,4)"; //List n to 4
        //List1To4Position=ThisRow.ThisColumn-1 : List1To4LastPosition
        namedCell.setRefersToFormula(reference);*/

        sheet.createRow(0).createCell(0).setCellValue("1 to 4");

        sheet.setActiveCell(new CellAddress("A2"));

        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);

        //data validations

        DataValidationHelper dvHelper = sheet.getDataValidationHelper();
        DataValidationConstraint dvConstraint = dvHelper.createFormulaListConstraint("List1To4");
        CellRangeAddressList addressList = new CellRangeAddressList(1, 4, 0, 0);
        DataValidation validation = dvHelper.createValidation(dvConstraint, addressList);
        sheet.addValidationData(validation);

/*        dvConstraint = dvHelper.createFormulaListConstraint("ListLeftCellTo4");
        addressList = new CellRangeAddressList(1, 1, 1, 1);
        validation = dvHelper.createValidation(dvConstraint, addressList);*/

        sheet.addValidationData(validation);

        //hide the ListSheet
        workbook.setSheetHidden(0, true);
        //set Sheet1 active
        workbook.setActiveSheet(1);

        FileOutputStream out = new FileOutputStream("CreateExcelDependentDataValidationLists.xlsx");
        workbook.write(out);
        workbook.close();
        out.close();

    }
}