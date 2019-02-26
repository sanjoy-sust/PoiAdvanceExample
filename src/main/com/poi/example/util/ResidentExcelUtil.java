package main.com.poi.example.util;

import main.com.poi.example.model.Resident;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Lenovo on 23/10/2018.
 */
public class ResidentExcelUtil {


    public static Boolean buildFile(String fileName,List<Resident> residentList) throws IOException {

        List<String> residentHeaders = new ArrayList<>();
        residentHeaders.add(ExcelFileHeaderConstants.NAME);
        residentHeaders.add(ExcelFileHeaderConstants.ADDRESS);
        residentHeaders.add(ExcelFileHeaderConstants.MOBILE);
        residentHeaders.add(ExcelFileHeaderConstants.EMAIL);
        residentHeaders.add(ExcelFileHeaderConstants.AGE);
        residentHeaders.add(ExcelFileHeaderConstants.NATIONAL_ID);

        Workbook workbook = buildWorkbook(residentHeaders,"TEST_SHEET",residentList);
        FileOutputStream outputStream = new FileOutputStream(fileName);
        workbook.write(outputStream);
        outputStream.close();
        workbook.close();
        return true;
    }

    private static Workbook buildWorkbook(List<String> headers, String sheetName, List<Resident> data) {
        //1. Workbook creation
        Workbook workbook = new XSSFWorkbook();
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        font.setColor((short) Font.COLOR_NORMAL);
        XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
        cellStyle.setFont(font);

        //2. Sheet creation
        Sheet sheet = workbook.createSheet();
        sheet.setColumnWidth((short) 0, (short) ((50 * 8) / ((double) 1 / 20)));
        sheet.setColumnWidth((short) 1, (short) ((50 * 8) / ((double) 1 / 20)));
        workbook.setSheetName(0,sheetName);

        //Header creation
        int count=0;
        Row headerRow = sheet.createRow(count);
        for (String header : headers) {
            Cell cell1 = headerRow.createCell(count++);
            cell1.setCellValue(header);
            cell1.setCellStyle(cellStyle);
        }

        int rownum = 1;
        Row row = null;
        Cell cell = null;
        count=0;
        for (Resident resident:data) {
            count=0;
            row = sheet.createRow(rownum++);
            cell = row.createCell(count++);
            cell.setCellValue(resident.getName());
            cell = row.createCell(count++);
            cell.setCellValue(resident.getAddress());
            cell = row.createCell(count++);
            cell.setCellValue(resident.getMobile());
            cell = row.createCell(count++);
            cell.setCellValue(resident.getEmail());
            cell = row.createCell(count++);
            cell.setCellValue(resident.getAge());
            cell = row.createCell(count++);
            cell.setCellValue(resident.getNationalId());
        }

        return workbook;
    }
}
