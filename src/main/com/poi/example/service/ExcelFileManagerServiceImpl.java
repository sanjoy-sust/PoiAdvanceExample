package main.com.poi.example.service;

import main.com.poi.example.model.Resident;
import main.com.poi.example.util.ParseExcelFile;
import main.com.poi.example.util.ResidentExcelUtil;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

/**
 * Created by Lenovo on 24/09/2018.
 */
public class ExcelFileManagerServiceImpl implements ExcelFileManagerService {
    @Override
    public boolean importExcel(String fileName,List<Resident> residentList) throws IOException {
        ResidentExcelUtil.buildFile(fileName, residentList);
        return true;
    }

    @Override
    public void exportExcel(String fileName) throws IOException {
        List<Resident> residentList = new ArrayList<>();
        for(int i = 0;i<10;i++)
        {
            Resident resident = new Resident();
            resident.setName("Name"+i);
            resident.setMobile("0142485824" + i);
            resident.setAddress("ABC" + i);
            resident.setEmail("count" + i + "@gmail.com");
            resident.setNationalId("8687678687687" + i);
            resident.setAge(i+30);
            residentList.add(resident);
        }
        ResidentExcelUtil.buildFile(fileName, residentList);
    }

    @Override
    public boolean validateExcelFile(String fileName) throws OpenXML4JException, IOException, ParserConfigurationException, SAXException {

        File file = new File(fileName);
        OPCPackage opcPackage = OPCPackage.open(file);
        ParseExcelFile parsedFile = new ParseExcelFile(opcPackage,10,Resident.class);
        boolean isValid = parsedFile.process();
        return isValid;
    }
}
