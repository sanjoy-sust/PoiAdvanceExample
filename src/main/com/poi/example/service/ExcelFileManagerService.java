package main.com.poi.example.service;

import main.com.poi.example.model.Resident;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.util.List;

/**
 * Created by Lenovo on 24/09/2018.
 */
public interface ExcelFileManagerService {
    boolean importExcel(String fileName,List<Resident> residentList) throws IOException;

    void exportExcel(String fileName) throws IOException;

    boolean validateExcelFile(String fileName) throws OpenXML4JException, IOException, ParserConfigurationException, SAXException;
}
