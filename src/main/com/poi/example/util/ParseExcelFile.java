package main.com.poi.example.util;

import main.com.poi.example.model.SheetModel;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;
import java.io.PrintStream;

/**
 * Created by Lenovo on 14/11/2018.
 */
public class ParseExcelFile {

    private OPCPackage xlsxPackage;
    private int minColumns;
    private PrintStream output;
    private Class clazz;
    private SXSSFWorkbook outWorkbook;

    /**
     *
     * @param pkg
     *            The XLSX package to process
     * @param minColumns
     *            The minimum number of columns to output, or -1 for no minimum
     */
    public ParseExcelFile(OPCPackage pkg, int minColumns, Class clazz) {
        this.xlsxPackage = pkg;
        this.minColumns = minColumns;
        this.clazz = clazz;
        this.outWorkbook = new SXSSFWorkbook();
    }
    /**
     * Parses and shows the content of one sheet using the specified styles and
     * shared-strings tables.
     *
     * @param styles
     * @param strings
     * @param sheetInputStream
     */
    public void processSheet(StylesTable styles,
                             ReadOnlySharedStringsTable strings,String sheetName, InputStream sheetInputStream)
            throws IOException, ParserConfigurationException, SAXException {
        SheetModel sheetModel = new SheetModel();
        sheetModel.setNumberOfColumns(0);
        sheetModel.setNumberOfRows(0);
        sheetModel.setWorkbook(outWorkbook);
        sheetModel.setCurSheetName(sheetName);
        sheetModel.setCurOutSheet(outWorkbook.createSheet(sheetName));
        long startTime = System.currentTimeMillis();
        InputSource sheetSource = new InputSource(sheetInputStream);
        SAXParserFactory saxFactory = SAXParserFactory.newInstance();
        SAXParser saxParser = saxFactory.newSAXParser();
        XMLReader sheetParser = saxParser.getXMLReader();
        ContentHandler handler = new XSSFSheetEventHandler(styles, strings,
                this.minColumns,  sheetModel);
        sheetParser.setContentHandler(handler);
        sheetParser.parse(sheetSource);
        sheetModel.setTimeToProcess(System.currentTimeMillis() - startTime);
        System.out.println("Number Of Rows : "+ sheetModel.getNumberOfRows());
        System.out.println("Time to process : "+ sheetModel.getTimeToProcess());
    }

    /**
     * Initiates the processing of the XLS workbook file to CSV.
     *
     * @throws IOException
     * @throws OpenXML4JException
     * @throws ParserConfigurationException
     * @throws SAXException
     */
    public boolean process() throws IOException, OpenXML4JException,
            ParserConfigurationException, SAXException {

        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(
                this.xlsxPackage);
        XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);

        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader
                .getSheetsData();
        while (iter.hasNext()) {
            InputStream stream = iter.next();
            String sheetName = iter.getSheetName();
            System.out.println("Processing Sheet : "+sheetName);
            processSheet(styles, strings,sheetName, stream);
            stream.close();
        }
        return true;
    }
}
