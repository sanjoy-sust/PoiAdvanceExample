package main.com.poi.example.util;

import main.com.poi.example.model.SheetModel;
import main.com.poi.example.model.XSSFDataTypes;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import java.io.PrintStream;
import java.util.ArrayList;
import java.util.List;

class XSSFSheetEventHandler extends DefaultHandler {

    /**
     * Table with styles
     */
    private StylesTable stylesTable;

    /**
     * Table with unique strings
     */
    private ReadOnlySharedStringsTable sharedStringsTable;

    private List<?> list = new ArrayList();

    private SheetModel clazz;

    /**
     * Number of columns to read starting with leftmost
     */
    private final int minColumnCount;

    // Set when V start element is seen
    private boolean vIsOpen;

    // Set when cell start element is seen;
    // used when cell close element is seen.
    private XSSFDataTypes nextDataType;

    // Used to format numeric cell values.
    private short formatIndex;
    private String formatString;
    private final DataFormatter formatter;
    // The last column printed to the output stream
    private int lastColumnNumber = -1;

    // Gathers characters as they are seen.
    private StringBuffer value;

    /**
     * Accepts objects needed while parsing.
     *
     * @param styles
     *            Table of styles
     * @param strings
     *            Table of shared strings
     * @param cols
     *            Minimum number of columns to show
     */
    public XSSFSheetEventHandler(StylesTable styles,
                            ReadOnlySharedStringsTable strings, int cols, SheetModel clazz) {
        this.stylesTable = styles;
        this.sharedStringsTable = strings;
        this.minColumnCount = cols;
        this.value = new StringBuffer();
        this.nextDataType = XSSFDataTypes.NUMBER;
        this.formatter = new DataFormatter();
        this.clazz = clazz;
    }

    public void startElement(String uri, String localName, String name,
                             Attributes attributes) throws SAXException {

        if ("inlineStr".equals(name) || "v".equals(name)) {
            vIsOpen = true;
            // Clear contents cache
            value.setLength(0);
        }
        // c => cell
        else if ("c".equals(name)) {
            // Get the cell reference
            String r = attributes.getValue("r");

            // Set up defaults.
            this.nextDataType = XSSFDataTypes.NUMBER;
            this.formatIndex = -1;
            this.formatString = null;
            String cellType = attributes.getValue("t");
            String cellStyleStr = attributes.getValue("s");
            if ("b".equals(cellType))
                nextDataType = XSSFDataTypes.BOOL;
            else if ("e".equals(cellType))
                nextDataType = XSSFDataTypes.ERROR;
            else if ("inlineStr".equals(cellType))
                nextDataType = XSSFDataTypes.INLINESTR;
            else if ("s".equals(cellType))
                nextDataType = XSSFDataTypes.SSTINDEX;
            else if ("str".equals(cellType))
                nextDataType = XSSFDataTypes.FORMULA;
            else if (cellStyleStr != null) {
                // It's a number, but almost certainly one
                // with a special style or format
                int styleIndex = Integer.parseInt(cellStyleStr);
                XSSFCellStyle style = stylesTable.getStyleAt(styleIndex);
                this.formatIndex = style.getDataFormat();
                this.formatString = style.getDataFormatString();
                if (this.formatString == null)
                    this.formatString = BuiltinFormats
                            .getBuiltinFormat(this.formatIndex);
            }
        }

    }

    public void endElement(String uri, String localName, String name)
            throws SAXException {

        String thisStr = null;

        // v => contents of a cell
        if ("v".equals(name)) {
            // Process the value contents as required.
            // Do now, as characters() may be called more than once
            switch (nextDataType) {

                case BOOL:
                    char first = value.charAt(0);
                    thisStr = first == '0' ? "FALSE" : "TRUE";
                    break;

                case ERROR:
                    thisStr = "\"ERROR:" + value.toString() + '"';
                    break;

                case FORMULA:
                    // A formula could result in a string value,
                    // so always add double-quote characters.
                    thisStr = '"' + value.toString() + '"';
                    break;

                case INLINESTR:
                    // TODO: have seen an example of this, so it's untested.
                    XSSFRichTextString rtsi = new XSSFRichTextString(value
                            .toString());
                    thisStr = '"' + rtsi.toString() + '"';
                    break;

                case SSTINDEX:
                    String sstIndex = value.toString();
                    try {
                        int idx = Integer.parseInt(sstIndex);
                        XSSFRichTextString rtss = new XSSFRichTextString(
                                sharedStringsTable.getEntryAt(idx));
                        thisStr = '"' + rtss.toString() + '"';
                    } catch (NumberFormatException ex) {

                    }
                    break;

                case NUMBER:
                    String n = value.toString();
                    if (this.formatString != null)
                        thisStr = formatter.formatRawCellContents(Double
                                        .parseDouble(n), this.formatIndex,
                                this.formatString);
                    else
                        thisStr = n;
                    break;

                default:
                    thisStr = "(TODO: Unexpected type: " + nextDataType + ")";
                    break;
            }

            // Output after we've seen the string contents
            // Emit commas for any fields that were missing on this row
            if (lastColumnNumber == -1) {
                lastColumnNumber = 0;
            }

        } else if ("row".equals(name)) {
            clazz.setNumberOfRows(clazz.getNumberOfRows()+1);
        }

    }

    /**
     * Captures characters only if a suitable element is open. Originally
     * was just "v"; extended for inlineStr also.
     */
    public void characters(char[] ch, int start, int length)
            throws SAXException {
        if (vIsOpen)
            value.append(ch, start, length);
    }
}
