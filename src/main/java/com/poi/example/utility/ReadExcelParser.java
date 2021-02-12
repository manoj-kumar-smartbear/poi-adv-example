package com.poi.example.utility;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.util.SAXHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;

public class ReadExcelParser {
    private List<Row> excelData;
    private int minColumns;
    private String excelFile;

    /**
     *
     * @param fileName
     *            The XLSX file to process
     * @param minColumns
     *            The minimum number of columns to output, or -1 for no minimum
     */
    public ReadExcelParser(String fileName, int minColumns) {
        this.excelFile = fileName;
        this.minColumns = minColumns;
    }

    /**
     * Initiates the processing of the XLS workbook file to CSV.
     *
     * @throws IOException
     * @throws OpenXML4JException
     * @throws ParserConfigurationException
     * @throws SAXException
     */
    public List<Row> process() throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {
        File xlsxFile = new File(excelFile);
        if (!xlsxFile.exists()) {
            throw new FileNotFoundException("Not found or not a file: " + xlsxFile.getPath());
        }
        OPCPackage opcPackage = OPCPackage.open(xlsxFile);
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(opcPackage);
        XSSFReader xssfReader = new XSSFReader(opcPackage);

        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader .getSheetsData();
        while (iter.hasNext()) {
            excelData = new ArrayList<>();

            InputStream stream = iter.next();
            String sheetName = iter.getSheetName();
            System.out.println("Processing Sheet : " + sheetName);

            // Now process sheet
            this.processSheet(styles, strings, sheetName, stream);

            stream.close();
        }
        return excelData;
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
                             ReadOnlySharedStringsTable strings, String sheetName, InputStream sheetInputStream)
            throws IOException, ParserConfigurationException, SAXException {
        InputSource sheetSource = new InputSource(sheetInputStream);
        XMLReader sheetParser = SAXHelper.newXMLReader();
        ContentHandler handler = new XSSFSheetXMLHandler(styles, strings, new XSSFSheetContentHandler(),false);
        sheetParser.setContentHandler(handler);
        sheetParser.parse(sheetSource);
    }

    private class XSSFSheetContentHandler implements SheetContentsHandler {
        Row rowData;
        private int currentRow = -1;
        private int currentCol = -1;

        @Override
        public void startRow(int rowNum) {
            rowData = new Row();
            currentRow = rowNum;
            currentCol = -1;
        }
        @Override
        public void endRow(int rowNum) {
            for (int i = currentCol; i < minColumns; i++) {
                rowData.addCell(getBlankCell());
            }
            excelData.add(rowData);
        }
        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            if (cellReference == null) {
                cellReference = new CellAddress(currentRow, currentCol).formatAsString();
            }

            int thisCol = (new CellReference(cellReference)).getCol();
            int missedCols = thisCol - currentCol - 1;
            for (int i = 0; i < missedCols; i++) {
                rowData.addCell(getBlankCell());
            }
            currentCol = thisCol;

            if (minColumns < currentCol) {
                minColumns = currentCol;
            }
            rowData.addCell(getCell(formattedValue));

        }
        @Override
        public void headerFooter(String text, boolean isHeader, String tagName) {
            //codes for reading header/footer if required
        }
        private Cell getCell(String formattedValue) {
            Cell cell;
            if (isNumeric(formattedValue)) {
                cell = new Cell(CellType.NUMERIC, Double.parseDouble(formattedValue));
            } else if (isBoolean(formattedValue)) {
                cell = new Cell(CellType.BOOLEAN, Boolean.parseBoolean(formattedValue));
            } else {
                cell = new Cell(CellType.STRING, formattedValue);
            }
            return cell;
        }
        private boolean isNumeric(String value) {
            try {
                Double.parseDouble(value);
                return true;
            } catch (NumberFormatException e) {
                return false;
            }
        }
        private boolean isBoolean(String value) {
            return value != null
                    && Arrays.stream(new String[] { "true", "false" }).anyMatch(b -> b.equalsIgnoreCase(value));
        }
        private Cell getBlankCell() {
            return new Cell(CellType.BLANK, null);
        }
    }

    public class Row {
        private List<Cell> cells;

        public Row() {
            this.cells = new ArrayList<>();
        }
        public List<Cell> getCells() {
            return cells;
        }
        public void addCell(Cell cell) {
            cells.add(cell);
        }
    }

    class Cell {
        private CellType type;
        private Object data;

        public Cell(CellType type, Object data) {
            super();
            this.type = type;
            this.data = data;
        }
        public CellType getType() {
            return type;
        }
        public String getStringCellValue() {
            return (String) data;
        }
        public double getNumericCellValue() {
            return (double) data;
        }
        public boolean getBooleanCellValue() {
            return (boolean) data;
        }
    }
}