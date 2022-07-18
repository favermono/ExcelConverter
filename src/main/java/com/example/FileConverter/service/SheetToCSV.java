package com.example.FileConverter.service;

import com.example.FileConverter.exceptions.ParserException;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.io.IOException;
import java.io.OutputStreamWriter;
import java.io.PrintStream;
import java.util.ArrayList;

public class SheetToCSV implements XSSFSheetXMLHandler.SheetContentsHandler {
    private ExcelSheetReadConfig readConfig;
    CSVFormat csvFormat;

    private boolean firstCellOfRow;
    private boolean skipRow;
    private int currentRow = -1;
    private int currentCol = -1;
    private int rowCount = 0;
    private boolean rowHasValues = false;
    private int skippedColumns = 0;

    private CSVPrinter printer;

    private boolean firstRow = false;

    private ArrayList<Object> fieldValues;

    public int getRowCount() {
        return rowCount;
    }

    public void setOutput(PrintStream output) {
        final OutputStreamWriter streamWriter = new OutputStreamWriter(output);

        try {
            printer = new CSVPrinter(streamWriter, csvFormat);
        } catch (IOException e) {
            throw new ParserException("Failed to create CSV Printer for file.");
        }
    }

    public SheetToCSV(ExcelSheetReadConfig readConfig, CSVFormat csvFormat) {
        this.readConfig = readConfig;
        this.csvFormat = csvFormat;
    }

    @Override
    public void startRow(int rowNum) {
        if (rowNum <= readConfig.getOverrideFirstRow()) {
            skipRow = true;
            return;
        }

        // Prepare for this row
        skipRow = false;
        firstCellOfRow = true;
        firstRow = currentRow == -1;
        currentRow = rowNum;
        currentCol = -1;
        rowHasValues = false;

        fieldValues = new ArrayList<>();
    }

    @Override
    public void endRow(int rowNum) {
        if (skipRow) {
            return;
        }

        if (firstRow) {
            readConfig.setLastColumn(currentCol);
        }

        //if there was no data in this row, don't write it
        if (!rowHasValues) {
            return;
        }

        // Ensure the correct number of columns
        int columnsToAdd = (readConfig.getLastColumn() - currentCol) - readConfig.getColumnsToSkip().size();
        for (int i = 0; i < columnsToAdd; i++) {
            fieldValues.add(null);
        }

        try {
            printer.printRecord(fieldValues);
        } catch (IOException e) {
            e.printStackTrace();
        }

        rowCount++;
    }

    @Override
    public void cell(String cellReference, String formattedValue, XSSFComment comment) {
        if (skipRow) {
            return;
        }

        // gracefully handle missing CellRef here in a similar way as XSSFCell does
        if (cellReference == null) {
            cellReference = new CellAddress(currentRow, currentCol).formatAsString();
        }

        // Did we miss any cells?
        int thisCol = (new CellReference(cellReference)).getCol();

        // Should we skip this

        //Use the first row of the file to decide on the area of data to export
        if (firstRow && firstCellOfRow) {
            readConfig.setFirstRow(currentRow);
            readConfig.setFirstColumn(thisCol);
        }

        //if this cell falls outside our area, or has been explcitely marked as a skipped column, return and don't write it out.
        if (!firstRow && (thisCol < readConfig.getFirstColumn() || thisCol > readConfig.getLastColumn())) {
            return;
        }

        if (readConfig.getColumnsToSkip().contains(thisCol)) {
            skippedColumns++;
            return;
        }

        int missedCols = (thisCol - readConfig.getFirstColumn()) - (currentCol - readConfig.getFirstColumn())
                - 1;
        if (firstCellOfRow) {
            missedCols = (thisCol - readConfig.getFirstColumn());
        }

        missedCols -= skippedColumns;

        if (firstCellOfRow) {
            firstCellOfRow = false;
        }

        for (int i = 0; i < missedCols; i++) {
            fieldValues.add(null);
        }
        currentCol = thisCol;

        fieldValues.add(formattedValue);

        rowHasValues = true;
        skippedColumns = 0;
    }

    @Override
    public void headerFooter(String s, boolean b, String s1) {

    }

    public void close() throws IOException {
        printer.close();
    }
}