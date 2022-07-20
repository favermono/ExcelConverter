package com.example.FileConverter.service;

import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;

import java.util.List;

public class ExcelSheetReadConfig {

    private int firstColumn;
    private int lastColumn;

    private int firstRow;
    private int lastRow;
    private int overrideFirstRow;
    private String sheetName;
    private boolean formatValues;

    private ReadOnlySharedStringsTable sst;
    private StylesTable styles;

    private List<Integer> columnsToSkip;

    public ExcelSheetReadConfig(List<Integer> columnsToSkip, int overrideFirstRow, String sheetName,
                                boolean formatValues, ReadOnlySharedStringsTable sst, StylesTable styles) {

        this.sheetName = sheetName;
        this.columnsToSkip = columnsToSkip;
        this.overrideFirstRow = overrideFirstRow;
        this.formatValues = formatValues;

        this.sst = sst;
        this.styles = styles;
    }

    public String getSheetName() {
        return sheetName;
    }

    public int getFirstColumn() {
        return firstColumn;
    }

    public void setFirstColumn(int value) {
        this.firstColumn = value;
    }

    public int getLastColumn() {
        return lastColumn;
    }

    public void setLastColumn(int lastColumn) {
        this.lastColumn = lastColumn;
    }

    public int getOverrideFirstRow() {
        return overrideFirstRow;
    }

    public boolean getFormatValues() {
        return formatValues;
    }

    public int getFirstRow() {
        return firstRow;
    }

    public void setFirstRow(int value) {
        firstRow = value;
    }

    public int getLastRow() {
        return lastRow;
    }

    public void setLastRow(int value) {
        lastRow = value;
    }

    public List<Integer> getColumnsToSkip() {
        return columnsToSkip;
    }

    public ReadOnlySharedStringsTable getSharedStringsTable() {
        return sst;
    }

    public StylesTable getStyles() {
        return styles;
    }

}