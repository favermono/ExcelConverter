package com.example.FileConverter.service;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;


import com.example.FileConverter.exceptions.BadDtoException;
import com.example.FileConverter.exceptions.BadLinkException;
import com.example.FileConverter.exceptions.WrongFileFormatException;
import com.example.FileConverter.odt.GetFileDto;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVPrinter;
import org.apache.commons.csv.QuoteMode;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.text.StringEscapeUtils;

import org.apache.nifi.components.PropertyDescriptor;

import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;


import static org.apache.nifi.csv.CSVUtils.*;


@Service
public class ConvertExcelToCSVService {

    private static final String DESIRED_SHEETS_DELIMITER = ",";


    public byte[] convertExcelToCSV(Map<String, String> params, Object model)
            throws IOException, IllegalArgumentException {


        InputStream fis = null;
        File targetFile = null;
        try {
            String directory = String.valueOf(generateUniqueString().hashCode()); // уникальное название папки и файлов для каждого запроса
            GetFileDto getFileDto;
            if (model.getClass() == GetFileDto.class) {     //создаем файл из приянтого inputStream'а
                getFileDto = (GetFileDto) model;
                fis = urlToInputStream(getFileDto.getURL());

                targetFile = new File(directory+".xlsx");
                FileUtils.copyInputStreamToFile(fis, targetFile);
                fis.close();
            } else {
                MultipartFile multipartFile = (MultipartFile) model;
                fis =  new BufferedInputStream(multipartFile.getInputStream());

                getFileDto = MapToDto (params);

                targetFile = new File(directory+".xlsx");
                FileUtils.copyInputStreamToFile(fis, targetFile);
                fis.close();
            }

            File fileDirectory = new File(directory); //создаем папку для хранения преобразованных csv файлов
            if (!fileDirectory.exists()){
                fileDirectory.mkdir();
            }

            ConvertFileToExcel(getFileDto,directory); //Записываем в папку "directory" преобразованные листы ексель в csv формат


            FileOutputStream fos = new FileOutputStream(directory +".zip");
            ZipOutputStream zipOut = new ZipOutputStream(fos);

            CreateZipFile(fileDirectory, directory,zipOut); //записываем все файлы из папки в зип файл

            zipOut.close();
            fos.close();

            FileUtils.deleteDirectory(fileDirectory); //удаляем папку c csv файлами


            File file = new File(directory +".zip");
            byte[] zipBytes = FileUtils.readFileToByteArray(file);
            FileUtils.delete(file);                     //удаляем zip file
            FileUtils.delete(targetFile);               //удаляем xlsx file
            return zipBytes; // FIXME: 14.07.2022 Возвращаем файл в виде массива байтов (как оптимальнее?)

        } catch (Exception e) {
            System.out.println("Unexpected exception");
            e.printStackTrace();
            return null;       // FIXME: 14.07.2022 Ошибки пока не обрабатываются

        }
    }



    private void ConvertFileToExcel(final GetFileDto GetFileDto, String filename) throws  Exception {


        final String desiredSheetsDelimited = GetFileDto.getDESIRED_SHEETS();
        final boolean formatValues = GetFileDto.isFORMAT_VALUES();

        final CSVFormat csvFormat = createCSVFormat(GetFileDto);

        //Switch to 0 based index
        final int firstRow = GetFileDto.getROWS_TO_SKIP() - 1 ;
        final String[] sColumnsToSkip = StringUtils.split(GetFileDto.getCOLUMNS_TO_SKIP(), ",");

        final List<Integer> columnsToSkip = new ArrayList<>();

        if (sColumnsToSkip != null && sColumnsToSkip.length > 0) {
            for (String c : sColumnsToSkip) {
                try {
                    //Switch to 0 based index
                    columnsToSkip.add(Integer.parseInt(c) - 1);
                } catch (NumberFormatException e) {
                    throw new Exception("Invalid column in Columns to Skip list.", e);
                }
            }
        }

        try {

            File initialFile = new File(filename+".xlsx");
            InputStream inputStream = new FileInputStream(initialFile);

            OPCPackage pkg = OPCPackage.open(inputStream);
            XSSFReader r = new XSSFReader(pkg);
            ReadOnlySharedStringsTable sst = new ReadOnlySharedStringsTable(pkg);
            StylesTable styles = r.getStylesTable();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) r.getSheetsData();

            if (desiredSheetsDelimited != null) {
                String[] desiredSheets = StringUtils.split(desiredSheetsDelimited,
                        DESIRED_SHEETS_DELIMITER);

                if (desiredSheets != null) {
                    while (iter.hasNext()) {
                        InputStream sheet = iter.next();
                        String sheetName = iter.getSheetName();

                        for (String desiredSheet : desiredSheets) {
                            //If the sheetName is a desired one parse it
                            if (sheetName.equalsIgnoreCase(desiredSheet)) {
                                ExcelSheetReadConfig readConfig = new ExcelSheetReadConfig(
                                        columnsToSkip, firstRow, sheetName, formatValues, sst, styles);

                                handleExcelSheet(sheet, readConfig,
                                        csvFormat, filename+"\\"+updateFilenameToCSVExtension(filename,sheetName));
                                break;
                            }
                        }
                        sheet.close();
                    }
                } else {
                    System.out.println("Excel document was parsed but no sheets with the specified desired names were found.");
                }

            } else {
                //Get all of the sheets in the document.
                while (iter.hasNext()) {
                    InputStream sheet = iter.next();
                    String sheetName = iter.getSheetName();

                    ExcelSheetReadConfig readConfig = new ExcelSheetReadConfig(columnsToSkip, firstRow,
                            sheetName, formatValues, sst, styles);
                    File targetFile = new File(sheetName+".csv");
                    OutputStream outStream = new FileOutputStream(targetFile);
                    handleExcelSheet( sheet, readConfig,
                            csvFormat, filename+"\\"+updateFilenameToCSVExtension(filename,sheetName));
                    outStream.close();
                }
            }
            inputStream.close();
            pkg.close();
        } catch (InvalidFormatException ife) {
            System.out.println("Only .xlsx Excel 2007 OOXML files are supported");
            throw new UnsupportedOperationException("Only .xlsx Excel 2007 OOXML files are supported", ife);
        } catch (OpenXML4JException | SAXException e) {
            System.out.println("Error occurred while processing Excel document metadata");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
        finally {

        }
    }


    private void handleExcelSheet(final InputStream sheetInputStream, ExcelSheetReadConfig readConfig, CSVFormat csvFormat, String filename)
            throws IOException {
        try {
            final DataFormatter formatter = new DataFormatter();
            final InputSource sheetSource = new InputSource(sheetInputStream);

            final SheetToCSV sheetHandler = new SheetToCSV(readConfig, csvFormat);

            final XMLReader parser = SAXHelper.newXMLReader();

            //If Value Formatting is set to false then don't pass in the styles table.
            // This will cause the XSSF Handler to return the raw value instead of the formatted one.
            final StylesTable sst = readConfig.getFormatValues() ? readConfig.getStyles() : null;

            final XSSFSheetXMLHandler handler = new XSSFSheetXMLHandler(sst, null,
                    readConfig.getSharedStringsTable(), sheetHandler, formatter, false);

            parser.setContentHandler(handler);


            File targetFile = new File(filename);
            OutputStream out= new FileOutputStream(targetFile); // FIXME: 14.07.2022 не забыть проверить
            PrintStream outPrint = new PrintStream(out);
            sheetHandler.setOutput(outPrint);

            try {
                parser.parse(sheetSource);

                sheetInputStream.close();
                sheetHandler.close();
                outPrint.close();
                out.close();
            } catch (SAXException se) {
                System.out.println("Error occurred while processing Excel sheet {}" + readConfig.getSheetName());
            }



        } catch (SAXException | ParserConfigurationException saxE) {
            System.out.println("Failed to create instance of Parser.");
        } catch (Exception e) {
            throw new RuntimeException(e);
        } finally {
            sheetInputStream.close();
        }
    }

    private class SheetToCSV implements XSSFSheetXMLHandler.SheetContentsHandler {
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

        public void setOutput(PrintStream output) throws Exception {
            final OutputStreamWriter streamWriter = new OutputStreamWriter(output);

            try {
                printer = new CSVPrinter(streamWriter, csvFormat);
            } catch (IOException e) {
                throw new Exception("Failed to create CSV Printer.", e);
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

    /**
     * Takes the original input filename and updates it by removing the file extension and replacing it with
     * the .csv extension.
     */
    private String updateFilenameToCSVExtension(String origFileName, String sheetName) {

        StringBuilder stringBuilder = new StringBuilder();

        if (StringUtils.isNotEmpty(origFileName)) {
            String ext = FilenameUtils.getExtension(origFileName);
            if (StringUtils.isNotEmpty(ext)) {
                stringBuilder.append(StringUtils.replace(origFileName, ("." + ext), ""));
            } else {
                stringBuilder.append(origFileName);
            }
        } else {
            stringBuilder.append(generateUniqueString());
        }

        stringBuilder.append("_");
        stringBuilder.append(sheetName);
        stringBuilder.append(".");
        stringBuilder.append("csv");

        return stringBuilder.toString();
    }

    private class ExcelSheetReadConfig {
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
    }

    private CSVFormat createCSVFormat(GetFileDto dto) {
        String formatName = dto.getCSV_FORMAT();
        if (formatName.equalsIgnoreCase("custom")) {
            return buildCustomFormat(dto);
        } else if (formatName.equalsIgnoreCase("rfc-4180")) {
            return CSVFormat.RFC4180;
        } else if (formatName.equalsIgnoreCase("excel")) {
            return CSVFormat.EXCEL;
        } else if (formatName.equalsIgnoreCase("tdf")) {
            return CSVFormat.TDF;
        } else if (formatName.equalsIgnoreCase("mysql")) {
            return CSVFormat.MYSQL;
        } else if (formatName.equalsIgnoreCase("informix-unload")) {
            return CSVFormat.INFORMIX_UNLOAD;
        } else {
            return formatName.equalsIgnoreCase("informix-unload-csv") ? CSVFormat.INFORMIX_UNLOAD_CSV : CSVFormat.DEFAULT;
        }
    }
    private CSVFormat buildCustomFormat(GetFileDto GetFileDto) {
        Character valueSeparator = getValueSeparatorCharUnescapedJava(GetFileDto.getVALUE_SEPARATOR());
        CSVFormat format = CSVFormat.newFormat(valueSeparator).withAllowMissingColumnNames().withIgnoreEmptyLines();
        if (GetFileDto.getFIRST_LINE_IS_HEADER() != null && GetFileDto.getFIRST_LINE_IS_HEADER()) {
            format = format.withFirstRecordAsHeader();
        }

        Character quoteChar = getCharUnescaped(GetFileDto.getQUOTE_CHAR(), QUOTE_CHAR);
        format = format.withQuote(quoteChar);
        Character escapeChar = GetFileDto.getESCAPE_CHAR().isEmpty() ? null : getCharUnescaped(GetFileDto.getESCAPE_CHAR(), ESCAPE_CHAR);
        format = format.withEscape(escapeChar);
        format = format.withTrim(GetFileDto.getTRIM_FIELDS());
        if (GetFileDto.getCOMMENT_MAKER() != null) {
            Character commentMarker = getCharUnescaped(GetFileDto.getCOMMENT_MAKER(), COMMENT_MARKER);
            if (commentMarker != null) {
                format = format.withCommentMarker(commentMarker);
            }
        }
        if (GetFileDto.getNULL_STRING() != null) {
            format = format.withNullString(unescape(GetFileDto.getNULL_STRING()));
        }


        if (GetFileDto.getQUOTE_MODE() != null) {
            QuoteMode quoteMode = QuoteMode.valueOf(GetFileDto.getQUOTE_MODE());
            format = format.withQuoteMode(quoteMode);
        }


        if (GetFileDto.getTRAILING_DELIMITER() != null) {
            format = format.withTrailingDelimiter(GetFileDto.getTRAILING_DELIMITER());
        }

        if (GetFileDto.getRECORD_SEPARATOR() != null) {
            format = format.withRecordSeparator(GetFileDto.getRECORD_SEPARATOR());
        }

        if (GetFileDto.getALLOW_DUPLICATE_HEADER_NAMES()!= null) {
            format = format.withAllowDuplicateHeaderNames(GetFileDto.getALLOW_DUPLICATE_HEADER_NAMES());
        }

        return format;
    }

    private Character getValueSeparatorCharUnescapedJava(String value) {
        if (value != null) {
            String unescaped = unescape(value);
            if (unescaped.length() == 1) {
                return unescaped.charAt(0);
            }
        }

        //LOG.warn("'{}' property evaluated to an invalid value: \"{}\". It must be a single character. The property value will be ignored.", VALUE_SEPARATOR.getName(), value);
        return VALUE_SEPARATOR.getDefaultValue().charAt(0);
    }

    private Character getCharUnescaped(String value, PropertyDescriptor property) {

        if (value != null) {
            String unescaped = unescape(value);
            if (unescaped.length() == 1) {
                return unescaped.charAt(0);
            }
        }

        //LOG.warn("'{}' property evaluated to an invalid value: \"{}\". It must be a single character. The property value will be ignored.", property.getName(), value);
        return property.getDefaultValue() != null ? property.getDefaultValue().charAt(0) : null;
    }

    public String unescape(String input) {
        if (input != null && input.length() > 1) {
            input = StringEscapeUtils.unescapeJava(input);
        }

        return input;
    }


    //nach

    private String generateUniqueString() {
        DateFormat dateFormat = new SimpleDateFormat("yyyyMMdd_HHmmss");
        Date date = new Date();
        return dateFormat.format(date);

    }

    private InputStream urlToInputStream(String urlString) throws IOException, BadLinkException, WrongFileFormatException {
        HttpURLConnection con;
        InputStream inputStream;
        final URL url = new URL(urlString);
        con = (HttpURLConnection) url.openConnection();
        con.setConnectTimeout(15000);
        con.setReadTimeout(15000);
        con.connect();
        int responseCode = con.getResponseCode();
        if (responseCode != 200) { // FIXME: 11.07.2022 Возможно не работает
            throw new BadLinkException("Unable to read from given URL");
        }
        String filename =urlString.toLowerCase().substring(urlString.lastIndexOf("/") + 1);
        if (!(filename.endsWith("xls") || filename.endsWith("xlsx"))) {
            throw new WrongFileFormatException("Wrong file type: Only support .xls or .xlsx file!");
        }
        inputStream = con.getInputStream();
        return inputStream;
    }



    private GetFileDto MapToDto(Map<String,String> map) throws BadDtoException { // FIXME: 14.07.2022 Пока не работает должным образом
        try {
            return GetFileDto.builder().URL("Empty")
                    .CSV_FORMAT("EXCEL").build();
        } catch (Exception e){
            throw new BadDtoException("Wrong parameters!");
        }
    }


    private void CreateZipFile(File fileToZip, String fileName, ZipOutputStream zipOut) throws IOException {
        if (fileToZip.isHidden()) {
            return;
        }
        if (fileToZip.isDirectory()) {
            if (fileName.endsWith("/")) {
                zipOut.putNextEntry(new ZipEntry(fileName));
                zipOut.closeEntry();
            } else {
                zipOut.putNextEntry(new ZipEntry(fileName + "/"));
                zipOut.closeEntry();
            }
            File[] children = fileToZip.listFiles();
            for (File childFile : children) {
                CreateZipFile(childFile, fileName + "/" + childFile.getName(), zipOut);
            }
            return;
        }
        FileInputStream fis = new FileInputStream(fileToZip);
        ZipEntry zipEntry = new ZipEntry(fileName);
        zipOut.putNextEntry(zipEntry);
        byte[] bytes = new byte[1024];
        int length;
        while ((length = fis.read(bytes)) >= 0) {
            zipOut.write(bytes, 0, length);
        }
        fis.close();
    }
}