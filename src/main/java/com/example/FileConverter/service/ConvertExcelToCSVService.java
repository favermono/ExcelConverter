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
import com.example.FileConverter.exceptions.ParserException;
import com.example.FileConverter.exceptions.WrongFileFormatException;
import com.example.FileConverter.odt.GetFileDto;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.QuoteMode;
import org.apache.commons.io.FileUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.EnumUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.text.StringEscapeUtils;

import org.apache.nifi.components.PropertyDescriptor;

import org.apache.poi.ooxml.util.SAXHelper;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;


import static org.apache.commons.lang3.StringUtils.split;
import static org.apache.nifi.csv.CSVUtils.*;


@Service
public class ConvertExcelToCSVService {


    private static final String DESIRED_SHEETS_DELIMITER = ",";


    public byte[] convertExcelToCSV(Map<String, String> params, Object model)
            throws IOException {
        String filename;
        File fileDirectory = null;
        File file = null;
        InputStream fis=InputStream.nullInputStream();
        File targetFile = null;
        FileOutputStream fos = null;
        ZipOutputStream zipOut = null;
        try {
            String directory = String.valueOf(generateUniqueString().hashCode()); // уникальное название папки и файлов для каждого запроса
            GetFileDto getFileDto;
            if (model.getClass() == GetFileDto.class) {     //создаем файл из приянтого inputStream'а
                getFileDto = (GetFileDto) model;
                filename =getFileDto.getUrl().toLowerCase()
                        .substring(getFileDto.getUrl().lastIndexOf("/") + 1)
                        .split("\\.")[0];
                fis = urlToInputStream(getFileDto.getUrl());

            } else {
                MultipartFile multipartFile = (MultipartFile) model;
                filename = Objects.requireNonNull(multipartFile.getOriginalFilename()).split("\\.")[0] != null?
                        multipartFile.getOriginalFilename().split("\\.")[0] : "unknown.xlsx" ;
                fis =  new BufferedInputStream(multipartFile.getInputStream());
                getFileDto = mapToDto(params);
            }
            targetFile = new File(filename+directory+".xlsx");
            FileUtils.copyInputStreamToFile(fis, targetFile);
            fis.close();

            fileDirectory = new File(directory); //создаем папку для хранения преобразованных csv файлов
            if (!fileDirectory.exists()) {
                fileDirectory.mkdir();
            }

            convertFileToExcel(getFileDto,filename,directory); //Записываем в папку "directory" преобразованные листы ексель в csv формат

            fos = new FileOutputStream(filename+directory +".zip");
            zipOut = new ZipOutputStream(fos);

            createZipFile(fileDirectory, directory,zipOut); //записываем все файлы из папки в зип файл

            zipOut.close();
            fos.close();


            file = new File(filename+directory +".zip");
            return FileUtils.readFileToByteArray(file);

        }
        finally {
            fis.close();
            if (fos !=null) { fos.close(); }
            if (zipOut != null) { zipOut.close(); }
            if (fos != null) { fos.close(); }
            if (fileDirectory!= null) { FileUtils.deleteDirectory(fileDirectory); }//удаляем папку c csv файлами
            if (file != null)  { FileUtils.delete(file); }                        //удаляем zip file
            if (targetFile != null) { FileUtils.delete(targetFile); }             //удаляем xlsx file
        }
    }



    private void convertFileToExcel(final GetFileDto getFileDto, String filename, String directory) {


        final String desiredSheetsDelimited = getFileDto.getDesired_sheets();
        final boolean formatValues = getFileDto.isFormat_values();

        final CSVFormat csvFormat = createCSVFormat(getFileDto);
        //Switch to 0 based index
        final int firstRow = getFileDto.getRow_to_skip()!= null ? Integer.parseInt(getFileDto.getRow_to_skip()) - 1 : -1;
        final String[] sColumnsToSkip = split(getFileDto.getColumns_to_skip(), ",");
        final List<Integer> columnsToSkip = new ArrayList<>();

        if (sColumnsToSkip != null && sColumnsToSkip.length > 0) {
            for (String c : sColumnsToSkip) {
                try {
                    //Switch to 0 based index
                    columnsToSkip.add(Integer.parseInt(c) - 1);
                } catch (NumberFormatException e) {
                    throw new BadDtoException("Invalid column in Columns to Skip list.");
                }
            }
        }

        try {
            File initialFile = new File(filename+directory+".xlsx");
            InputStream inputStream = new FileInputStream(initialFile);

            OPCPackage pkg = OPCPackage.open(inputStream);
            XSSFReader r = new XSSFReader(pkg);
            ReadOnlySharedStringsTable sst = new ReadOnlySharedStringsTable(pkg);
            StylesTable styles = r.getStylesTable();
            XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) r.getSheetsData();

            if (desiredSheetsDelimited != null && !desiredSheetsDelimited.isEmpty()) {
                String[] desiredSheets = split(desiredSheetsDelimited,
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
                                        csvFormat, directory+"\\"+updateFilenameToCSVExtension(filename,sheetName));
                                break;
                            }
                        }
                        sheet.close();
                    }
                } else {
                    throw new BadDtoException("Excel document was parsed but no sheets with the specified desired names were found.");
                }

            } else {
                //Get all of the sheets in the document.
                while (iter.hasNext()) {
                    InputStream sheet = iter.next();
                    String sheetName = iter.getSheetName();

                    ExcelSheetReadConfig readConfig = new ExcelSheetReadConfig(columnsToSkip, firstRow,
                            sheetName, formatValues, sst, styles);

                    handleExcelSheet( sheet, readConfig,
                            csvFormat, directory+"\\"+updateFilenameToCSVExtension(filename,sheetName));
                    sheet.close();
                }
            }
            inputStream.close();
            pkg.close();
        } catch (InvalidFormatException ife) {
            throw new WrongFileFormatException("Only .xlsx Excel 2007 OOXML files are supported");
        } catch (OpenXML4JException | SAXException e) {
            throw new ParserException("Error occurred while processing Excel document metadata");
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }


    private void handleExcelSheet(final InputStream sheetInputStream, ExcelSheetReadConfig readConfig, CSVFormat csvFormat, String filename)
            throws IOException, ParserException {
        try (sheetInputStream) {
            DataFormatter formatter = new DataFormatter();
            InputSource sheetSource = new InputSource(sheetInputStream);

            SheetToCSV sheetHandler = new SheetToCSV(readConfig, csvFormat);

            XMLReader parser = SAXHelper.newXMLReader();

            //If Value Formatting is set to false then don't pass in the styles table.
            // This will cause the XSSF Handler to return the raw value instead of the formatted one.
            StylesTable sst = readConfig.getFormatValues() ? readConfig.getStyles() : null;

            XSSFSheetXMLHandler handler = new XSSFSheetXMLHandler(sst, null,
                    readConfig.getSharedStringsTable(), sheetHandler, formatter, false);

            parser.setContentHandler(handler);


            File targetFile = new File(filename);
            OutputStream out = new FileOutputStream(targetFile);
            PrintStream outPrint = new PrintStream(out);
            sheetHandler.setOutput(outPrint);

            try {
                parser.parse(sheetSource);

                sheetHandler.close();
                outPrint.close();
                out.close();
            } catch (SAXException se) {
                throw new ParserException("Error occurred while processing Excel sheet {}" + readConfig.getSheetName());
            }
        } catch (SAXException | ParserConfigurationException saxE) {
            throw new ParserException("Failed to create instance of Parser while proceed file.");
        } catch (Exception e) {
            throw new RuntimeException("Failed to convert file.");
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


    private CSVFormat createCSVFormat(GetFileDto dto) throws BadDtoException {
        String formatName = dto.getCsv_format() != null ? dto.getCsv_format() : "custom" ;
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
    private CSVFormat buildCustomFormat(GetFileDto getFileDto) throws BadDtoException {
        try {
            Character valueSeparator = getValueSeparatorCharUnescapedJava(getFileDto.getValue_separator());
            CSVFormat format = CSVFormat.newFormat(valueSeparator).withAllowMissingColumnNames().withIgnoreEmptyLines();
            if (getFileDto.getFirst_line_is_header() == null || getFileDto.getFirst_line_is_header()) {
                format = format.withFirstRecordAsHeader();
            }

            Character quoteChar = getCharUnescaped(getFileDto.getQuote_char(), QUOTE_CHAR);
            format = format.withQuote(quoteChar);
            Character escapeChar;
            if (getFileDto.getEscape_char() == null || getFileDto.getEscape_char().isEmpty()) {
                escapeChar = null;
            } else {
                escapeChar = getCharUnescaped(getFileDto.getEscape_char(), ESCAPE_CHAR);
            }
            format = format.withEscape(escapeChar);

            format = format.withTrim(getFileDto.getTrim_fields() == null || getFileDto.getTrim_fields());
            if (getFileDto.getComment_maker() != null) {
                Character commentMarker = getCharUnescaped(getFileDto.getComment_maker(), COMMENT_MARKER);
                if (commentMarker != null) {
                    format = format.withCommentMarker(commentMarker);
                }
            }
            if (getFileDto.getNull_string() != null) {
                format = format.withNullString(unescape(getFileDto.getNull_string()));
            }

            if (getFileDto.getQuote_mode() != null && EnumUtils.isValidEnum(QuoteMode.class, getFileDto.getQuote_mode())
                    && !getFileDto.getQuote_mode().equals("ALL_NON_NULL")) {
                QuoteMode quoteMode = QuoteMode.valueOf(getFileDto.getQuote_mode());
                format = format.withQuoteMode(quoteMode);
            } else {
                format = format.withQuoteMode(QuoteMode.MINIMAL);
            }
            format = format.withTrailingDelimiter(getFileDto.getTrailing_delimiter() != null &&
                                                                            getFileDto.getTrailing_delimiter());
            if (getFileDto.getRecord_separator() != null) {
                String separator = unescape(getFileDto.getRecord_separator());
                format = format.withRecordSeparator(separator);
            } else {
                format = format.withRecordSeparator("\n");
            }
            format = format.withAllowDuplicateHeaderNames((getFileDto.getAllow_duplicate_header_names() == null ||
                                                                            getFileDto.getAllow_duplicate_header_names()));
            return format;
        } catch (Exception e){
            throw new BadDtoException("Given parameters are incorrect!");
        }

    }

    private Character getValueSeparatorCharUnescapedJava(String value) {
        if (value != null) {
            String unescaped = unescape(value);
            if (unescaped.length() == 1) {
                return unescaped.charAt(0);
            }
        }

        return VALUE_SEPARATOR.getDefaultValue().charAt(0);
    }

    private Character getCharUnescaped(String value, PropertyDescriptor property) {

        if (value != null) {
            String unescaped = unescape(value);
            if (unescaped.length() == 1) {
                return unescaped.charAt(0);
            }
        }

        return property.getDefaultValue() != null ? property.getDefaultValue().charAt(0) : null;
    }

    private String unescape(String input) {
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
        HttpURLConnection con = null;
        InputStream inputStream = null;
        try {
            try {
                final URL url = new URL(urlString);
                con = (HttpURLConnection) url.openConnection();
                con.setConnectTimeout(15000);
                con.setReadTimeout(15000);
                con.connect();
            } catch (Exception e) {
                if (con != null) {
                    con.disconnect();
                }
                throw new RuntimeException("Unable to read from given URL.");
            }
            int responseCode = con.getResponseCode();
            if (responseCode != 200) {
                throw new BadLinkException("Unable to read from given URL.");
            }
            String filename = urlString.toLowerCase().substring(urlString.lastIndexOf("/") + 1);
            if (!filename.endsWith("xlsx")) {
                throw new WrongFileFormatException("Wrong file type: Only support .xlsx file!");
            }
            inputStream = con.getInputStream();
            return inputStream;
        }
        finally {
            if (con != null) {con.connect();}
        }
    }



    private GetFileDto mapToDto(Map<String,String> map) throws BadDtoException {
        try {
            GetFileDto getFileDto = GetFileDto.builder().url("MULTIPART_FILE").build();
            BeanUtils.populate(getFileDto,map);
            return getFileDto;
        } catch (Exception e){
            throw new BadDtoException("Given parameters are incorrect!");
        }
    }





    private void createZipFile(File fileToZip, String fileName, ZipOutputStream zipOut) throws IOException {
        FileInputStream fis=null;
        try {
            if (fileToZip.isHidden()) {
                return;
            }
            if (fileToZip.isDirectory()) {
                File[] children = fileToZip.listFiles();
                for (File childFile : children) {
                    createZipFile(childFile, childFile.getName(), zipOut);
                }
                return;
            }
            fis = new FileInputStream(fileToZip);
            ZipEntry zipEntry = new ZipEntry(fileName);
            zipOut.putNextEntry(zipEntry);
            byte[] bytes = new byte[1024];
            int length;
            while ((length = fis.read(bytes)) >= 0) {
                zipOut.write(bytes, 0, length);
            }
        } catch (IOException e) {
            throw new RuntimeException("Failed to create zip file.");
        } finally {
            if (fis != null) {
                fis.close();
            }
        }
    }
}