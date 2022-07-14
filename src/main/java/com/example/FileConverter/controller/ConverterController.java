package com.example.FileConverter.controller;

import java.io.OutputStream;
import java.util.Map;

import com.example.FileConverter.odt.GetFileDto;
import com.example.FileConverter.service.ConvertExcelToCSVService;

import org.modelmapper.ModelMapper;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;


@RestController
public class ConverterController {
    @Autowired
    ConvertExcelToCSVService convertExcelToCSV;

    @Autowired
    private ModelMapper modelMapper;

    @PostMapping(value = "/file",
 //           consumes = "multipart/form-data",
            produces = "application/zip")
    public void parseURL(@RequestBody GetFileDto getfiledto, HttpServletResponse response) throws Exception {

        try {
            byte[] Zip = convertExcelToCSV.convertExcelToCSV(Map.of(), getfiledto); //  размер файла lenght и рар

            try (OutputStream os = response.getOutputStream()) {
                os.write(Zip, 0, Zip.length);
                response.setContentType("application/zip");
                response.setHeader(HttpHeaders.CONTENT_DISPOSITION,"attachment; filename=test.zip");
            } catch (Exception e) {
                System.out.println("Не удалось вернуть файл!");
            }
        }
        //catch (WrongFileFormatException | CorruptedFileException e)
        catch (Exception e) {
            System.out.println("Не удалось создать файл!");
        }
    }

    @PostMapping(value = "/file1",
            produces = "application/zip")
    public void parseMultipartFile(@RequestParam Map<String, String> params,
                                                     @RequestBody MultipartFile multipartFile,
                                                     HttpServletResponse response) throws Exception {
        try {
            byte[] Zip = convertExcelToCSV.convertExcelToCSV(params, multipartFile); //  размер файла lenght и рар
            try (OutputStream os = response.getOutputStream()) {
                os.write(Zip, 0, Zip.length);
                os.close();
                response.setContentType("application/zip");
                response.setHeader(HttpHeaders.CONTENT_DISPOSITION,"attachment; filename=test.zip");
                response.setHeader(HttpHeaders.CONTENT_LENGTH, String.valueOf(Zip.length));
            } catch (Exception e) {
                System.out.println("Не удалось вернуть файл!");
            }

        }
        //catch (WrongFileFormatException | CorruptedFileException e)
        catch (Exception e) {
            System.out.println("Не удалось создать файл!");//controlleradvice
        }
    }
}