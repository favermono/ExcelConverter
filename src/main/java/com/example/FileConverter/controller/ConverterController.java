package com.example.FileConverter.controller;

import java.io.IOException;
import java.io.OutputStream;
import java.util.Map;

import com.example.FileConverter.dto.GetFileDto;
import com.example.FileConverter.service.ConvertExcelToCSVService;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;


@RestController
public class ConverterController {
    @Autowired
    ConvertExcelToCSVService convertExcelToCSV;

    @PostMapping(value = "/json")
    public void parseURL(@RequestBody GetFileDto getfiledto, HttpServletResponse response) throws IOException {

            byte[] zip = convertExcelToCSV.convertExcelToCSV(Map.of(), getfiledto); //  размер файла lenght и рар
            OutputStream os = response.getOutputStream();
            os.write(zip, 0, zip.length);
            response.setContentType("application/zip");
            response.setHeader(HttpHeaders.CONTENT_DISPOSITION,"attachment; filename=output.zip");
            response.setContentLength(zip.length);
            os.close();
    }


    @PostMapping(value = "/multipart")
    public void parseMultipartFile(@RequestParam Map<String, String> params,
                                                     @RequestBody MultipartFile multipartFile,
                                                     HttpServletResponse response) throws IOException {
            if (multipartFile == null){
                throw new RuntimeException("For this request, you need to attach a \"multipartFile\" file. " +
                        "If you only have a link to the file, use \"host:8080/json\".");
            }
            byte[] zip = convertExcelToCSV.convertExcelToCSV(params, multipartFile); //  размер файла lenght и рар
            OutputStream os = response.getOutputStream();
            os.write(zip, 0, zip.length);
            response.setContentType("application/zip");
            response.setHeader(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=output.zip");
            response.setContentLength(zip.length);
            os.close();
    }

}