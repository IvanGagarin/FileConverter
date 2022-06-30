package com.example.FileConverter.controller;

import java.io.*;

import org.springframework.core.io.InputStreamResource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import static com.example.FileConverter.service.ConverterService.xlsConvert;


@RestController
public class ConverterController
{

    @PostMapping(value = "/file",
            consumes = MediaType.MULTIPART_FORM_DATA_VALUE,
            produces = "text/csv")
    public ResponseEntity<byte[]> parseMultipartFile(@RequestParam MultipartFile multipartFile) throws Exception {
        String fileContentType = multipartFile.getContentType();
        try {
            String mime = multipartFile.getOriginalFilename().split("\\.")[1];
            if ((mime.equalsIgnoreCase("xls")) || mime.equalsIgnoreCase("xlsx")) {

                byte[] output = xlsConvert(multipartFile,mime);
                if (output == null){
                    throw new Exception("meme");
                }
                HttpHeaders responseHeaders = new HttpHeaders();
                responseHeaders.add("Content-Type", "text/csv");
                responseHeaders.add("Content-Disposition", "attachment; filename=output.csv");
                return new ResponseEntity<>(output, responseHeaders, HttpStatus.OK);
            }else {
                throw new Exception();
            }
        } catch (Exception e){
            throw new Exception("Wrong file type: Import .xls or .xlsx file!");
        }
    }
}