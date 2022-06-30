package com.example.FileConverter.service;

import org.apache.commons.io.FilenameUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.*;
import java.util.Iterator;

@Service
public class ConverterService {

    public static byte[] xlsConvert(MultipartFile inputFile, String mime) {

        StringBuffer data = new StringBuffer();
        try {
            InputStream fis = new ByteArrayInputStream(inputFile.getBytes());

            Workbook workbook;
            if (mime == "xls"){
                workbook = new HSSFWorkbook(fis);
            }else {
                workbook = new XSSFWorkbook(fis);
            }
            int numberOfSheets = workbook.getNumberOfSheets();

            Row row;
            Cell cell;

            for (int i = 0; i < numberOfSheets; i++) {
                Sheet sheet = workbook.getSheetAt(0);
                Iterator<Row> rowIterator = sheet.iterator();

                while (rowIterator.hasNext()) {
                    row = rowIterator.next();

                    short firstCell = row.getFirstCellNum();
                    short lastCell = row.getLastCellNum();

                    String DELIMITER = "";

                    for(int j = firstCell; j<lastCell; j++){
                        cell = row.getCell(j);
                        if(cell==null){
                            data.append(DELIMITER);
                        }else {
                            switch (cell.getCellType()) {
                                case Cell.CELL_TYPE_BOOLEAN:
                                    data.append(DELIMITER + cell.getBooleanCellValue() );

                                    break;
                                case Cell.CELL_TYPE_NUMERIC:
                                    data.append(DELIMITER + cell.getNumericCellValue() );

                                    break;
                                case Cell.CELL_TYPE_STRING:
                                    data.append(DELIMITER + cell.getStringCellValue());
                                    break;

                                case Cell.CELL_TYPE_BLANK:
                                    data.append(DELIMITER );
                                    break;

                                default:
                                    data.append(DELIMITER + cell);
                            }
                        }
                        DELIMITER=",";
                    }
                    data.append('\n'); //
                }
            }
            return data.toString().getBytes();

        } catch (Exception ioe) {
            ioe.printStackTrace();
            return new byte[0];
        }
    }
}