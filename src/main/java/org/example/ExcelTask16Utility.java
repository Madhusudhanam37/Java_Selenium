package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.util.*;

public class ExcelTask16Utility {

    String path= "src/main/resources/SuperStoreUS-2015.xlsx";
    DataFormatter dataFormatter = new DataFormatter();
    FileInputStream fis;
    public List<LinkedHashMap<String, String>> readRangeOFValues() {
        List<LinkedHashMap<String, String>> storeData = new LinkedList<>();
        List<String> fr = new LinkedList<>();

        try {
            fis = new FileInputStream(new File(path));
             Workbook workbook = new XSSFWorkbook(fis);
            Sheet sheet = workbook.getSheetAt(0);
            Row firstRow = sheet.getRow(0);
            if (firstRow != null) {
                for (Cell cell : firstRow) {
                    fr.add(dataFormatter.formatCellValue(cell));
                }
            }
            for (int rowIndex = 22; rowIndex < 56; rowIndex++) {
                Row row = sheet.getRow(rowIndex);
                if (row == null) {
                    continue;
                }

                LinkedHashMap<String, String> mapData = new LinkedHashMap<>();
                int cellIndex = 0;
                for (Cell cell : row) {
                    String cellValue = dataFormatter.formatCellValue(cell);
                    if (cellIndex < fr.size()) {
                        mapData.put(fr.get(cellIndex), cellValue);
                    }
                    cellIndex++;
                }
                storeData.add(mapData);
            }

            for(LinkedHashMap<String,String> values:storeData){
                System.out.println(values);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return storeData;
    }

    public String readCorrespondingValues(int rowIndex,String columnHeader){
        List<String> fr = new LinkedList<>();
        Cell cell = null;
        try {
            fis=new FileInputStream(new File(path));
            Workbook workbook=new XSSFWorkbook(fis);
            Sheet sheet= workbook.getSheetAt(0);
            Row firstRow = sheet.getRow(0);
            if (firstRow != null) {
                for (Cell cells : firstRow) {
                    fr.add(dataFormatter.formatCellValue(cells));
                }
            }
            Row row=sheet.getRow(rowIndex);
            int colIndex = fr.indexOf(columnHeader);
            cell = row.getCell(colIndex);
        }catch (Exception e){
            e.printStackTrace();
        }
        return dataFormatter.formatCellValue(cell);
    }
}
