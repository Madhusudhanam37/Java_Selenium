package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.LinkedList;
import java.util.List;

public class ExcelWritingTask16Utility {

    String path = "src/main/resources/SuperStoreUS-2015.xlsx";
    DataFormatter dataFormatter = new DataFormatter();
    FileInputStream fis;
    List<String> fr;

    public String writeSingleValue(int rowIndex, String columnHeader, String value){
        fr = new LinkedList<>();
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
            cell.setCellValue(value);
            try (FileOutputStream fos = new FileOutputStream(new File(path))) {
                workbook.write(fos);
            }
            workbook.close();
        }catch (Exception e){
            e.printStackTrace();
        }
        return dataFormatter.formatCellValue(cell);
    }

    public void copyRows(int startRowIndex, int endRowIndex, int targetStartRowIndex) {
        try (FileInputStream fis = new FileInputStream(new File(path));
             Workbook workbook = new XSSFWorkbook(fis)) {
            Sheet sheet = workbook.getSheetAt(0);
            Row firstRow = sheet.getRow(0);
            if (firstRow != null) {
                for (Cell cells : firstRow) {
                    fr.add(dataFormatter.formatCellValue(cells));
                }
            }

            for (int i = 0; i <= endRowIndex - startRowIndex; i++) {
                Row targetRow = sheet.getRow(targetStartRowIndex + i);
                if (targetRow == null) {
                    sheet.createRow(targetStartRowIndex + i);
                }
            }

            for (int i = 0; i <= endRowIndex - startRowIndex; i++) {
                Row sourceRow = sheet.getRow(startRowIndex + i);
                Row targetRow = sheet.getRow(targetStartRowIndex + i);

                for (int j = 0; j < fr.size(); j++) {
                    Cell sourceCell = sourceRow.getCell(j);
                    Cell targetCell = targetRow.getCell(j);
                }
            }

            try (FileOutputStream fos = new FileOutputStream(new File(path))) {
                workbook.write(fos);
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
