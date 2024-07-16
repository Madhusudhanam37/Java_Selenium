package org.example;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.util.Iterator;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;

public class ExcelUtility {

    DataFormatter dataFormatter=new DataFormatter();

    public void excelReading(int rownum,int columnnum) {
        String path = "src/main/resources/Bank Ledger with bank details.xlsx";
        Workbook workbook = null;
        try {
            File file = new File(path);
            workbook = new XSSFWorkbook(file);
        } catch (Exception e) {
            e.printStackTrace();
        }
        Sheet sheet = workbook.getSheetAt(0);
        Row row= sheet.getRow(rownum);
        Cell cell= row.getCell(columnnum);
        String cellValue=dataFormatter.formatCellValue(cell);
        System.out.println(cellValue);
    }

    public List<LinkedHashMap<String,String>> getFIrstRowValue(){
        List<LinkedHashMap<String,String>> testData=new LinkedList<LinkedHashMap<String, String>>();
        List<String> firstRow=new LinkedList<>();
        Workbook workbook=null;
        String path="src/main/resources/Bank Ledger with bank details.xlsx";
        try {
            File file=new File(path);
            workbook=new XSSFWorkbook(file);
        }catch (Exception e){
            e.printStackTrace();
        }
        String sheetName=workbook.getSheetName(0);
        System.out.println("SheetName is: "+sheetName);
        Sheet sheet=workbook.getSheet(sheetName);
        Row row;
        Iterator rowiterator=sheet.rowIterator();
        Iterator celliterator;
        while (rowiterator.hasNext()){
            row= (Row) rowiterator.next();
            celliterator=row.cellIterator();
            while (celliterator.hasNext()){
                firstRow.add(dataFormatter.formatCellValue((Cell) celliterator.next()));
            }
            break;
        }
        for(String value:firstRow) {
            System.out.println(value);
        }
        return  null;
    }

    public List<LinkedHashMap<String,String>> getValues(){
        List<LinkedHashMap<String,String>> allData=new LinkedList<LinkedHashMap<String,String>>();
        List<String> firstRow=new LinkedList<>();
        String path="src/main/resources/Bank Ledger with bank details.xlsx";
        Workbook workbook=null;
        try{
            File file=new File(path);
            workbook=new XSSFWorkbook(file);
        }catch (Exception e){
            e.printStackTrace();
        }
        String sheetName=workbook.getSheetName(0);
        Sheet sheet=workbook.getSheet(sheetName);
        Row row;
        Iterator rowiterator= sheet.rowIterator();
        Iterator celliterator;
        while (rowiterator.hasNext()){
            row= (Row) rowiterator.next();
            celliterator=row.cellIterator();
            while (celliterator.hasNext()){
                firstRow.add(dataFormatter.formatCellValue((Cell) celliterator.next()));
            }
            break;
        }
        int count=0;
        rowiterator= sheet.rowIterator();
        while (rowiterator.hasNext()){
            if(count==4){
                break;
            }
            LinkedHashMap<String,String> restRow=new LinkedHashMap<>();
            row= (Row) rowiterator.next();
            celliterator=row.cellIterator();
            int firstRowIndex=0;
            while (celliterator.hasNext()){
                Cell cell= (Cell) celliterator.next();
                String cellValue=dataFormatter.formatCellValue(cell);
                restRow.put(firstRow.get(firstRowIndex), cellValue);
                firstRowIndex++;
            }
            count++;
            allData.add(restRow);
        }
        for(LinkedHashMap<String,String> map:allData){
            System.out.println(map);
        }
        return allData;
    }
}
