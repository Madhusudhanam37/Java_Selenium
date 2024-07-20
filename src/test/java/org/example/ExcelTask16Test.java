package org.example;

import org.junit.Test;

public class ExcelTask16Test {

    ExcelTask16Utility excelTask16Utility=new ExcelTask16Utility();
    ExcelWritingTask16Utility excelWritingTask16Utility=new ExcelWritingTask16Utility();

//    @Test
    public void readDatas(){
        excelTask16Utility.readRangeOFValues();
    }

//    @Test
    public void readValue(){
        System.out.println(excelTask16Utility.readCorrespondingValues(24,"Postal Code"));
    }


//    @Test
    public void writingSingleData(){
        System.out.println(excelWritingTask16Utility.writeSingleValue(24,"Postal Code","123456"));
    }

    @Test
    public void writingCpoyData(){
        excelWritingTask16Utility.copyRows(70, 79, 10);
    }

}
