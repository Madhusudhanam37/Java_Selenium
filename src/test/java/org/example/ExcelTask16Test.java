package org.example;

import org.junit.Test;

public class ExcelTask16Test {

    ExcelTask16Utility excelTask16Utility=new ExcelTask16Utility();

//    @Test
    public void readDatas(){
        excelTask16Utility.readRangeOFValues();
    }

//    @Test
    public void readValue(){
        System.out.println(excelTask16Utility.readCorrespondingValues(24,"Postal Code"));
    }


}
