package org.example;

import org.junit.Test;

public class ExcelTest {

    ExcelUtility excelUtility=new ExcelUtility();

//    @Test
    public void exceltest(){
        excelUtility.excelReading(0,3);
    }

//    @Test
    public void excelFirstRowValue(){
        excelUtility.getFIrstRowValue();
    }

    @Test
    public void excelGetValues(){
        excelUtility.getValues();
    }
}
