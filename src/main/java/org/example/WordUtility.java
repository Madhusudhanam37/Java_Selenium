package org.example;

import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;


public class WordUtility {

    public void readWordUsingScanner() {
        String path = "src/main/resources/WordFile.docx";
        File file = new File(path);

        try(FileInputStream fileInputStream=new FileInputStream(file)) {
            XWPFDocument xwpfDocument = new XWPFDocument(fileInputStream);
            XWPFWordExtractor xwpfWordExtractor=new XWPFWordExtractor(xwpfDocument);
            String text=xwpfWordExtractor.getText();
            System.out.println(text);
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
//public void readWordUsingScanner() {
//    String path = "src/main/resources/WordFile.docx";
//    File file = new File(path);
//
//    if (!file.exists()) {
//        System.out.println("File does not exist at path: " + path);
//        return;
//    }
//
//    System.out.println("File found, attempting to read...");
//
//    try (FileInputStream fileInputStream = new FileInputStream(file);
//         XWPFDocument xwpfDocument = new XWPFDocument(fileInputStream)) {
//
//        String text = xwpfDocument.getDocument().getBody().toString();
//        System.out.println("File content: " + text);
//
//    } catch (Exception e) {
//        System.out.println("An error occurred while reading the file.");
//        e.printStackTrace();
//    }
//}
}
