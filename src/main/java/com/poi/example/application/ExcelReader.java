package com.poi.example.application;

import com.poi.example.Application;
import com.poi.example.model.Resident;
import com.poi.example.utility.ParseExcelFile;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.springframework.boot.SpringApplication;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.File;
import java.io.IOException;
import java.util.Date;

public class ExcelReader {
    public static void main(String[] args) {
        final String FILE_NAME = "Resident_INFO.xlsx";

        long startTime = System.currentTimeMillis();
        System.out.println("EXCEL READ PROCESS ---> START AT --> " + new Date());

        try {
            // Read workbook
            readFile(FILE_NAME);
        } catch(Error e) {
            e.printStackTrace();
        } catch(Exception ex) {
            ex.printStackTrace();
        } finally {
            long endTime = System.currentTimeMillis();
            long timeElapsed = endTime - startTime;
            System.out.println("EXCEL READ PROCESS ---> FINISH AT --> " + new Date());
            System.out.println("TOTAL EXECUTION TIME: " + timeElapsed + " milliseconds");
        }
    }
    private static void readFile(String fileName)
            throws IOException, OpenXML4JException, ParserConfigurationException, SAXException {

        File file = new File(fileName);
        OPCPackage opcPackage = OPCPackage.open(file);
        ParseExcelFile parsedFile = new ParseExcelFile(opcPackage,10, Resident.class);
        boolean isValid = parsedFile.process();
        System.out.println("File is valid ? " + isValid);
    }
}