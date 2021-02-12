package com.poi.example.application;

import com.poi.example.utility.ReadExcelParser;

import java.util.Date;
import java.util.List;

public class ExcelReader {
    public static void main(String[] args) {
        final String FILE_NAME = "Resident_INFO.xlsx";

        long startTime = System.currentTimeMillis();
        System.out.println("EXCEL READ PROCESS ---> START AT --> " + new Date());

        try {
            // Read workbook
            ReadExcelParser excelReader = new ReadExcelParser(FILE_NAME, 10);
            List<ReadExcelParser.Row> data = excelReader.process();
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
}