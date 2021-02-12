package com.poi.example.application;

import com.poi.example.utility.ReadExcelParser;

import java.util.Date;
import java.util.List;

public class ExcelReader {
    public static void main(String[] args) {
        final String FILE_NAME = "Resident_INFO.xlsx";

        long startTime = System.currentTimeMillis();
        System.out.println("EXCEL READ PROCESS ---> STARTED AT --> " + new Date());

        try {
            // Read Excel
            ReadExcelParser excelReader = new ReadExcelParser(FILE_NAME, 10);
            List<ReadExcelParser.Row> data = excelReader.process();
            System.out.println("TOTAL ROWS READ --> " + data.size());
        } catch(Error e) {
            e.printStackTrace();
        } catch(Exception ex) {
            ex.printStackTrace();
        } finally {
            long endTime = System.currentTimeMillis();
            long timeElapsed = endTime - startTime;
            System.out.println("EXCEL READ PROCESS ---> FINISHED AT --> " + new Date());
            System.out.println("TOTAL EXECUTION TIME --> " + timeElapsed + " milliseconds");
        }
    }
}