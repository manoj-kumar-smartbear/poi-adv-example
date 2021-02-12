package com.poi.example.application;

import com.poi.example.utility.WriteExcelParser;

import java.util.Date;

public class ExcelWriter {
    public static void main(String[] args) {
        final String FILE_NAME = "Resident_INFO.xlsx";

        long startTime = System.currentTimeMillis();
        System.out.println("EXCEL WRITE PROCESS ---> STARTED AT --> " + new Date());

        try {
            // Write Excel
            WriteExcelParser.buildFile(FILE_NAME);
        } catch(Error e) {
            e.printStackTrace();
        } catch(Exception ex) {
            ex.printStackTrace();
        } finally {
            long endTime = System.currentTimeMillis();
            long timeElapsed = endTime - startTime;
            System.out.println("EXCEL WRITE PROCESS ---> FINISHED AT --> " + new Date());
            System.out.println("TOTAL EXECUTION TIME ---> " + timeElapsed + " milliseconds");
        }
    }
}