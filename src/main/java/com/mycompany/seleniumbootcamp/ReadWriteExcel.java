/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.seleniumbootcamp;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author antony.raphy
 * Date last modified: October 18, 2017
 * Purpose: Read and write from excel(*xls, *xlsx) files
 */
public class ReadWriteExcel {

    public ArrayList<String> readExcel(String filePath, String fileName) throws IOException {
        //Create an arraylist object to return the searchwords in Excel
        ArrayList<String> searchWords = new ArrayList<>();
        //Create an object of File class to open excel file
        File file = new File(filePath);
        //Create an object of FileInputStream class to read excel file
        FileInputStream inputStream = new FileInputStream(file);
        //Find the file extension by splitting file name in substring  and getting only extension name
        String fileExtensionName = fileName.substring(fileName.indexOf("."));
        //Check condition if the file is xlsx file
        if (fileExtensionName.equals(".xlsx")) {
            //If it is xlsx file then create object of XSSFWorkbook class
            XSSFWorkbook inputWorkbook = new XSSFWorkbook(inputStream);
            //Read sheet inside the workbook by its name
            XSSFSheet inputSheet = inputWorkbook.getSheetAt(0);
            //Find number of rows in excel file
            int rowCount = inputSheet.getLastRowNum();
            //create a data formatter object to convert cell object to string
            DataFormatter df = new DataFormatter();
            //create a loop over all the rows of excel file to read it
            for (int i = 0; i <= rowCount; i++) {
                XSSFRow row = inputSheet.getRow(i);
                //Create a loop to read cell values in a row
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    //read cell value and store it in a string 
                    String value = df.formatCellValue(row.getCell(j));
                    //add string to arraylist
                    searchWords.add(value); 
                }
                System.out.println();
            } 
        } 
        //Check condition if the file is xls file
        else if (fileExtensionName.equals(".xls")) {
            //If it is xls file then create object of HSSFWorkbook class
            HSSFWorkbook inputWorkbook = new HSSFWorkbook(inputStream);
            //Read sheet inside the workbook by its name
            HSSFSheet inputSheet = inputWorkbook.getSheetAt(0);
            //Find number of rows in excel file
            int rowCount = inputSheet.getLastRowNum();
            //create a data formatter object to convert cell object to string
            DataFormatter df = new DataFormatter();
            //create a loop over all the rows of excel file to read it
            for (int i = 0; i <= rowCount; i++) {
                HSSFRow row = inputSheet.getRow(i);
                //Create a loop to read cell values in a row
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    //read cell value and store it in a string 
                    String value = df.formatCellValue(row.getCell(j));
                    //add string to arraylist
                    searchWords.add(value); 
                }
                System.out.println();
            } 
        }
    //Close input stream
    inputStream.close();
    return searchWords;
    }    
    
    public void writeExcel(String filePath, String fileName, String dataToWrite) throws IOException {
        //try block to see if file already exist and if yes append output string to last row 
        try{
            //Create an object of File class to open excel file
            File file = new File(filePath);
            //Create an object of FileInputStream class to read output excel file
            FileInputStream inputStream = new FileInputStream(file);
            //Find the file extension by splitting file name in substring  and getting only extension name
            String fileExtensionName = fileName.substring(fileName.indexOf("."));
             //Check condition if the file is xlsx file
            if (fileExtensionName.equals(".xlsx")) {
                ///If it is xls file then create object of HSSFWorkbook class
                XSSFWorkbook outputWorkbook = new XSSFWorkbook(inputStream);
                //Read sheet inside the workbook by its name
                XSSFSheet outputSheet = outputWorkbook.getSheetAt(0);
                //Find number of rows in excel file
                int rowCount = outputSheet.getLastRowNum();
                 if(rowCount == 0){
                    //create a new row and append it below the last row of the sheet
                    XSSFRow newRow = outputSheet.createRow(rowCount);
                    //Fill data in first cell of row
                    Cell cell = newRow.createCell(0);
                    //write the dataToWrite string in the first cell of new row
                    cell.setCellValue(dataToWrite);
                }
                else{
                    //create a new row and append it below the last row of the sheet
                    XSSFRow newRow = outputSheet.createRow(rowCount+1);
                    //Fill data in first cell of row
                    Cell cell = newRow.createCell(0);
                    //write the dataToWrite string in the first cell of new row
                    cell.setCellValue(dataToWrite);  
                 }
                //Close input stream
                inputStream.close();
                //write data in the excel file
                try ( //Create an object of FileOutputStream class to write data in excel file
                        FileOutputStream outputStream = new FileOutputStream(file)) {
                    //write data in the excel file
                    outputWorkbook.write(outputStream);
                    //close output stream
                    outputStream.close();
                }
            } 
            //Check condition if the file is xls file
            else if (fileExtensionName.equals(".xls")) {
                ///If it is xls file then create object of HSSFWorkbook class
                HSSFWorkbook outputWorkbook = new HSSFWorkbook(inputStream);
                //Read sheet inside the workbook by its name
                HSSFSheet outputSheet = outputWorkbook.getSheetAt(0);
                //Find number of rows in excel file
                int rowCount = outputSheet.getLastRowNum();
                if(rowCount == 0){
                //create a new row as the first row of the sheet
                HSSFRow newRow = outputSheet.createRow(rowCount);
                //Fill data in first cell of row
                Cell cell = newRow.createCell(0);
                //write the dataToWrite string in the first cell of new row
                cell.setCellValue(dataToWrite);
                }
                else{
                //create a new row and append it below the last row of the sheet
                HSSFRow newRow = outputSheet.createRow(rowCount+1);
                //Fill data in first cell of row
                Cell cell = newRow.createCell(0);
                //write the dataToWrite string in the first cell of new row
                cell.setCellValue(dataToWrite);
                }
                //Close input stream
                inputStream.close();
                //write data in the excel file
                try ( 
                    //Create an object of FileOutputStream class to write data in excel file
                    FileOutputStream outputStream = new FileOutputStream(file)) {
                    //write data in the excel file
                    outputWorkbook.write(outputStream);
                    //close output stream
                    outputStream.close();
                }
            }
        }catch (IOException e){
            e.printStackTrace();
         }
    }
}

