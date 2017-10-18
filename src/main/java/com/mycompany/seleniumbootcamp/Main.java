/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.seleniumbootcamp;

import java.io.*;
import java.util.ArrayList;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;

/**
 *
 * @author antony.raphy
 * Date last modified: October 18, 2017
 * Purpose for main class: 
 * -Calls getSelecteFile() method twice defined in JFileChooser class which enables 
 *  the user to pick input and output excel files
 * -Stores the input values from input excel file chosen by the user to an array 
 *  list by calling readExcel() 
 *  method defined in ReadWriteExcel class
 * -Calls openPage() method defined in TraversePage class by passing the first 
 *  element of the array list
 * -Calls inputSearch() defined in TraversePage class  method by sequentially 
 *  passing second and third elements in the array list  
 * -Stores the error message from web site by calling errorMessageCheck() method
 *  defined in TraversePage class by passing the fourth element in the array list
 * -Calls writeExcel() method defined in ReadWriteExcel class to write the error 
 *  message from we site to the output excel file chosen by the user
 */
public class Main {

    public static void main(String args[]) throws InterruptedException, IOException {
        
        //Creating an object for REExcel class
        ReadWriteExcel RWExcel = new ReadWriteExcel();
        //Code to create a JFileChooser so that the user can select the file for.. 
        //..input by browsing through drives and clicking
        //create a JFileChooser object
        JFileChooser jfc = new JFileChooser(FileSystemView.getFileSystemView().getHomeDirectory());
        // create FileNameExtensionFilter object to filter excel files
        FileNameExtensionFilter excelFilter = new FileNameExtensionFilter("All Excel Files", "xls","xlsx");
        //set filter for jfc object
        jfc.setFileFilter(excelFilter);
        //use JFileChooser object to set the dialog box title
        jfc.setDialogTitle("Choose Input Excel File: ");
        int returnValue ;
        //call method to display dialogue box
        returnValue = jfc.showOpenDialog(null);
        //create file object to get the file user clicked
        File inputExcel = jfc.getSelectedFile();
        //if file selected print path and file name
	if (returnValue == JFileChooser.APPROVE_OPTION) {
            System.out.println(inputExcel.getAbsolutePath());
            System.out.println(inputExcel.getName());
	}
        //use JFileChooser object to set the dialog box title
        jfc.setDialogTitle("Choose Output Excel File: ");
        //call method to display dialogue box
        returnValue = jfc.showOpenDialog(null);
        //create file object to get the file user clicked
        File outputExcel = jfc.getSelectedFile();
        //if file selected print path and file name
	if (returnValue == JFileChooser.APPROVE_OPTION) {
            System.out.println(outputExcel.getAbsolutePath());
            System.out.println(outputExcel.getName());
	}
        //Create an arraylist to read data from ExcelSheet
        ArrayList<String> searchWords = new ArrayList<>();
        System.out.println("Size of arraylist before:"+searchWords.size());
        //Call readExcel method of the class to read data into an arraylist
        searchWords = RWExcel.readExcel(inputExcel.getAbsolutePath(), inputExcel.getName());
        System.out.println("Size of arraylist after:"+searchWords.size());
        //Creating an object for TraversePage class
        TraversePage TP = new TraversePage();
        //Calling function OpenPage()with url in first cell of first sheet
        TP.OpenPage(searchWords.get(0));
        //Calling function InputSearch()with second and third cells in the first sheet as strings
        TP.InputSearch(searchWords.get(1), searchWords.get(2));
        //Calling function ErrorMessageCheck()with fourth cell in the first sheet as string
        String websiteErrorMessage;
        websiteErrorMessage = TP.ErrorMessageCheck(searchWords.get(3));
        RWExcel.writeExcel(outputExcel.getAbsolutePath(), outputExcel.getName(), websiteErrorMessage);
    }
}
