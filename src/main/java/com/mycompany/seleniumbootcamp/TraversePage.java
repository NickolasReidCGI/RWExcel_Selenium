/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.mycompany.seleniumbootcamp;

import java.io.*;
import java.util.List;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;

/**
 *
 * @author antony.raphy
 * Date last modified: October 18, 2017
 * class to open a web page, enters a search word and picks the first article in
 * search results, search for another search word which would likely return an error
 * message, compares the error message with a pre-defined error message.
 */
//class to traverse through webpages
public class TraversePage {

    //create a web driver object
    WebDriver driver;

    //constructor for TraverPage class
    public TraversePage() {
        // create file object to store chrome binary file
        File chromeBinary = new File("chromedriver.exe");
        //get absolute path for chrome binary file
        System.setProperty("webdriver.chrome.driver", chromeBinary.getAbsolutePath());
        //create a new driver object for chrome driver
        driver = new ChromeDriver();
    }

    // method to open a web page in chrome browser
    public void OpenPage(String url) throws InterruptedException {
        //opens a url
        driver.get(url);
    }

    //method that imputs two words to search
    public void InputSearch(String searchWord1, String searchWord2) throws InterruptedException {

        WebElement searchWordWE = driver.findElement(By.xpath(".//input[@id='orb-search-q']"));

        searchWordWE.sendKeys(searchWord1);
        WebElement submitWE = driver.findElement(By.xpath(".//button[@class='se-searchbox__submit']"));
        submitWE.click();
        //store first article in a web element object
        WebElement firstArticle = driver.findElement(By.xpath("(.//article)[1]"));
        //check if a display date element exist in the article attribute
        try {
            firstArticle.findElement(By.xpath(".//time[@class='display-date'][1]"));
            //create WebELement for the article heading
            WebElement articleHeading = firstArticle.findElement(By.xpath(".//h1[@itemprop='headline']"));
            //click on the article heading
            articleHeading.click();
        } catch (Exception e) {   //catch the exception if the date element is not in the article attribute
            //store second article in a web element object
            WebElement secondArticle = driver.findElement(By.xpath("(.//article)[2]"));
            //check if a display date element exist in the article web element
            if (secondArticle.findElement(By.xpath(".//time[@class='display-date'][1]")) != null) {
                //create web element for the article heading
                WebElement articleHeading = secondArticle.findElement(By.xpath(".//h1[@itemprop='headline']"));
                //click on the article heading
                articleHeading.click();
            }
        }

        //
        searchWordWE = driver.findElement(By.xpath(".//input[@id='orb-search-q']"));
        searchWordWE.sendKeys(searchWord2);
        submitWE = driver.findElement(By.xpath(".//button[@class='se-searchbox__submit']"));
        submitWE.click();
    }

    public String ErrorMessageCheck(String storedErrorMessage) throws InterruptedException {
        WebElement errorMessageWE = driver.findElement(By.xpath("(.//p[@tabindex='0'])[1]"));
        String websiteErrorMessage;
        websiteErrorMessage = errorMessageWE.getText();
        System.out.println("Error Message: " + storedErrorMessage);
        if (websiteErrorMessage.equals(storedErrorMessage)) {
            System.out.println("Error messages match");
            System.out.println("Error Message: " + websiteErrorMessage);
        } else {
            System.out.println("Error messages do not match");
            System.out.println("Error Message: " + websiteErrorMessage);
        }
        driver.close();
        return websiteErrorMessage;
    }
}
