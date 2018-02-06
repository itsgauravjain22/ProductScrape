/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.justbestbooks.productscrape;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.WebDriverWait;

/**
 *
 * @author itsga
 */
public class FlipkartScrape {

    static String productBaseLocator = "//div[@id='container']/div/div[1]/div/div/div/div/div/div[2]";
    static String productTitleLocator = productBaseLocator + "/div[2]/div[1]/div/h1";
    static String productPriceLocator = productBaseLocator + "/div[2]/div[3]/div/div/div[2]";
    static String productSalePriceLocator = productBaseLocator + "/div[2]/div[3]/div/div/div[1]";
    static String productBookAuthor = "//span[text()='Author']/parent::div/parent::div/div[2]//a";
    static String productBookLanguage = "//li[contains(text(),'Language')]";
    static String productBookBinding = "//li[contains(text(),'Binding')]";
    static String productBookPublisher = "//li[contains(text(),'Publisher')]";
    static String productBookEdition = "//li[contains(text(),'Edition')]";
    static String productBookPages = "//li[contains(text(),'Pages')]";
    static String productBookImage = "//div[@id='container']/div/div[1]/div/div/div/div/div/div[1]/div/div/div[2]/div[2]/img";
    static String productBookDescription = "//span[text()='Description']/parent::div/parent::div/div[2]/div";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Map productMap = new HashMap();
        ExcelAdapter excelAdapter = new ExcelAdapter();
        for (String isbn : excelAdapter.getIsbnList()) {
            System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir") + "/driver/chromedriver.exe");
            WebDriver driver = new ChromeDriver();
            try {
                JavascriptExecutor js = (JavascriptExecutor) driver;
                driver.manage().window().maximize();
                driver.get("https://www.flipkart.com/search?q=" + isbn);
                WebElement element = driver.findElement(By.xpath("//div[@id='container']/div/div/div/div[2]/div/div[2]/div/div[3]/div/div[1]/div[1]/div/a"));
                element.click();
                for (String handle : driver.getWindowHandles()) {
                    driver.switchTo().window(handle);
                }
                js.executeScript("window.scrollTo(0, document.body.scrollHeight)");
                Thread.sleep(1000);
                js.executeScript("window.scrollTo(0, 0)");

                productMap.put("isbn", isbn);

                try {
                    element = driver.findElement(By.xpath(productTitleLocator));
                    productMap.put("title", ((String) element.getText().split("\\(")[0]).trim());
                } catch (Exception e) {

                }

                try {
                    element = driver.findElement(By.xpath(productPriceLocator));
                    productMap.put("regular_price", element.getText().replaceAll("[^0-9\\.]", "").trim());
                } catch (Exception e) {

                }

                try {
                    element = driver.findElement(By.xpath(productSalePriceLocator));
                    productMap.put("sale_price", element.getText().replaceAll("[^0-9\\.]", "").trim());
                } catch (Exception e) {

                }

                try {
                    element = driver.findElement(By.xpath(productBookAuthor));
                    productMap.put("author", element.getText().trim());
                } catch (Exception e) {

                }

                try {
                    element = driver.findElement(By.xpath(productBookLanguage));
                    productMap.put("language", ((String) element.getText().split(": ")[1]).trim());

                } catch (Exception e) {

                }

                try {
                    element = driver.findElement(By.xpath(productBookBinding));
                    productMap.put("binding", ((String) element.getText().split(": ")[1]).trim());
                } catch (Exception e) {

                }

                try {
                    element = driver.findElement(By.xpath(productBookPublisher));
                    productMap.put("publisher", ((String) element.getText().split(": ")[1]).trim());
                } catch (Exception e) {

                }

                try {
                    element = driver.findElement(By.xpath(productBookEdition));
                    productMap.put("edition", ((String) element.getText().split(": ")[1]).trim());
                } catch (Exception e) {

                }
                try {
                    element = driver.findElement(By.xpath(productBookPages));
                    productMap.put("numberofpages", ((String) element.getText().split(": ")[1]).trim());
                } catch (Exception e) {

                }

                try {
                    element = driver.findElement(By.xpath(productBookImage));
                    productMap.put("image", element.getAttribute("src").trim());
                } catch (Exception e) {

                }

                try {
                    element = driver.findElement(By.xpath(productBookDescription));
                    productMap.put("description", element.getAttribute("innerHTML").trim().replaceAll("\n", ""));
                } catch (Exception e) {

                }

                excelAdapter.updateExcelSheet(productMap);

            } catch (InterruptedException ex) {
                Logger.getLogger(FlipkartScrape.class.getName()).log(Level.SEVERE, null, ex);
            } finally {
                driver.quit();
            }
        }
    }
}
