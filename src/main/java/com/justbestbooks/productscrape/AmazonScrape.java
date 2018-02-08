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
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;

/**
 *
 * @author itsga
 */
public class AmazonScrape {

    static String productBaseLocator = "//div[@id='container']/div/div[1]/div/div/div/div/div/div[2]";
    static String productTitleLocator = "//span[@id='productTitle']";
    static String productPriceLocator = "//span[contains(text(),'M.R.P.')]//parent::*/span[2]";
    static String productSalePriceLocator = "//span[contains(@class,'offer-price')]";
    static String productBookAuthor = "//span[contains(@class,'author')]/span[1]/a[1]";
    static String productBookLanguage = "//*[contains(@id,'detail_bullets')]/table//div[contains(@class,'content')]/ul/li/b[contains(text(),'Language')]/parent::li";
    static String productBookBinding = "//*[contains(@id,'detail_bullets')]/table//div[contains(@class,'content')]/ul/li[1]/b[1]";
    static String productBookPublisher = "//*[contains(@id,'detail_bullets')]/table//div[contains(@class,'content')]/ul/li/b[contains(text(),'Publisher')]/parent::li";
    static String productBookImage = "//div[@id='container']/div/div[1]/div/div/div/div/div/div[1]/div/div/div[2]/div[2]/img";
    static String productBookDescription = "//div[@data-feature-name='productDescription']//div[@id='productDescription']";

    public static void main(String[] args) throws IOException, InvalidFormatException {
        Map productMap = new HashMap();
        ExcelAdapter excelAdapter = new ExcelAdapter("amazon");
        for (String isbn : excelAdapter.getIsbnList()) {
            WebDriverInit webDriverInit = new WebDriverInit();
            WebDriver driver = webDriverInit.getDriver();
            try {
                JavascriptExecutor js = (JavascriptExecutor) driver;
                driver.manage().window().maximize();
                driver.get("https://www.amazon.in/s?url=field-keywords=" + isbn);
                WebElement element = driver.findElement(By.xpath("(//a[contains(@class,'s-access-detail-page')])[1]"));
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
                    productMap.put("title", ((String) element.getText()).trim());
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
                    productMap.put("language", ((String) element.getText().split(":")[1]).trim());

                } catch (Exception e) {

                }

                try {
                    element = driver.findElement(By.xpath(productBookBinding));
                    productMap.put("binding", ((String) element.getText().split(":")[0]).trim());
                } catch (Exception e) {

                }

                try {
                    element = driver.findElement(By.xpath(productBookPublisher));
                    productMap.put("publisher", (element.getText().split(":")[1]).split(";")[0].trim());
                } catch (Exception e) {

                }

                try {
                    element = driver.findElement(By.xpath(productBookPublisher));
                    productMap.put("edition", (element.getText().split(":")[1]).split(";")[1].trim());
                } catch (Exception e) {

                }
                try {
                    element = driver.findElement(By.xpath("//*[contains(@id,'detail_bullets')]/table//div[contains(@class,'content')]/ul/li/b[contains(text(),'"+productMap.get("binding")+"')]/parent::li"));
                    productMap.put("numberofpages", element.getText().split(":")[0].trim());
                } catch (Exception e) {

                }

                try {
                    element = driver.findElement(By.xpath(productBookImage));
                    productMap.put("image", element.getAttribute("src").trim());
                } catch (Exception e) {

                }

                try {
                    element = driver.findElement(By.xpath(productBookDescription));
                    productMap.put("description", element.getAttribute("innerHTML").trim());
                } catch (Exception e) {

                }
                excelAdapter.updateExcelSheet(productMap);
            } catch (InterruptedException ex) {
                Logger.getLogger(AmazonScrape.class.getName()).log(Level.SEVERE, null, ex);
            } finally {
                driver.quit();
            }
        }
    }
}
