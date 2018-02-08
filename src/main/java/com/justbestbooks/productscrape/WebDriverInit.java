/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.justbestbooks.productscrape;

import java.io.IOException;
import java.util.Properties;
import org.openqa.selenium.Capabilities;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.phantomjs.PhantomJSDriver;
import org.openqa.selenium.phantomjs.PhantomJSDriverService;
import org.openqa.selenium.remote.DesiredCapabilities;

/**
 *
 * @author itsga
 */
public class WebDriverInit {

    public WebDriver getDriver() {
        WebDriver driver = null;
        String browser = getConfigProperty("browser");
        switch (browser) {
            case "phantomjs":
                System.setProperty("phantomjs.binary.path", System.getProperty("user.dir") + getConfigProperty("phantomjsBrowserPath"));
                Capabilities caps = new DesiredCapabilities();
                ((DesiredCapabilities) caps).setJavascriptEnabled(true);
                ((DesiredCapabilities) caps).setCapability("takesScreenshot", true);
                ((DesiredCapabilities) caps).setCapability(
                        PhantomJSDriverService.PHANTOMJS_EXECUTABLE_PATH_PROPERTY,
                        System.getProperty("user.dir") + getConfigProperty("phantomjsBrowserPath")
                );
                String[] phantomArgs = new String[]{
                    "--webdriver-loglevel=NONE"
                };
                ((DesiredCapabilities) caps).setCapability(PhantomJSDriverService.PHANTOMJS_CLI_ARGS, phantomArgs);
                driver = new PhantomJSDriver(caps);
                break;
            case "chrome":
                System.setProperty("webdriver.chrome.driver", System.getProperty("user.dir") + "/driver/chromedriver.exe");
                driver = new ChromeDriver();
                break;
        }
        return driver;
    }

    public String getConfigProperty(String key) {
        Properties prop = new Properties();
        String value = null;
        try {
            prop.load(AmazonScrape.class.getClassLoader().getResourceAsStream("config.properties"));
            value = prop.getProperty(key);
        } catch (IOException io) {
            io.printStackTrace();
        }
        return value;
    }
}
