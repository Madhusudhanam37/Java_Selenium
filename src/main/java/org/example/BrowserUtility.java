package org.example;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

public class BrowserUtility {

    WebDriver driver;
    String path="src/main/resources/driver/";
    public void initialzeChromeBrowser(){
        System.setProperty("webdriver.chrome.driver",path+"chromedriver.exe");
        driver=new ChromeDriver();
        driver.manage().window().maximize();
        driver.get("https://the-internet.herokuapp.com/");
        try {
            Thread.sleep(5000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        if (driver != null) {
            driver.quit();
        }

    }

    public void initialzeFireFoxBrowser(){

        System.setProperty("webdriver.gecko.driver",path+"geckodriver.exe");
        driver=new FirefoxDriver();
        driver.manage().window().maximize();
        driver.get("https://the-internet.herokuapp.com/");
        try {
            Thread.sleep(5000);
        } catch (InterruptedException e) {
            e.printStackTrace();
        }
        if (driver != null) {
            driver.quit();
        }
    }
}
