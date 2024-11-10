package com.automation.code;

import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;

public class HandleMegaMenu {
    public static void main(String[] args) throws InterruptedException {
        WebDriver driver = new ChromeDriver();
        driver.get("https://www.flipkart.com/");
        driver.manage().window().maximize();
        Thread.sleep(5000);

        WebElement target = driver.findElement(By.xpath("//div[@class='_1ch8e_'][2]"));
        Actions act = new Actions(driver);
        act.moveToElement(target).perform();
        WebElement target1 = driver.findElement(By.xpath("//a[normalize-space()='Laptop and Desktop']"));
        act.moveToElement(target1).perform();
        WebElement target2 = driver.findElement(By.xpath("//a[@class ='_3490ry'][2]"));
        act.click(target2).perform();

    }

}
