package captureScreenshotSaveToWord;

import org.apache.commons.io.FileUtils;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.*;
import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

public class MultiScreenshotsSaveToWordFromExcel {
    static WebDriver driver;
    static String[] ScreenshotNames = new String[100];
    static int array_increment = 0;

    public static void main(String[] args) throws IOException, InterruptedException {
        String excelFilePath = System.getProperty("user.dir") + "/TestData/phoneNumbers.xlsx";
        List<String> phoneNumbers = readPhoneNumbersFromExcel(excelFilePath);

        driver = new FirefoxDriver();
        driver.manage().window().maximize();

        for (String phoneNumber : phoneNumbers) {
            driver.get("https://www.wbsedcl.in/irj/go/km/docs/internet/new_website/Home.html");

            // Reset screenshot names array for each test
            ScreenshotNames = new String[100];
            array_increment = 0;

            // Perform steps and capture screenshots for each scenario
            driver.findElement(By.xpath("//input[@class='search-box d-none d-sm-block']")).sendKeys(phoneNumber);
            MultiScreenshotsSaveToWordFromExcel.CaptureScreenshot(driver, ScreenshotNames[array_increment++] = "Homepage_" + phoneNumber);
            driver.findElement(By.xpath("//div[@class = 'col-xs-9 col-md-9 paddleft0 home-menu-nav1']/ul/li[2]")).click();
            Thread.sleep(1000);
            driver.findElement(By.xpath("//ul[@id='aboutusSubmenu1']//li//a[@href='contactUs.html'][normalize-space()='Contact Us']")).click();
            MultiScreenshotsSaveToWordFromExcel.CaptureScreenshot(driver, ScreenshotNames[array_increment++] = "Contact Us_" + phoneNumber);

            // Save screenshots to a Word document for each scenario
            MultiScreenshotsSaveToWordFromExcel.SaveScreenShotsT0WordDocument(phoneNumber + "_TestResult", ScreenshotNames);
        }

        driver.quit();
    }

    // Method to read phone numbers from Excel file
    public static List<String> readPhoneNumbersFromExcel(String excelFilePath) {
        List<String> phoneNumbers = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(new File(excelFilePath));
             Workbook workbook = new XSSFWorkbook(fis)) {

            Sheet sheet = workbook.getSheetAt(0); // Get the first sheet
            Iterator<Row> rowIterator = sheet.iterator();

            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Cell cell = row.getCell(0); // Assuming phone numbers are in the first column

                if (cell != null) {
                    phoneNumbers.add(cell.toString().trim());
                }
            }
        } catch (IOException e) {
            System.out.println("Error reading Excel file: " + e.getMessage());
        }

        return phoneNumbers;
    }

    public static void CaptureScreenshot(WebDriver driver, String screenshotName) throws IOException {
        File src = ((FirefoxDriver) driver).getFullPageScreenshotAs(OutputType.FILE);
        String screenshotPath = System.getProperty("user.dir") + "\\Excel_TestResults\\" + screenshotName + ".jpg";
        File dest = new File(screenshotPath);

        try {
            FileUtils.copyFile(src, dest);
            if (dest.exists()) {
                System.out.println("Screenshot saved: " + dest.getAbsolutePath());
            } else {
                System.out.println("Screenshot failed to save: " + dest.getAbsolutePath());
            }
        } catch (IOException e) {
            System.out.println("Error saving screenshot: " + e.getMessage());
        }
    }

    public static void SaveScreenShotsT0WordDocument(String documentName, String[] screenshotNames) throws IOException {
        XWPFDocument doc = new XWPFDocument();
        XWPFParagraph p = doc.createParagraph();
        XWPFRun r = p.createRun();
        p.setAlignment(ParagraphAlignment.CENTER);
        r.setBold(true);
        r.setFontFamily("Verdana");
        r.setText(documentName);
        r.addBreak();

        String screenshotFolderPath = System.getProperty("user.dir") + "\\Excel_TestResults\\";

        for (String file : screenshotNames) {
            if (file == null) continue; // Skip null entries

            try {
                File dest = new File(screenshotFolderPath + file + ".jpg");
                if (!dest.exists()) {
                    System.out.println("File not found: " + dest.getAbsolutePath());
                    continue;
                }

                BufferedImage bimg1 = ImageIO.read(dest);
                int width = 500;
                int height = 280;

                String imgFile = dest.getName();
                int imgFormat = getImageFormat(imgFile);

                r.addBreak();
                r.addBreak();
                r.setText(file);
                r.addPicture(new FileInputStream(dest), imgFormat, imgFile, Units.toEMU(width), Units.toEMU(height));
            } catch (Exception e) {
                System.out.println("Error adding image: " + file + " - " + e.getMessage());
                continue;
            }
        }

        try (FileOutputStream out = new FileOutputStream(screenshotFolderPath + documentName + ".doc")) {
            doc.write(out);
        }
        System.out.println("Word document with screenshots created successfully: " + documentName);
    }

    public static int getImageFormat(String imgFileName) {
        if (imgFileName.endsWith(".emf")) return XWPFDocument.PICTURE_TYPE_EMF;
        if (imgFileName.endsWith(".wmf")) return XWPFDocument.PICTURE_TYPE_WMF;
        if (imgFileName.endsWith(".pict")) return XWPFDocument.PICTURE_TYPE_PICT;
        if (imgFileName.endsWith(".jpeg") || imgFileName.endsWith(".jpg")) return XWPFDocument.PICTURE_TYPE_JPEG;
        if (imgFileName.endsWith(".png")) return XWPFDocument.PICTURE_TYPE_PNG;
        if (imgFileName.endsWith(".dib")) return XWPFDocument.PICTURE_TYPE_DIB;
        if (imgFileName.endsWith(".gif")) return XWPFDocument.PICTURE_TYPE_GIF;
        if (imgFileName.endsWith(".tiff")) return XWPFDocument.PICTURE_TYPE_TIFF;
        if (imgFileName.endsWith(".eps")) return XWPFDocument.PICTURE_TYPE_EPS;
        if (imgFileName.endsWith(".bmp")) return XWPFDocument.PICTURE_TYPE_BMP;
        if (imgFileName.endsWith(".wpg")) return XWPFDocument.PICTURE_TYPE_WPG;
        return 0;
    }
}

