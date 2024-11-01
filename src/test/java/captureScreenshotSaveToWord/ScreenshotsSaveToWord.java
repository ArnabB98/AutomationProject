package captureScreenshotSaveToWord;

import org.apache.commons.io.FileUtils;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.ParagraphAlignment;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlException;
import org.openqa.selenium.By;
import org.openqa.selenium.OutputType;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;

public class ScreenshotsSaveToWord {
    static WebDriver driver;
    static String[] ScreenshotNames = new String[100];
    static int array_increment = 0;
    static String phoneNumber = "0423567889";

    public static void main(String[] args) throws IOException, XmlException, InterruptedException {
        driver = new FirefoxDriver();
        driver.get("https://www.wbsedcl.in/irj/go/km/docs/internet/new_website/Home.html");
        driver.manage().window().maximize();
       // driver.findElement(By.xpath("//div[@id='search_header']//input[@id='tipue_search_input']")).sendKeys(phoneNumber);
       // ScreenshotsSaveToWord.CaptureScreenshot(driver,"Checkpoint");
        ScreenshotsSaveToWord.CaptureScreenshot(driver, ScreenshotNames[array_increment++]="Homepage");
        driver.findElement(By.xpath("//div[@class = 'col-xs-9 col-md-9 paddleft0 home-menu-nav1']/ul/li[2]")).click();
        Thread.sleep(1000);
        driver.findElement(By.xpath("//ul[@id='aboutusSubmenu1']//li//a[@href='contactUs.html'][normalize-space()='Contact Us']")).click();
        ScreenshotsSaveToWord.CaptureScreenshot(driver, ScreenshotNames[array_increment++]="Contact Us");
       //save screenshot to word doc
        ScreenshotsSaveToWord.SaveScreenShotsT0WordDocument(phoneNumber + "_TestResult",ScreenshotNames);

    }

    public static void CaptureScreenshot(WebDriver driver,String screenshotName) throws IOException {
      //  File src=((TakesScreenshot)driver).getFullPageScreenshotAs(OutputType.FILE);
        File src = ((FirefoxDriver) driver).getFullPageScreenshotAs(OutputType.FILE);
        try {
            File dest =new File(System.getProperty("user.dir") + "\\TestResults\\" + screenshotName + ".jpg");
            FileUtils.copyFile(src, dest);
        }
        catch (IOException e){
            System.out.println(e.getMessage());
        }

//        // Take a full-page screenshot using Ashot
//        Screenshot screenshot = new AShot()
//                .shootingStrategy(ShootingStrategies.viewportPasting(1000))  // Scroll and capture the entire page
//                .takeScreenshot(driver);
//
//        // Save the screenshot
//        try {
//            ImageIO.write(screenshot.getImage(), "JPG", new File(new File("user.dir") + "\\TestResults\\" + screenshotName + ".jpg"));
//            System.out.println("Full page screenshot saved successfully.");
//        } catch (IOException e) {
//            System.out.println("Error saving screenshot: " + e.getMessage());
//        }
    }

    public static void SaveScreenShotsT0WordDocument (String documentName, String[] screenshotNames) throws IOException, XmlException {

        //Create Instance for document, paragraphs
        XWPFDocument doc = new XWPFDocument(); //Document Object
        XWPFParagraph p = doc.createParagraph(); //Paragraph alignments
        XWPFRun r = p.createRun(); //Set font styles, colors, next line

        //Title in the center ( Test Scenraio Name )
        p.setAlignment(ParagraphAlignment.CENTER);
        r.setBold(true);
        r.setFontFamily("Verdana");
        r.setText(documentName);
        r.addBreak();
        BufferedImage bimg1;

        for (String file : screenshotNames) {


            try {

                File dest = new File(System.getProperty("user.dir") + "\\TestResults\\" + file + ".jpg");
                bimg1 = ImageIO.read(dest);
                //BufferedImage bufferedImage = ImageIO.read(dest);
                //int width = Units.pixelToEMU(bufferedImage.getWidth());
                //int height = Units.pixelToEMU(bufferedImage.getHeight());

               int width = 500;
               int height = 280;

                String imgFile = dest.getName();
                int imgFormat = getImageFormat(imgFile);

                r.addBreak();
                r.addBreak();

                r.setText(file);
                r.addPicture(new FileInputStream(dest), imgFormat, imgFile, Units.toEMU(width), Units.toEMU(height));

            }
            catch(Exception e)
            {
                continue;
            }

            FileOutputStream out=new FileOutputStream(System.getProperty("user.dir")+ "\\TestResults\\" +  documentName + ".doc");
            doc.write(out);
            out.close();

        }
        System.out.println("Word document with screenshots created successfully");

    }

    public static int getImageFormat(String imgFileName)
    {
        int format;
        if (imgFileName.endsWith(".emf"))
            format = XWPFDocument.PICTURE_TYPE_EMF;
        else if (imgFileName.endsWith(".wmf"))
            format = XWPFDocument.PICTURE_TYPE_WMF;
        else if (imgFileName.endsWith(".pict"))
            format = XWPFDocument.PICTURE_TYPE_PICT;
        else if (imgFileName.endsWith(".jpeg") || imgFileName.endsWith(".jpg"))
            format = XWPFDocument.PICTURE_TYPE_JPEG;
        else if (imgFileName.endsWith(".png"))
            format = XWPFDocument.PICTURE_TYPE_PNG;
        else if (imgFileName.endsWith(".dib"))
            format = XWPFDocument.PICTURE_TYPE_DIB;
        else if (imgFileName.endsWith(".gif"))
            format = XWPFDocument.PICTURE_TYPE_GIF;
        else if (imgFileName.endsWith(".tiff"))
            format = XWPFDocument.PICTURE_TYPE_TIFF;
        else if (imgFileName.endsWith(".eps"))
            format = XWPFDocument.PICTURE_TYPE_EPS;
        else if (imgFileName.endsWith(".bmp"))
            format = XWPFDocument.PICTURE_TYPE_BMP;
        else if (imgFileName.endsWith(".wpg"))
            format = XWPFDocument.PICTURE_TYPE_WPG;
        else {
            return 0;
        }
        return format;
    }

}
