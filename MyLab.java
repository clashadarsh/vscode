package casemgmt;


import java.util.Date;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;

import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Part;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMultipart;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
// import org.openqa.selenium.winium.DesktopOptions;
// import org.openqa.selenium.winium.WiniumDriver;
import org.openqa.selenium.interactions.Actions;

public class MyLab {
    
    public void createAttachment(Message mail, String caseNumber, String subFinal) throws IOException, MessagingException, InterruptedException {
        File file = new File("D:\\Users\\F3C3XRT\\Documents", caseNumber+".eml");
        file.getParentFile().mkdirs();
        file.createNewFile();
        FileOutputStream outputStream = new FileOutputStream(file, false);
        mail.writeTo(outputStream);
        outputStream.close();

        // File file1 = new File("D:\\Users\\F3C3XRT\\Documents", caseNumber+".doc");
        // file1.getParentFile().mkdirs();
        // file1.createNewFile();
        // FileOutputStream outputStream1 = new FileOutputStream(file1, false);
        // mail.writeTo(outputStream1);
        // outputStream1.close();


        // DesktopOptions option = new DesktopOptions();
        // option.setApplicationPath("C:\\Program Files\\Microsoft Office\\root\\Office16\\OUTLOOK.exe");
        // // Runtime.getRuntime().exec("D:\\Winium.Desktop.Driver.exe", null, new File("D:\\"));
        // WiniumDriver windriver = new WiniumDriver(new URL("http://localhost:9999"), option);
        // // windriver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        // Thread.sleep(2000);
        // windriver.findElement(By.name("Unique")).click();
        // Thread.sleep(1000);
        // windriver.findElement(By.name("File Tab")).click();
        // Thread.sleep(1000);
        // windriver.findElement(By.name("Save As")).click();
        // Thread.sleep(1000);
        // windriver.findElement(By.name("File name:")).sendKeys("hahah");
        // Thread.sleep(1000);
        // windriver.findElement(By.name("Save")).click();
        // windriver.findElement(By.name("Close")).click();
       


        String filePath = "D:\\Users\\F3C3XRT\\Documents\\"+caseNumber+".eml";
        String zipPath = "D:\\Users\\F3C3XRT\\Documents\\"+caseNumber+".zip";
        try (ZipOutputStream zipOut = new ZipOutputStream(new FileOutputStream(zipPath))) {
        File fileToZip = new File(filePath);
        zipOut.putNextEntry(new ZipEntry(fileToZip.getName()));
        Files.copy(fileToZip.toPath(), zipOut);
        }
    }

    public static void switchingoo(WebDriver driver, String title) {
        Set<String> windowIds = driver.getWindowHandles();
        for (String currentWindow : windowIds) {
            driver.switchTo().window(currentWindow);
            String currentTitle = driver.getTitle();
            if (currentTitle.contentEquals(title)) {
                break;
            }
        }
    }
    
    public String dateFormatter(Date receivedDate) throws ParseException {
        String input = receivedDate.toString();
        DateFormat inputFormatter = new SimpleDateFormat("EEE MMM dd HH:mm:ss zzzz yyyy");
        Date date = inputFormatter.parse(input);
        
        DateFormat outputFormatter = new SimpleDateFormat("MM/dd/yyyy");
        String formattedDate = outputFormatter.format(date);
        return formattedDate;
    }

    public boolean hasAttachments(Message email) throws Exception {

        // suppose 'message' is an object of type Message
        String contentType = email.getContentType();
        System.out.println(contentType);

        if (contentType.toLowerCase().contains("multipart/mixed")) {
            // this message must contain attachment
            Multipart multiPart = (Multipart) email.getContent();

            for (int i = 0; i < multiPart.getCount(); i++) {
                MimeBodyPart part = (MimeBodyPart) multiPart.getBodyPart(i);
                if (Part.ATTACHMENT.equalsIgnoreCase(part.getDisposition())) {
                    System.out.println("Attached filename is:" + part.getFileName());

                    MimeBodyPart mimeBodyPart = (MimeBodyPart) part;
                    String fileName = mimeBodyPart.getFileName();

                    String destFilePath = System.getProperty("user.dir") + "\\Resources\\";

                    File fileToSave = new File(fileName);
                    mimeBodyPart.saveFile(destFilePath + fileToSave);

                    // download the pdf file in the resource folder to be read by PDFUTIL api.

                    Path fileToDeletePath = Paths.get(destFilePath + fileToSave);
                    Files.delete(fileToDeletePath);
                }
            }

            return true;
        }

        return false;
    }

    public String getTextFromMessage(Message message) throws MessagingException, IOException {
        String result = "";
        if (message.isMimeType("text/plain")) {
            result = message.getContent().toString();
        } else if (message.isMimeType("multipart/*")) {
            MimeMultipart mimeMultipart = (MimeMultipart) message.getContent();
            result = getTextFromMimeMultipart(mimeMultipart);
        }
        return result;
    }
    
    public static String getTextFromMimeMultipart(
            MimeMultipart mimeMultipart)  throws MessagingException, IOException{
        String result = "";
        int count = mimeMultipart.getCount();
        for (int i = 0; i < count; i++) {
            BodyPart bodyPart = mimeMultipart.getBodyPart(i);
            if (bodyPart.isMimeType("text/plain")) {
                result =  "\n" + bodyPart.getContent();
                break; // without break same text appears twice in my tests
            } else if (bodyPart.isMimeType("text/html")) {
                String html = (String) bodyPart.getContent();
                result = result + org.jsoup.Jsoup.parse(html).toString();
            } else if (bodyPart.getContent() instanceof MimeMultipart){
                result = result + getTextFromMimeMultipart((MimeMultipart)bodyPart.getContent());
            }
        }
        return result;
    }

    public String cardRemover(String text) {
        // String text = "Testing card number one 4444444444444449 second card number 5444444444444448 Sent from my iPhone";
        Pattern p = Pattern.compile("(?:[0-9]{13}(?:[0-9]{3})?|5[1-5][0-9]{14})");
        // get a matcher object
        // String v = "my visa 4444444444444448 number 4444444444444448";
        
        Matcher match = p.matcher(text);
        System.out.println("match");
        int count = 0;
        while (match.find()) {
          System.out.println("Match number " + count);
          System.out.println("start(): " + match.start());
          System.out.println("end(): " + match.end());
          text = text.replaceAll("(?:[0-9]{13}(?:[0-9]{3})?|5[1-5][0-9]{14})", "XXXXxxxxxxxxXXXX");
          
          count++;
          // count=match.end();
        }
        System.out.println(text);
        return text;
        
    }

    public String prefixSelector(String clientDeployment) {
        String prefix = null;
        if (clientDeployment.contains("Accel")) {
            prefix = "Accel: ";
        } else if (clientDeployment.contains("STAR")){
            prefix = "STAR: ";
        } else if (clientDeployment.contains("Money")){
            prefix = "Moneypass: ";
        } else if (clientDeployment.contains("POS")){
            prefix = "POS: ";
        }
        return prefix;
    }

    public void assignmentGroupSelector(String prefix, WebDriver driver) throws InterruptedException {
        if (prefix.contains("POS")){
            driver.findElement(By.id("sys_display.sn_customerservice_case.assignment_group")).sendKeys("Card-Client Services POS-Deb");
            Thread.sleep(2000);
            Actions act3 = new Actions(driver);
            act3.sendKeys(Keys.DOWN).perform();
            act3.sendKeys(Keys.ENTER).perform();
        } else {
            driver.findElement(By.id("sys_display.sn_customerservice_case.assignment_group")).sendKeys("CARD.1.Network Member Relati");
            Thread.sleep(2000);
            Actions act3 = new Actions(driver);
            act3.sendKeys(Keys.DOWN).perform();
            act3.sendKeys(Keys.ENTER).perform();
        }
    }

    public String[] readExcel(int rowPicker) throws IOException {
        String filePath = "D:\\Users\\F3C3XRT\\Documents";
        String fileName = "Domain & Company Name - DB.xlsx";
        String sheetName = "Sheet1";
    
        File file = new File(filePath + "\\" + fileName);
        FileInputStream inputStream = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        // Read sheet inside the workbook by its name
        XSSFSheet sheet = workbook.getSheet(sheetName);
        // Find number of rows in excel file
        int rowCount = sheet.getLastRowNum();
        String rows =Integer.toString(rowCount);
        // System.out.println(rowCount);
    
        String domain = sheet.getRow(rowPicker).getCell(0).getStringCellValue();
        String company = sheet.getRow(rowPicker).getCell(1).getStringCellValue();
        
        workbook.close();
        return new String[]{domain , company, rows};
    }
    
}
