package Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.openxml4j.util.ZipSecureFile;
import org.apache.poi.xssf.usermodel.*;
import org.openqa.selenium.By;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.support.ui.Select;

public class Pad {

    public String dateFormatter(String receivedDate) throws ParseException {
        DateFormat inputFormatter = new SimpleDateFormat("MM-dd-yyyy HH:mm:ss");
        Date date = inputFormatter.parse(receivedDate);
        
        DateFormat outputFormatter = new SimpleDateFormat("MM/dd/yyyy");
        String formattedDate = outputFormatter.format(date);
        return formattedDate;
    }

    public LinkedHashSet<String> readExcel() throws IOException {
        String filePath = "D:\\Users\\F3C3XRT\\Documents";
        String fileName = "ragi.xlsx";
        String sheetName = "Sheet1";

        File file = new File(filePath + "\\" + fileName);
        FileInputStream inputStream = new FileInputStream(file);
        ZipSecureFile.setMinInflateRatio(0);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        // Read sheet inside the workbook by its name
        XSSFSheet sheet = workbook.getSheet(sheetName);
        // Find number of rows in excel file
        LinkedHashSet<String> hash_Set = new LinkedHashSet<String>();
        int rowCount = sheet.getLastRowNum();
        int colCount = sheet.getRow(1).getLastCellNum();
        // Create a loop over all the rows of excel file to read it
        for (int r = 1; r <= rowCount; r++) {
          XSSFRow row = sheet.getRow(r);
          for (int c = 0; c < colCount; c++) {
            XSSFCell cell = row.getCell(c);
            String var = cell.getStringCellValue();
            // System.out.println(cell.getStringCellValue());
            hash_Set.add(var);
          }
        }
        System.out.println(hash_Set);
        workbook.close();
        return hash_Set;
    }

    public void writeExcel(int rowPicker, String addInfo, String cardAcceptorName, String amount, String date, String stan, String logo, String cardNumber,String claimNumber, String entryMode, String posInput, String ID) throws IOException {
    // public void writeExcel(int rowPicker, String addInfo, String cardAcceptorName, String amount, String date, String stan, String logo, String cardNumber,String claimNumber) throws IOException {

        FileInputStream fis = new FileInputStream("D:\\Users\\F3C3XRT\\Documents\\ragi.xlsx");
        ZipSecureFile.setMinInflateRatio(0);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        XSSFSheet sheet = workbook.getSheetAt(0);

        sheet.getRow(rowPicker).createCell(1).setCellValue(addInfo);
        sheet.getRow(rowPicker).createCell(2).setCellValue(cardAcceptorName);
        sheet.getRow(rowPicker).createCell(3).setCellValue(amount);
        sheet.getRow(rowPicker).createCell(4).setCellValue(date);
        sheet.getRow(rowPicker).createCell(5).setCellValue(stan);
        sheet.getRow(rowPicker).createCell(6).setCellValue(logo);
        sheet.getRow(rowPicker).createCell(7).setCellValue(cardNumber);
        sheet.getRow(rowPicker).createCell(8).setCellValue(claimNumber);
        sheet.getRow(rowPicker).createCell(9).setCellValue(entryMode);
        sheet.getRow(rowPicker).createCell(10).setCellValue(posInput);
        // sheet.getRow(rowPicker).createCell(11).setCellValue(expiry);
        sheet.getRow(rowPicker).createCell(11).setCellValue(ID);


        fis.close();
        FileOutputStream outputStream = new FileOutputStream("D:\\Users\\F3C3XRT\\Documents\\ragi.xlsx");
        workbook.write(outputStream);
        outputStream.close();
    }

    public void touchExcel(int rowPicker, String defaoultItem) throws IOException {
        // public void writeExcel(int rowPicker, String addInfo, String cardAcceptorName, String amount, String date, String stan, String logo, String cardNumber,String claimNumber) throws IOException {
    
        FileInputStream fis = new FileInputStream("D:\\Users\\F3C3XRT\\Documents\\ragi.xlsx");
        ZipSecureFile.setMinInflateRatio(0);
        XSSFWorkbook workbook = new XSSFWorkbook(fis);
        XSSFSheet sheet = workbook.getSheetAt(0);

        sheet.getRow(rowPicker).createCell(1).setCellValue(defaoultItem);

        fis.close();
        FileOutputStream outputStream = new FileOutputStream("D:\\Users\\F3C3XRT\\Documents\\ragi.xlsx");
        workbook.write(outputStream);
        outputStream.close();
    }

    public void switchingoo(WebDriver driver, String title) {
        Set<String> windowIds = driver.getWindowHandles();
        for (String currentWindow : windowIds) {
            driver.switchTo().window(currentWindow);
            String currentTitle = driver.getTitle();
            if (currentTitle.contentEquals(title)) {
                break;
            }
        }
        // System.out.println("Switching To: " + driver.getTitle());
    }

    public String claimFetcher(WebDriver driver, String homeWindowHandleID, String spc, String logo,String cardNumber, String stan) {
        driver.switchTo().window(homeWindowHandleID);
        driver.switchTo().frame("main");
        driver.findElement(By.xpath("//a[contains(text(),'Single Point Corrections')]")).click();
        // switching to cardManagement window
        switchingoo(driver, spc);
        driver.findElement(By.id("pmenuanchor")).click();
        driver.findElement(By.id("cjSearchLoad_initiatedByA")).click();
        Select drpTypeLogo = new Select(driver.findElement(By.name("clientSelectionOption")));/**Handling dropdown */
        drpTypeLogo.selectByVisibleText("Enter Logo");
        driver.findElement(By.id("cjSearchLoad_logoEnter")).clear();
        driver.findElement(By.id("cjSearchLoad_logoEnter")).sendKeys(logo);
        driver.findElement(By.id("cjSearchLoad_cardholderNum")).clear();
        driver.findElement(By.id("cjSearchLoad_cardholderNum")).sendKeys(cardNumber);
        driver.findElement(By.id("cjSearchLoad_cjSearch")).click();
        // List <WebElement> col = driver.findElements(By.xpath("//*[@id='fmForm:j_id833']/div/table/thead/tr/th"));
        // System.out.println("No of cols are : " +col.size()); 
        //No.of rows 
        List <WebElement> rows = driver.findElements(By.xpath("//*[contains(@id,'-4')]")); 
        // System.out.println("No of rows are : " + rows.size());

        String claimNumber = "Not Found";
        int poi;
        claimFetcher: for (poi = 0; poi < rows.size(); poi++) {
            String originalStan = driver.findElement(By.id("dt1-"+poi+"-4")).getText(); 
            if (originalStan.equals(stan)) {
                claimNumber = driver.findElement(By.id("dt1-"+poi+"-1")).getText();
                break claimFetcher;
            }
        }
        driver.close();
        return claimNumber;
        
    }

    public String[] posFetcher(WebDriver driver, String homeWindowHandleID, String spc, String logo, String cardNumber, String formattedDate, String stan) throws ParseException, IOException {
        driver.switchTo().window(homeWindowHandleID);
        driver.switchTo().frame("main");
        driver.findElement(By.xpath("//a[contains(text(),'Single Point Corrections')]")).click();
        // switching to SPC
        switchingoo(driver, spc);
        // driver.findElement(By.id("pmenuanchor")).click();
        Select drpTypeLogo = new Select(driver.findElement(By.name("clientSelectionOption")));/**Handling dropdown */
        drpTypeLogo.selectByVisibleText("Enter Logo");
        driver.findElement(By.id("welcome_logoEnter")).clear();
        driver.findElement(By.id("welcome_logoEnter")).sendKeys(logo);
        driver.findElement(By.id("welcome_cardNumber")).clear();
        driver.findElement(By.id("welcome_cardNumber")).sendKeys(cardNumber);
        // take month and selectDateRange
        driver.findElement(By.id("welcome_dateTimeSelectionOptionRANGE")).click();
        DateFormat inputFormatter = new SimpleDateFormat("MM/dd/yyyy");
        Date date = inputFormatter.parse(formattedDate);
        
        DateFormat outputFormatter = new SimpleDateFormat("MM");
        String month = outputFormatter.format(date);

        String entryMode = "Oops"; String posInput = "Oops";
        
        Select drpType = new Select(driver.findElement(By.name("selectDate")));/** Handling dropdown */
        switch (month) {
            case "03":
            drpType.selectByVisibleText("This month");
            break;
            case "02":
            drpType.selectByVisibleText("Last month");
            break;
            case "01":
            drpType.selectByVisibleText("2 months ago");
            break;
            case "12":
            drpType.selectByVisibleText("3 months ago");
            break;
            
            default: System.out.print("yavdu illappaa");
            driver.close();
            return new String[]{entryMode,posInput};
        }
        driver.findElement(By.id("welcome_sequenceNum")).clear();
        driver.findElement(By.id("welcome_sequenceNum")).sendKeys(stan);
        driver.findElement(By.id("welcome_switchTime")).click();
        driver.findElement(By.id("searchButton")).click();
        driver.findElement(By.id("buttonDetails")).click();
        entryMode = driver.findElement(By.id("loadDetails_tjTranDetailsForm_transactionEntryMode")).getAttribute("value");
        posInput = driver.findElement(By.id("loadDetails_tjAcquirerDetailsForm_termInputCapability")).getAttribute("value");
        driver.close();
        return new String[]{entryMode,posInput};

    }
    
    public String getExpiry(WebDriver driver,String homeWindowHandleID, String cardManagement, String logo, String cardNumber) {
        driver.switchTo().window(homeWindowHandleID);
        driver.switchTo().frame("main");
        driver.findElement(By.xpath("//a[contains(text(),'Card Management')]")).click();
        // switching to cardManagement window
        switchingoo(driver, cardManagement);
        driver.findElement(By.xpath("//label[contains(text(),'Enter Card Number')]")).click();
        driver.findElement(By.id("chooseClient_cardNumber")).sendKeys(cardNumber);
        driver.findElement(By.id("chooseClient_continue")).click();
        String eString = new String();
        List<WebElement> x = driver.findElements(By.xpath("//input[@id='viewCardDetails_cardExpirationDate']"));
        if (x.size() != 0) {
            eString = driver.findElement(By.xpath("//input[@id='viewCardDetails_cardExpirationDate']"))
                    .getAttribute("value");
        } else {
            driver.findElement(By.id("chooseCardHolder_chooseCardHolderSearch")).click();
            driver.findElement(By.id("submenu0")).click();
            eString = driver.findElement(By.id("loadChDetails_cardExpirationDate")).getAttribute("value");
        }

        driver.close();
        return eString;
        
    }
}
