package Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedHashSet;
import java.util.Set;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

public class Bull {
    
    public static void main(String[] args) throws InterruptedException, IOException {
        String disputeManager = "Dispute Manager [www.client-central.com]";
        // String systemNotReady = "System Not Ready [www.client-central.com]";
        String disputeDetail = "Repository Viewer - Dispute Detail [www.client-central.com]";
        String cardManagement = "CWSi | Card Management - Choose Client ";
        String disputeActionCreate = "Dispute Action Create [www.client-central.com]";
        String disputeSearch = "Repository Viewer - Dispute Search [www.client-central.com]";
        String filePath = "C:\\Users\\f3c3xrt\\Music";
        String fileName = "Book.xlsx";
        String sheetName = "Sheet1";
        System.setProperty("webdriver.chrome.driver" , "C:\\Users\\f3c3xrt\\Music\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.get("https://www.client-central.com");
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        driver.findElement(By.id("Ecom_User_ID")).sendKeys("amurth00");
        driver.manage().window().maximize();
        driver.findElement(By.id("Ecom_Token")).click();
        Thread.sleep(12000);
        driver.findElement(By.id("loginButton2")).click();
        //switching to iframe
        String homeWindowHandleID = driver.getWindowHandle();
        // String homeWindowTitle = driver.getTitle();    Switchingoo nu use madbodu
        System.out.println("Parent window: "+homeWindowHandleID);
        driver.switchTo().frame("main");
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        System.out.println("We are switching to the iframe");
        driver.findElement(By.xpath("//a[contains(text(),'Dispute Manager')]")).click();
        //switching to Dispute Manager
        switchingoo(driver,disputeManager);
        if (driver.getTitle().contains("Dispute Manager [www.client-central.com]")) {
            driver.findElement(By.xpath("//*[@id='fmForm:j_id103:1:j_id108']")).click();/**Dispute */
            Thread.sleep(1000);
            driver.findElement(By.id("fmForm:j_id124:0:j_id137")).click();/**Rep Viewer */

            File file =    new File(filePath+"\\"+fileName);
            FileInputStream inputStream = new FileInputStream(file);
            XSSFWorkbook workbook= new XSSFWorkbook(inputStream);
            //Read sheet inside the workbook by its name
            XSSFSheet sheet = workbook.getSheet(sheetName);
            //Find number of rows in excel file
            LinkedHashSet<String> hash_Set = new LinkedHashSet<String>();
            int rowCount = sheet.getLastRowNum();
            int colCount = sheet.getRow(1).getLastCellNum();
            //Create a loop over all the rows of excel file to read it
            for (int r = 0; r <= rowCount; r++) {
                XSSFRow row = sheet.getRow(r);
                for (int c = 0; c <colCount; c++) {
                    XSSFCell cell = row.getCell(c);
                    String var = cell.getStringCellValue();
                    // System.out.println(cell.getStringCellValue());  
                    hash_Set.add(var); 
                }
            } 
            System.out.println(hash_Set);
            workbook.close();

            for(String iter: hash_Set){
                System.out.println(iter);
                driver.findElement(By.xpath("//*[@id='fmForm:evalId']")).clear();
                driver.findElement(By.xpath("//*[@id='fmForm:evalId']")).sendKeys(iter);
                Select drpType = new Select(driver.findElement(By.name("fmForm:actnTypeByNameFndrElem")));/**Handling dropdown */
                drpType.selectByVisibleText("Accel Dispute Inquiry");
                driver.findElement(By.xpath("//*[@id='fmForm:btnScan']")).click();
                driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);
                // bringing in javascript since element is hidden
                WebElement searchedDispute = driver.findElement(By.xpath("//*[@id='fmForm:rptEval:0:j_id854']"));
                JavascriptExecutor js = (JavascriptExecutor)driver;
                js.executeScript("arguments[0].click()", searchedDispute);
                Thread.sleep(3000);
                // switching to Repository Viewer;
                switchingoo(driver, disputeDetail);
                String cardNumber = driver.findElement(By.xpath("//*[@id='fmForm:j_id249']/td[3]/span")).getText();
                System.out.println(cardNumber);
                driver.findElement(By.id("fmForm:btnActn")).click();
                driver.findElement(By.id("fmForm:btnmActn:evalInqySitnDtrm")).click();
                Actions actions = new Actions(driver);
                actions.contextClick(driver.findElement(By.xpath("//*[@id='fmForm:j_id1160']/table/tbody/tr[2]/td[1] | //*[@id='fmForm:j_id1203']/table/tbody/tr[2]/td[1]"))).perform(); /**id id damn dynamic*/
                driver.findElement(By.id("fmForm:evalActnMgrMpnlCmnuActnType")).click();
                driver.close();
                
                // swithing to disputeActionCreate
                switchingoo(driver, disputeActionCreate);
                Select drpType2 = new Select(driver.findElement(By.id("fmForm:ActyDtl_FisvMsgResnCode_elem")));
                drpType2.selectByVisibleText("0081 - Chargeback - Fraud, unauthorized");
                Select drpType3 = new Select(driver.findElement(By.name("fmForm:ActyDtl_FisvFrdType_elem")));
                drpType3.selectByVisibleText("6 - Fraudulent Use");
                // switching to parent window >>cardManagemt
                driver.switchTo().window(homeWindowHandleID);
                driver.switchTo().frame("main");
                driver.findElement(By.xpath("//a[contains(text(),'Card Management')]")).click();
                // switching to cardManagement window
                switchingoo(driver, cardManagement);
                driver.findElement(By.xpath("//label[contains(text(),'Enter Card Number')]")).click();
                driver.findElement(By.id("chooseClient_cardNumber")).sendKeys(cardNumber);
                driver.findElement(By.id("chooseClient_continue")).click();
                String estring = driver.findElement(By.xpath("//input[@id='viewCardDetails_cardExpirationDate']")).getAttribute("value");
                System.out.println(estring);  
                char[] ch = new char[4];
                ch[0] = estring.charAt(3);
                ch[1] = estring.charAt(4);
                ch[2] = estring.charAt(0);
                ch[3] = estring.charAt(1);
                String date = String.valueOf(ch);
                System.out.println(date);
                driver.close(); /*CLOSE CARD MANAGEMENT*/
                switchingoo(driver, disputeActionCreate);
                driver.findElement(By.id("fmForm:ActyDtl_DateExp_elem")).sendKeys(date);
                driver.findElement(By.id("fmForm:ActyDtl_FisvMbrMsgText_elem")).sendKeys("Cardholder did not authorize transaction");
                driver.findElement(By.id("fmForm:ActyDtl_FisvAddlCmnt_elem")).sendKeys("Fraudulent use for online transactions. Cardholder did not participate in the transaction");
                Select drpType4 = new Select(driver.findElement(By.id("fmForm:ActyDtl_FisvDocInd_elem")));
                drpType4.selectByVisibleText("0 - No document");
                driver.findElement(By.id("fmForm:j_id611")).click();
                driver.findElement(By.id("fmForm:btnHeaderAdd")).click();
                driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
                Select drpType5 = new Select(driver.findElement(By.id("fmForm:evalResnByNamePckrElem")));
                drpType5.selectByVisibleText("*Chargeback Submitted");
                driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);/**thread.seleep remove madidini */
                driver.findElement(By.id("fmForm:btnUpd")).click();
                Thread.sleep(2000);
                System.out.println("Chargeback Submitted: " + iter);
                driver.close(); /*CLOSE CHARGEBACK FORM*/   

                switchingoo(driver, disputeSearch);
                                
            }

            JavascriptExecutor js = (JavascriptExecutor)driver;
            js.executeScript("alert('<<Khatam!! Tata bye bye>>')");
    
        } else {
            
            JavascriptExecutor js = (JavascriptExecutor)driver;
            js.executeScript("alert('               ********System NOT Ready Magaaa********')");
        
        }
        
    }
    public static void switchingoo(WebDriver driver,String title) {
        Set<String> windowIds = driver.getWindowHandles();
            for(String currentWindow : windowIds){
                driver.switchTo().window(currentWindow);
                String currentTitle = driver.getTitle();
                if(currentTitle.contentEquals(title))
                {
                    break;
                }
            }
        // System.out.println("Switching To: " + driver.getTitle());
    }    
    
}