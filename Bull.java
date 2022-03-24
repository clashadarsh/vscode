package Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Scanner;
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
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

public class Bull {
    public static void main(String[] args) throws InterruptedException, IOException {
        String disputeManager = "Dispute Manager [www.client-central.com]";
        // String systemNotReady = "System Not Ready [www.client-central.com]";
        String disputeDetail = "Repository Viewer - Dispute Detail [www.client-central.com]";
        String cardManagement = "CWSi | Card Management - Choose Client ";
        String disputeActionCreate = "Dispute Action Create [www.client-central.com]";
        String createDisputeAction = "Create Dispute Action [www.client-central.com]";
        String disputeSearch = "Repository Viewer - Dispute Search [www.client-central.com]";
        String disputeActionDetail = "Repository Viewer - Dispute Action Detail [www.client-central.com]";
        String disputeActivityDetail = "Repository Viewer - Activity Detail [www.client-central.com]";
    
        String filePath = "D:\\Users\\F3C3XRT\\Documents";
        String fileName = "Book.xlsx";
        String sheetName = "Sheet1";
        Scanner scanner = new Scanner(System.in);
        System.out.print("Enter Token:");
        String token = scanner.nextLine();
        scanner.close();
        // System.setProperty("webdriver.chrome.driver", "D:\\chromedriver.exe");
        // ChromeOptions options = new ChromeOptions();
        // options.addArguments("--headless");    
        // options.addArguments("--disable-gpu");
        // options.addArguments("--window-size=1920,1080");
        // options.addArguments("--start-maximized");
        // WebDriver driver = new ChromeDriver(options);
        System.setProperty("webdriver.chrome.driver", "D:\\chromedriver.exe");
        WebDriver driver = new ChromeDriver();
        driver.get("https://www.client-central.com");
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        driver.findElement(By.id("Ecom_User_ID")).sendKeys("amurth00");
        driver.manage().window().maximize();
        driver.findElement(By.id("Ecom_Token")).sendKeys(token);
        driver.findElement(By.id("loginButton2")).click();
        // switching to iframe
        String homeWindowHandleID = driver.getWindowHandle();
        // String homeWindowTitle = driver.getTitle(); Switchingoo nu use madbodu
        System.out.println("Parent window: " + homeWindowHandleID);
        driver.switchTo().frame("main");
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        System.out.println("We are switching to the iframe");
        driver.findElement(By.xpath("//a[contains(text(),'Dispute Manager')]")).click();
        // switching to Dispute Manager
        switchingoo(driver, disputeManager);
        // driver.findElement(By.xpath("//*[@id='fmForm:j_id103:1:j_id108']")).click();/** Dispute */
        // Thread.sleep(500);
        // driver.findElement(By.id("fmForm:j_id124:0:j_id137")).click();/** Rep Viewer */
        driver.manage().window().maximize();

        driver.findElement(By.id("fmForm:pr16935975420:1:j_idt121")).click();/** Dispute */
        Thread.sleep(900);
        driver.findElement(By.linkText("Repository Viewer - Dispute Search")).click();/** Rep Viewer */
        // fmForm:j_id124:0:j_id137
        File file = new File(filePath + "\\" + fileName);
        FileInputStream inputStream = new FileInputStream(file);
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
        // Read sheet inside the workbook by its name
        XSSFSheet sheet = workbook.getSheet(sheetName);
        // Find number of rows in excel file
        LinkedHashSet<String> hash_Set = new LinkedHashSet<String>();
        int rowCount = sheet.getLastRowNum();
        int colCount = sheet.getRow(1).getLastCellNum();
        // Create a loop over all the rows of excel file to read it
        for (int r = 0; r <= rowCount; r++) {
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

        int count = 1;
        for (String disputeID : hash_Set) {
            driver.findElement(By.id("fmForm:btnClr")).click();
            Thread.sleep(2000);
            driver.findElement(By.xpath("//*[@id='fmForm:evalId']")).clear();
            driver.findElement(By.xpath("//*[@id='fmForm:evalId']")).sendKeys(disputeID);
            driver.findElement(By.id("fmForm:j_idt209")).click();
            // driver.findElement(By.id("fmForm:actnTypeByNameFndrElem_11")).click();
            driver.findElement(By.id("fmForm:actnTypeByNameFndrElem_label")).click();
            driver.findElement(By.id("fmForm:actnTypeByNameFndrElem_filter")).clear();
            driver.findElement(By.id("fmForm:actnTypeByNameFndrElem_filter")).sendKeys("Accel Dispute Inq");
            Thread.sleep(500);
            driver.findElement(By.id("fmForm:actnTypeByNameFndrElem_11")).click();
            Thread.sleep(500);
            WebElement sear = driver.findElement(By.xpath("//*[@id='fmForm:btnSrch']"));
            JavascriptExecutor js = (JavascriptExecutor) driver;
            js.executeScript("arguments[0].click()", sear);
            WebElement searchedDispute = driver.findElement(By.id("fmForm:confRsetTablerpsyvewrevalsrch:0:j_idt283"));
            String id = searchedDispute.getText();
            Thread.sleep(1000);
            // searchedDispute.click();
            JavascriptExecutor jsw = (JavascriptExecutor) driver;
            jsw.executeScript("arguments[0].click()", searchedDispute);
            // switching to Repository Viewer;
            switchingoo(driver, disputeDetail);

            String defaultItem = driver.findElement(By.id("fmForm:evalResnByNamePckrElem_label")).getText();
        
            if (defaultItem.isBlank()) {
                // driver.close();
                driver.findElement(By.id("fmForm:btnActn")).click();
                driver.findElement(By.id("fmForm:j_idt137")).click();
                Thread.sleep(2000);
                driver.close();
                // Actions actions = new Actions(driver);
                // actions.contextClick(driver.findElement(By.xpath("//*[@id='fmForm:j_id1160']/table/tbody/tr[2]/td[1] | //*[@id='fmForm:j_id1203']/table/tbody/tr[2]/td[1] | //*[@id='fmForm:j_id1296']/table/tbody/tr[2]/td[1] | //*[@id='fmForm:j_id1283']/table/tbody/tr[2]/td[1] | //*[@id='fmForm:j_id1253']/table/tbody/tr[2]/td[1] | //*[@id='fmForm:j_id1240']/table/tbody/tr[2]/td[1] | //*[@id='fmForm:j_id1203']/table/tbody/tr[2]/td[1] | //*[@id='fmForm:j_id1333']/table/tbody/tr[2]/td[1]"))).perform(); /** id id damn dynamic */
                // driver.findElement(By.id("fmForm:evalActnMgrMpnlCmnuActnType")).click();
                // driver.close();
                // swithing to disputeActionCreate
                switchingoo(driver, createDisputeAction);
                driver.findElement(By.id("fmForm:tabPaneCratEvalActn:arptPlanSetPlansRset0:0:arptManyToOneActnTypeTimeWndw0:1:j_idt198")).click();
                Thread.sleep(3000);
                driver.close();
                switchingoo(driver, disputeActionCreate);
                // Select drpType2 = new Select(driver.findElement(By.id("fmForm:ActyDtl_FisvMsgResnCode_elem")));
                // drpType2.selectByVisibleText("0081 - Chargeback - Fraud, unauthorized");
                String cardNumber = driver.findElement(By.xpath("//*[@id='fmForm:ActyDtl_magicvalue_PAN']")).getText();
                driver.findElement(By.id("fmForm:ActyDtl_magicvalue_FisvMsgResnCode_elem_label")).click();
                driver.findElement(By.id("fmForm:ActyDtl_magicvalue_FisvMsgResnCode_elem_filter")).sendKeys("0081");
                Thread.sleep(300);
                driver.findElement(By.id("fmForm:ActyDtl_magicvalue_FisvMsgResnCode_elem_7")).click();
                Thread.sleep(1500);
                driver.findElement(By.id("fmForm:ActyDtl_magicvalue_FisvMbrMsgText_elem")).sendKeys("Cardholder did not authorize transaction");
                Thread.sleep(300);
                driver.findElement(By.id("fmForm:j_idt517")).click();
                Thread.sleep(1000);

                driver.findElement(By.id("fmForm:ActyDtl_magicvalue_FisvFrdType_elem_label")).click();
                driver.findElement(By.id("fmForm:ActyDtl_magicvalue_FisvFrdType_elem_filter")).sendKeys("6");
                Thread.sleep(300);
                driver.findElement(By.id("fmForm:ActyDtl_magicvalue_FisvFrdType_elem_7")).click();
                Thread.sleep(1500);

                driver.findElement(By.id("fmForm:ActyDtl_magicvalue_FisvDocInd_elem_label")).click();
                driver.findElement(By.id("fmForm:ActyDtl_magicvalue_FisvDocInd_elem_filter")).sendKeys("0");
                Thread.sleep(300);
                driver.findElement(By.id("fmForm:ActyDtl_magicvalue_FisvDocInd_elem_1")).click();

                // switching to parent window >>cardManagemt
                driver.switchTo().window(homeWindowHandleID);
                driver.switchTo().frame("main");
                driver.findElement(By.xpath("//a[contains(text(),'Card Management')]")).click();
                // switching to cardManagement window
                switchingoo(driver, cardManagement);
                driver.findElement(By.xpath("//label[contains(text(),'Enter Card Number')]")).click();
                driver.findElement(By.id("chooseClient_cardNumber")).sendKeys(cardNumber);
                driver.findElement(By.id("chooseClient_continue")).click();

                String estring = getExpiry(driver);
                // String estring =
                // driver.findElement(By.xpath("//input[@id='viewCardDetails_cardExpirationDate']")).getAttribute("value");
                char[] ch = new char[4];
                ch[0] = estring.charAt(3);
                ch[1] = estring.charAt(4);
                ch[2] = estring.charAt(0);
                ch[3] = estring.charAt(1);

                String date = String.valueOf(ch);
                driver.close(); /* CLOSE CARD MANAGEMENT */
                switchingoo(driver, disputeActionCreate);
                driver.findElement(By.id("fmForm:ActyDtl_magicvalue_DateExp_elem")).sendKeys(date);
                driver.findElement(By.id("fmForm:j_idt479")).click();
                Thread.sleep(1500);
                driver.findElement(By.id("fmForm:ActyDtl_magicvalue_FisvAddlCmnt_elem")).sendKeys("Fraudulent use for online transactions. Cardholder did not participate in the transaction");
                Thread.sleep(300);
                driver.findElement(By.id("fmForm:j_idt517")).click();
                Thread.sleep(1000);

                // driver.findElement(By.id("fmForm:ActyDtl_magicvalue_FisvDocInd_elem_label")).click();
                // WebElement shit = driver.findElement(By.id("fmForm:ActyDtl_magicvalue_FisvDocInd_elem_filter"));
                // JavascriptExecutor jsww = (JavascriptExecutor) driver;
                // jsww.executeScript("arguments[0].click()", shit);
                
                driver.findElement(By.id("fmForm:btnAdd")).click();
                Thread.sleep(3000);
                driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
                driver.findElement(By.id("fmForm:evalResnByNamePckrElem_label")).click();
                Thread.sleep(200);
                driver.findElement(By.id("fmForm:evalResnByNamePckrElem_1")).click();
                Thread.sleep(100);

                driver.findElement(By.id("fmForm:btnUpd")).click();
                Thread.sleep(1000);
                System.out.println(count+" Chargeback Submitted: " + disputeID + " " + id +"|"+ date);
                driver.close(); /* CLOSE CHARGEBACK FORM */
                count++;
                switchingoo(driver, disputeSearch);

            }else{

                driver.close();
                switchingoo(driver, disputeSearch);
                System.out.println(count+" " + defaultItem + ": " + disputeID);        
                count++;
              
            }
            
        }
        JavascriptExecutor js = (JavascriptExecutor) driver;
        js.executeScript("alert('<<Khatam!! Tata bye bye>>')");

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

    public static String getExpiry(WebDriver driver) {
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
        return eString;
    }

}