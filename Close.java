package Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Scanner;
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
import org.openqa.selenium.support.ui.Select;

public class Close {
  public static void main(String[] args) throws InterruptedException, IOException {
    String disputeManager = "Dispute Manager [www.client-central.com]";
    // // String systemNotReady = "System Not Ready [www.client-central.com]";
    String disputeDetail = "Repository Viewer - Dispute Detail [www.client-central.com]";
    String cardManagement = "CWSi | Card Management - Choose Client ";
    // String disputeActionCreate = "Dispute Action Create [www.client-central.com]";
    String disputeActionCreate = "Dispute Action Create [www.client-central.com]";
    String createDisputeAction = "Create Dispute Action [www.client-central.com]";
    String disputeSearch = "Repository Viewer - Dispute Search [www.client-central.com]";
    String disputeActionDetail = "Repository Viewer - Dispute Action Detail [www.client-central.com]";
    String disputeActivityDetail = "Repository Viewer - Activity Detail [www.client-central.com]";
    String spc = "CWSi | SPC - Retrieve and Select Transaction";

    String lessThan10 = ": The claim investigation indicates no CB rights exist for transactions under $10 per Network Rules";
    // String lessThan10 = ": The claim investigation indicates no CB rights exist because the card was present at the time of the sale and the merchant terminal met all network requirements";
    // String lessThan10 = ": The claim investigation indicates no CB rights exist for transactions exceeding the 15-item limit per Network Rules";
    
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
    System.setProperty("webdriver.chrome.driver" , "D:\\chromedriver.exe");
    WebDriver driver = new ChromeDriver();
    driver.get("https://www.client-central.com");
    driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
    driver.findElement(By.id("Ecom_User_ID")).sendKeys("amurth00");
    driver.manage().window().maximize();
    driver.findElement(By.id("Ecom_Token")).sendKeys(token);
    driver.findElement(By.id("loginButton2")).click();
    //switching to iframe
    String homeWindowHandleID = driver.getWindowHandle();
    driver.switchTo().frame("main");
    driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
    System.out.println("We are switching to the iframe");
    driver.findElement(By.xpath("//a[contains(text(),'Dispute Manager')]")).click();
    //switching to Dispute Manager

    Pad pad = new Pad();

    pad.switchingoo(driver,disputeManager);
    driver.manage().window().maximize();

    driver.findElement(By.id("fmForm:pr16935975420:1:j_idt121")).click();/** Dispute */
    Thread.sleep(900);
    driver.findElement(By.linkText("Repository Viewer - Dispute Search")).click();/** Rep Viewer */

    LinkedHashSet<String> hash_Set =  readExcel();

    for(String disputeID: hash_Set){
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
      // driver.findElement(By.id("fmForm:j_idt216")).click();
      // Actions act = new Actions(driver);
      // act.sendKeys(Keys.RETURN);
      // Select drpType = new Select(driver.findElement(By.id("fmForm:actnTypeByNameFndrElem_label")));/** Handling dropdown */
      // drpType.selectByVisibleText("Accel Dispute Inquiry");//Accel Dispute Inquiry
      // driver.findElement(By.xpath("//*[@id='fmForm:btnSrch']")).click();
      driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);

      // bringing in javascript since element is hidden
      
      WebElement searchedDispute = driver.findElement(By.id("fmForm:confRsetTablerpsyvewrevalsrch:0:j_idt283"));
      // String id = searchedDispute.getText();
      Thread.sleep(1000);
      // searchedDispute.click();
      JavascriptExecutor jsw = (JavascriptExecutor) driver;
      jsw.executeScript("arguments[0].click()", searchedDispute);
      // switching to Repository Viewer;
      pad.switchingoo(driver, disputeDetail);

      // Select select = new Select(driver.findElement(By.id("fmForm:evalResnByNamePckrElem")));
      // WebElement option = select.getFirstSelectedOption();
      // String defaultItem = option.getText();//fmForm:evalResnByNamePckrElem_label
      String defaultItem = driver.findElement(By.id("fmForm:evalResnByNamePckrElem_label")).getText();
      if (defaultItem.isBlank()) {
        WebElement searcheddispute2 = driver.findElement(By.xpath("//a[contains(@id, 'fmForm:block_evalActns:j_idt536:0:j_idt')]"));
        JavascriptExecutor js2 = (JavascriptExecutor)driver;
        js2.executeScript("arguments[0].click()", searcheddispute2);
        
        pad.switchingoo(driver, disputeActionDetail);
        String addInfo = driver.findElement(By.id("fmForm:block_actnMgrPnl:ActyDtl_magicvalue_FisvAddlCmnt")).getText();
        System.out.println(addInfo);
        driver.findElement(By.id("fmForm:btnActn")).click();
        driver.findElement(By.id("fmForm:j_idt133")).click();
        Thread.sleep(2000);
        driver.close();
        
        pad.switchingoo(driver, disputeActivityDetail);
        String cardNumber = driver.findElement(By.xpath("//*[@id='fmForm:block_actyDtlDest:fitmRpsyVewrActyDtlDestInfoPan']/div[2]")).getText();
        String logo = driver.findElement(By.xpath("//*[@id='fmForm:block_actyDtlDest:fitmRpsyVewrActyDtlDestInfoBid']/div[2]")).getText();
        String stan = driver.findElement(By.xpath("//*[@id='fmForm:block_actyDtlTran:fitmRpsyVewrActyDtlTranInfoStan']/div[2]")).getText();
        driver.close();
        // String cardNumber = driver.findElement(By.xpath("//*[@id='fmForm:j_id249']/td[3]/span")).getText();
        char[] ch = new char[4];
        ch[0] = cardNumber.charAt(12);
        ch[1] = cardNumber.charAt(13);
        ch[2] = cardNumber.charAt(14);
        ch[3] = cardNumber.charAt(15);
        String last4Digits = String.valueOf(ch);
        // String amount = driver.findElement(By.xpath("//tbody/tr[@id='fmForm:j_id253']/td[3]")).getText();
        // String date = driver.findElement(By.xpath("//tbody/tr[@id='fmForm:j_id273']/td[3]")).getText();

        // fetch Claim Number
        String claimNumber = pad.claimFetcher(driver, homeWindowHandleID, spc, logo, cardNumber, stan);
        String memo = last4Digits+":"+claimNumber+":"+disputeID+lessThan10;
        
        pad.switchingoo(driver, disputeDetail);
        WebElement memoButton = driver.findElement(By.id("fmForm:block_evalMemos:cbtnAdd"));
        JavascriptExecutor js3 = (JavascriptExecutor)driver;
        js3.executeScript("arguments[0].click()", memoButton);
        // WebElement addButton = driver.findElement(By.id("fmForm:tabMemo:memoMgrPnlCbtnAdd"));
        // JavascriptExecutor js4 = (JavascriptExecutor)driver;
        // js4.executeScript("arguments[0].click()", addButton);
        driver.manage().timeouts().implicitlyWait(2, TimeUnit.SECONDS); 
        driver.findElement(By.id("fmForm:block_evalMemos:evalMemosMemoAddMpnlItaDscr")).sendKeys(memo);
        
        driver.findElement(By.id("fmForm:block_evalMemos:evalMemosMemoAddMpnlCbtnAdd")).click();
        Thread.sleep(1000);

        Thread.sleep(3000);
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        driver.findElement(By.id("fmForm:evalResnByNamePckrElem_label")).click();
        Thread.sleep(200);
        driver.findElement(By.id("fmForm:evalResnByNamePckrElem_2")).click();
        Thread.sleep(100);

        driver.findElement(By.id("fmForm:btnUpd")).click();
        Thread.sleep(1000);
        driver.close();

        // continiuw madpa after adding memp
        driver.switchTo().window(homeWindowHandleID);
        driver.switchTo().frame("main");
        driver.findElement(By.xpath("//a[contains(text(),'Card Management')]")).click();
        pad.switchingoo(driver, cardManagement);
        driver.findElement(By.xpath("//label[contains(text(),'Enter Card Number')]")).click();
        driver.findElement(By.id("chooseClient_cardNumber")).sendKeys(cardNumber);
        driver.findElement(By.id("chooseClient_continue")).click();
        
        driver.findElement(By.id("enotesAddButton")).click();
        driver.findElement(By.id("notetext")).sendKeys(memo);
        driver.findElement(By.id("addEnoteYes")).click();
        Thread.sleep(1000);
        
        driver.close();

        System.out.println(disputeID + ": Closed");
        pad.switchingoo(driver, disputeSearch);

      } else {

        driver.close();
        pad.switchingoo(driver, disputeSearch);
        System.out.println(disputeID + ": " + defaultItem);

      }

    }/**FOR LOOP END*/
      
    JavascriptExecutor js = (JavascriptExecutor)driver;
    js.executeScript("alert('<<Khatam!! Tata bye bye>>')");
    
  }

  // Internal Methods

  public static void addNote(WebDriver driver, String memo){
    List <WebElement> x = driver.findElements(By.xpath("//input[@id='viewCardDetails_cardExpirationDate']"));
    if(x.size()!=0){
      // click on addNote and send memo
      driver.findElement(By.id("enotesAddButton")).click();
      driver.findElement(By.id("notetext")).sendKeys(memo);
      driver.findElement(By.id("addEnoteYes")).click();

    } else {
      driver.findElement(By.id("chooseCardHolder_chooseCardHolderSearch")).click();
      driver.findElement(By.id("submenu0")).click();
      
    }
  }

  public static LinkedHashSet<String> readExcel() throws IOException {
    String filePath = "D:\\Users\\F3C3XRT\\Documents";
    String fileName = "Close.xlsx";
    String sheetName = "Sheet1";

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
    
}