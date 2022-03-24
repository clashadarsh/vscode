package Test;

import java.io.IOException;
import java.text.ParseException;
import java.util.LinkedHashSet;
import java.util.Scanner;
import java.util.concurrent.TimeUnit;

import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.chrome.ChromeOptions;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.Select;

public class Bear {
  public static void main(String[] args) throws IOException, InterruptedException, ParseException {

    String disputeManager = "Dispute Manager [www.client-central.com]";
    // String cardManagement = "CWSi | Card Management - Choose Client ";
    String disputeSearch = "Repository Viewer - Dispute Search [www.client-central.com]";
    String disputeDetail = "Repository Viewer - Dispute Detail [www.client-central.com]";
    String disputeActionDetail = "Repository Viewer - Dispute Action Detail [www.client-central.com]";
    String disputeActivityDetail = "Repository Viewer - Activity Detail [www.client-central.com]";
                                  //Repository Viewer - Activity Detail [www.client-central.com:443]
    String spc = "CWSi | SPC - Retrieve and Select Transaction";

    Scanner scanner = new Scanner(System.in);
    System.out.print("Enter Token:");
    String token = scanner.nextLine();
    scanner.close();
    System.setProperty("webdriver.chrome.driver", "D:\\chromedriver.exe");
    // ChromeOptions options = new ChromeOptions();
    // options.addArguments("--headless");     
    // options.addArguments("--disable-gpu");
    // options.addArguments("--window-size=1920,1080");
    // options.addArguments("--start-maximized");
    // WebDriver driver = new ChromeDriver(options);
    // System.setProperty("webdriver.chrome.driver", "D:\\Users\\chromedriver.exe");
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

    Pad pad = new Pad();

    pad.switchingoo(driver, disputeManager);
    // if (driver.getTitle().contains("Dispute Manager [www.client-central.com]")) {

    // }
    driver.manage().window().maximize();

    driver.findElement(By.id("fmForm:pr16935975420:1:j_idt121")).click();/** Dispute */
    Thread.sleep(900);
    driver.findElement(By.linkText("Repository Viewer - Dispute Search")).click();/** Rep Viewer */
    // fmForm:j_id124:0:j_id137
    LinkedHashSet<String> hash_Set =  pad.readExcel();

    int rowPicker = 1;
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
      Thread.sleep(500);
      // driver.findElement(By.id("fmForm:j_idt216")).click();
      // Actions act = new Actions(driver);
      // act.sendKeys(Keys.RETURN);
      // Select drpType = new Select(driver.findElement(By.id("fmForm:actnTypeByNameFndrElem_label")));/** Handling dropdown */
      // drpType.selectByVisibleText("Accel Dispute Inquiry");//Accel Dispute Inquiry
      // driver.findElement(By.xpath("//*[@id='fmForm:btnSrch']")).click();
      driver.manage().timeouts().implicitlyWait(15, TimeUnit.SECONDS);

      // bringing in javascript since element is hidden
      
      WebElement searchedDispute = driver.findElement(By.id("fmForm:confRsetTablerpsyvewrevalsrch:0:j_idt283"));
      String id = searchedDispute.getText();
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
        // String ID = driver.findElement(By.id("//*[@id='fmForm:j_id135']/td[3]")).getText();
        WebElement searcheddispute2 = driver.findElement(By.xpath("//a[contains(@id, 'fmForm:block_evalActns:j_idt536:0:j_idt')]"));
        JavascriptExecutor js2 = (JavascriptExecutor)driver;
        js2.executeScript("arguments[0].click()", searcheddispute2);
        Thread.sleep(1000);
        
        driver.close();
        
        pad.switchingoo(driver, disputeActionDetail);
        String addInfo = driver.findElement(By.id("fmForm:block_actnMgrPnl:ActyDtl_magicvalue_FisvAddlCmnt")).getText();
        driver.findElement(By.id("fmForm:btnActn")).click();
        driver.findElement(By.id("fmForm:j_idt133")).click();
        Thread.sleep(2000);
        driver.close();
        
        pad.switchingoo(driver, disputeActivityDetail);
        String cardNumber = driver.findElement(By.xpath("//*[@id='fmForm:block_actyDtlDest:fitmRpsyVewrActyDtlDestInfoPan']/div[2]")).getText();
        String logo = driver.findElement(By.xpath("//*[@id='fmForm:block_actyDtlDest:fitmRpsyVewrActyDtlDestInfoBid']/div[2]")).getText();
        String cardAcceptorName = driver.findElement(By.id("fmForm:block_actyDtlOrig:fitmRpsyVewrActyDtlOrigCardAcprInfoAcqOwn")).getText();
        String stan = driver.findElement(By.xpath("//*[@id='fmForm:block_actyDtlTran:fitmRpsyVewrActyDtlTranInfoStan']/div[2]")).getText();
        String amount = driver.findElement(By.xpath("//*[@id='fmForm:block_actyDtlTran:fitmRpsyVewrActyDtlTranInfoAmt']/div[2]/span")).getText();
        String date = driver.findElement(By.xpath("//*[@id='fmForm:block_actyDtlTranDateTime:fitmRpsyVewrActyDtlTranDateTimeLocl']/div[2]")).getText();
        
        driver.close();
        
        // String expiry = pad.getExpiry(driver, homeWindowHandleID, cardManagement, logo, cardNumber);

        // fetch Claim Number
        String claimNumber = pad.claimFetcher(driver, homeWindowHandleID, spc, logo, cardNumber, stan);
        
        // fetch pos input  and entry mode
        String formattedDate = pad.dateFormatter(date);
        
        String[] pos=  pad.posFetcher(driver, homeWindowHandleID, spc, logo, cardNumber, formattedDate, stan);
        String entryMode = pos[0];
        String posInput = pos[1];
        
        pad.switchingoo(driver, disputeSearch);
        pad.writeExcel(rowPicker, addInfo, cardAcceptorName, amount, date, stan, logo, cardNumber, claimNumber, entryMode, posInput, id);
        // pad.writeExcel(rowPicker, addInfo, cardAcceptorName, amount, date, stan, logo, cardNumber, claimNumber);
        
        System.out.println(rowPicker+" :Fetched Data for: " +disputeID);
        rowPicker++;

      }else{

        driver.close();
        pad.switchingoo(driver, disputeSearch);
        pad.touchExcel(rowPicker, defaultItem);
        System.out.println(rowPicker+" : else Fetched Data for: " +disputeID);
        rowPicker++;

      }
 
    }/**FOR LOOP END*/
    
    JavascriptExecutor js = (JavascriptExecutor)driver;
    js.executeScript("alert('*****Fetched ra mama*****')");
  
  }

    
}