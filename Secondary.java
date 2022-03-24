package Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.Driver;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;

public class Secondary {
  public static void main(String[] args) throws Exception {
    
    System.setProperty("webdriver.chrome.driver" , "D:\\Users\\chromedriver.exe");
    WebDriver driver = new ChromeDriver();
    driver.get("https://fiservservicepoint.fiservapps.com/nav_to.do?uri=%2F$pa_dashboard.do");
    driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
    // driver.findElement(By.id(id)
    driver.manage().window().maximize();
    driver.findElement(By.linkText("Fiserv Associates Click Here to Authenticate")).click();      
    driver.findElement(By.id("i0116")).sendKeys("adarsha.niranjanamurthy@fiserv.com");       
    driver.findElement(By.id("idSIButton9")).click();
    driver.findElement(By.id("passwordInput")).sendKeys("Llghjkl;'");
    driver.findElement(By.id("submitButton")).click();
    int rowPicker = 0;
    for(rowPicker = 4; rowPicker <= 29; rowPicker ++) {

      String[] data = readExcel(rowPicker);
      
      driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
      String homeWindowHandleID = driver.getWindowHandle();
      WebElement searchButton = driver.findElement(By.id("sysparm_search"));
      JavascriptExecutor js = (JavascriptExecutor) driver;
      js.executeScript("arguments[0].click()", searchButton);
      searchButton.clear();
      driver.findElement(By.id("sysparm_search")).sendKeys(data[0]);
      driver.findElement(By.id("sysparm_search")).sendKeys(Keys.ENTER);
      Thread.sleep(3000);
      driver.switchTo().frame("gsft_main");
      driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);

      driver.findElement(By.id("x_fise2_card_svs_projects.revised_live_date")).clear();
      driver.findElement(By.id("x_fise2_card_svs_projects.revised_live_date")).sendKeys("2022-05-31");

      // String[] ids = {"2009A", "2081A", "2085A", "2090A", "2153A", "2130A", "EMV ATM", "5011A-", "ATM - TLS", };
      // Actions act = new Actions(driver);

      // for(String id: ids){
      //   WebDriverWait wait = new WebDriverWait(driver, 60);
      //   wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//td[contains(text(),'Insert a new row...')]")));
      //   WebElement ele = driver.findElement(By.xpath("//td[contains(text(),'Insert a new row...')]"));
      //   act.doubleClick(ele).perform();
        
      //   driver.findElement(By.id("sys_display.LIST_EDIT_x_fise2_card_svs_secondary_job_code.secondary_job_code")).sendKeys(id);
      //   Thread.sleep(1000);
      //   act.sendKeys(Keys.DOWN).perform();
      //   act.sendKeys(Keys.ENTER).perform();
      //   Thread.sleep(2000);
      // }

      driver.findElement(By.id("sysverb_update_and_stay")).click();
      System.out.println(data[0] + ": Completed");
      driver.switchTo().window(homeWindowHandleID);      
    }
  }
  
  public static void writeExcel(int rowPicker, String workFlow, String term, String vru) throws IOException {

    FileInputStream fis = new FileInputStream("D:\\Users\\F3C3XRT\\Documents\\atm.xlsx");
    XSSFWorkbook workbook = new XSSFWorkbook(fis);
    XSSFSheet sheet = workbook.getSheetAt(0);

    sheet.getRow(rowPicker).createCell(11).setCellValue(workFlow);
    sheet.getRow(rowPicker).createCell(12).setCellValue(term);
    sheet.getRow(rowPicker).createCell(13).setCellValue(vru);
    
    fis.close();
    FileOutputStream outputStream = new FileOutputStream("D:\\Users\\F3C3XRT\\Documents\\atm.xlsx");
    workbook.write(outputStream);
    outputStream.close();
  }
    
  public static String[] readExcel(int rowPicker) throws IOException {
    String filePath = "D:\\Users\\F3C3XRT\\Documents";
    String fileName = "atm.xlsx";
    String sheetName = "Sheet1";

    File file = new File(filePath + "\\" + fileName);
    FileInputStream inputStream = new FileInputStream(file);
    XSSFWorkbook workbook = new XSSFWorkbook(inputStream);
    // Read sheet inside the workbook by its name
    XSSFSheet sheet = workbook.getSheet(sheetName);
    // Find number of rows in excel file
    int rowCount = sheet.getLastRowNum();
    // System.out.println(rowCount);

    String workFlow = sheet.getRow(rowPicker).getCell(0).getStringCellValue();
    
    workbook.close();
    return new String[]{workFlow};
  }


}