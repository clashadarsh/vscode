package Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.concurrent.TimeUnit;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openqa.selenium.By;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.By.ByXPath;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.support.ui.Select;

public class ATM {
  public static void main(String[] args) throws Exception {
    
    System.setProperty("webdriver.chrome.driver" , "D:\\chromedriver\\chromedriver.exe");
    WebDriver driver = new ChromeDriver();
    driver.get("https://fiservservicepoint.fiservapps.com/nav_to.do?uri=%2F$pa_dashboard.do");
    driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
    // driver.findElement(By.id(id)
    driver.manage().window().maximize();
    driver.findElement(By.linkText("Fiserv Associates Click Here to Authenticate")).click();      
    driver.findElement(By.id("i0116")).sendKeys("adarsha.niranjanamurthy@fiserv.com");       
    driver.findElement(By.id("idSIButton9")).click();
    driver.findElement(By.id("passwordInput")).sendKeys("Visitor$#@1");
    driver.findElement(By.id("submitButton")).click();
    int rowPicker = 1;
    for(rowPicker = 1; rowPicker <= 64; rowPicker ++) {

      String[] data = readExcel(rowPicker);
      
      driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
      String homeWindowHandleID = driver.getWindowHandle();
      WebElement searchButton = driver.findElement(By.id("sysparm_search"));
      JavascriptExecutor js = (JavascriptExecutor) driver;
      js.executeScript("arguments[0].click()", searchButton);
      searchButton.clear();
      driver.findElement(By.id("sysparm_search")).sendKeys(data[0]);
      driver.findElement(By.id("sysparm_search")).sendKeys(Keys.ENTER);
      
      driver.switchTo().frame("gsft_main");

      driver.findElement(By.xpath("//td[contains(text(),'In Progress')]")).click();

      driver.switchTo().window(homeWindowHandleID);


      // String terminal = driver.findElement(By.id("sys_original.x_fise2_card_svs_atm_add.terminal_id")).getAttribute("value");
      // if (terminal == "") {
      //   // driver.findElement(By.xpath("//span[contains(text(),'Install Info')]")).click();
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.requested_live_date")).clear();
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.requested_live_date")).sendKeys("12/6/2021");
      //   driver.findElement(By.id("status.x_fise2_card_svs_atm_add.requested_com_ready_date")).click();
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.requested_com_ready_date")).clear();
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.requested_com_ready_date")).sendKeys("11/1/2021");
      //   driver.findElement(By.id("sys_display.x_fise2_card_svs_atm_add.atm_model_type")).sendKeys(data[6]);
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.project_contact_name")).clear();
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.project_contact_name")).sendKeys("Jason Ries");
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.project_contact_phone")).clear();
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.project_contact_phone")).sendKeys("803-605-9887");
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.project_contact_email")).clear();
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.project_contact_email")).sendKeys("jason.ries@southstatebank.com");
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.site_contact_name")).clear();
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.site_contact_name")).sendKeys("Jason Ries");
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.site_contact_phone")).clear();
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.site_contact_phone")).sendKeys("803-605-9887");
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.site_contact_email")).clear();
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.site_contact_email")).sendKeys("jason.ries@southstatebank.com");
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.drop_location_site_name")).sendKeys(data[1]);
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.drop_location_street")).sendKeys(data[2]);
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.drop_location_city")).sendKeys(data[3]);
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.drop_location_state")).sendKeys(data[4]);
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.drop_location_zip")).sendKeys(data[5]);
      //   Select drp = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.circuit_owner")));
      //   drp.selectByVisibleText("Premier Central (A01)");
        
      //   Select drp1 = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.connectivity")));
      //   drp1.selectByVisibleText("Client WAN");
      //   Select drpp = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.first_client_wan_atm")));
      //   drpp.selectByVisibleText("Yes");
      //   Select drp2 = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.protocol")));
      //   drp2.selectByVisibleText("TCP/IP");
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.customer_atm_ip_address")).sendKeys(data[7]);
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.tcp_ip_default_gateway")).sendKeys(data[8]);
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.tcp_ip_subnet_mask")).sendKeys(data[9]);
      //   Select drp3 = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.atm_configured_as")));
      //   drp3.selectByVisibleText("Client");

      //   if (data[10].contentEquals("Yes")) {
      //     Select drp4 = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.image_processing")));
      //     drp4.selectByVisibleText("Yes");
      //     Select drp5 = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.image_processor")));
      //     drp5.selectByVisibleText("SCO");
      //   } else {
      //     Select drp4 = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.image_processing")));
      //     drp4.selectByVisibleText("No");
      //   }
      
      //   Select drp6 = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.vqp_associated_workflow")));
      //   drp6.selectByVisibleText("No");
      //   Select drp7 = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.emulation_operation_mode")));
      //   drp7.selectByVisibleText("NCR Native");
      //   Select drp8 = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.application_software")));
      //   drp8.selectByVisibleText("VISTA 5.4+");
      //   Select drp9 = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.device_handler_type")));
      //   drp9.selectByVisibleText("AMDH");
      //   Select drp10 = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.ej_upload")));
      //   drp10.selectByVisibleText("No");
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.surcharge_amount")).clear();
      //   driver.findElement(By.id("x_fise2_card_svs_atm_add.surcharge_amount")).sendKeys("3.50");
      //   Select drp11 = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.da")));
      //   drp11.selectByVisibleText("No");
      //   Select drp12 = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.tls_encryption")));
      //   drp12.selectByVisibleText("Version 1.2");
      //   Select drp13 = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.associated_workflow")));
      //   drp13.selectByVisibleText("Yes");
      //   driver.findElement(By.id("sys_display.x_fise2_card_svs_atm_add.host_workflow")).clear();
      //   driver.findElement(By.id("sys_display.x_fise2_card_svs_atm_add.host_workflow")).sendKeys("HOST0003832");
      //   driver.findElement(By.id("status.x_fise2_card_svs_atm_add.host_workflow")).click();
      //   // Generate terminal ID
      //   driver.findElement(By.xpath("//*[@id='element.x_fise2_card_svs_atm_add.terminal_id']/div[3]/a[1]")).click();
      //   Thread.sleep(1000);
      //   // driver.findElement(By.xpath("//span[contains(text(),'A98 Info')]")).click();
      //   driver.findElement(By.id("sys_display.x_fise2_card_svs_atm_add.term_mster_key_type")).clear();
      //   driver.findElement(By.id("sys_display.x_fise2_card_svs_atm_add.term_mster_key_type")).sendKeys("A98 - DIEB EPP7 SHA256 RKT/Triple DES");
      //   driver.findElement(By.id("status.x_fise2_card_svs_atm_add.term_mster_key_type")).click();
        
      //   // Generate VRU ID
      //   driver.findElement(By.xpath("//*[@id='element.x_fise2_card_svs_atm_add.vru_id']/div[3]/a")).click();
      //   Thread.sleep(1000);
      //   // driver.findElement(By.xpath("//span[contains(text(),'ATM Monitoring/Site Profile')]")).click();
      //   Select drp14 = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.monitor_attachment")));
      //   drp14.selectByVisibleText("Not Yet Determined");
        
      //   driver.findElement(By.xpath("//button[@id='sysverb_update_and_stay']")).click();
        
      //   String term = driver.findElement(By.id("sys_original.x_fise2_card_svs_atm_add.terminal_id")).getAttribute("value");
      //   // driver.findElement(By.xpath("//span[contains(text(),'A98 Info')]")).click();
      //   String vru = driver.findElement(By.id("x_fise2_card_svs_atm_add.vru_id")).getAttribute("value");
        
      //   writeExcel(rowPicker, data[0], term, vru);
        
      //   System.out.println("Workflow Created: " + data[0] + "|" +term + "|" + vru);
      //   driver.switchTo().window(homeWindowHandleID);
        
      // } else {
      //   String vru = driver.findElement(By.id("x_fise2_card_svs_atm_add.vru_id")).getAttribute("value");
      //   Select drpp = new Select(driver.findElement(By.id("x_fise2_card_svs_atm_add.first_client_wan_atm")));
      //   drpp.selectByVisibleText("Yes");
      //   driver.findElement(By.xpath("//button[@id='sysverb_update_and_stay']")).click();
      //   System.out.println("workflow Present : " + data[0] + "|" +terminal + "|" + vru);
      //   // writeExcel(rowPicker, data[0], terminal, vru);
      //   driver.switchTo().window(homeWindowHandleID);

      // }

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
    String site = sheet.getRow(rowPicker).getCell(1).getStringCellValue();
    String street = sheet.getRow(rowPicker).getCell(2).getStringCellValue();
    String city = sheet.getRow(rowPicker).getCell(3).getStringCellValue();
    String state = sheet.getRow(rowPicker).getCell(4).getStringCellValue();
    String zip = String.valueOf(sheet.getRow(rowPicker).getCell(5).getNumericCellValue());
    zip = zip.indexOf(".") < 0 ? zip : zip.replaceAll("0*$", "").replaceAll("\\.$", "");
    String model = sheet.getRow(rowPicker).getCell(6).getStringCellValue();
    String ip1 = sheet.getRow(rowPicker).getCell(7).getStringCellValue();
    String ip2 = sheet.getRow(rowPicker).getCell(8).getStringCellValue();
    String ip3 = sheet.getRow(rowPicker).getCell(9).getStringCellValue();
    String ImgProc = sheet.getRow(rowPicker).getCell(10).getStringCellValue();
    
    workbook.close();
    return new String[]{workFlow,site,street,city,state,zip,model,ip1,ip2,ip3,ImgProc};
  }


}