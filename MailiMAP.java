package casemgmt;

import java.util.Date;
import java.util.Properties;
import java.util.concurrent.TimeUnit;
import java.awt.Toolkit;
import java.awt.datatransfer.Clipboard;
import java.awt.datatransfer.StringSelection;
//   import com.j256.simplemagic.ContentInfo;
//   import com.j256.simplemagic.ContentInfoUtil;
//   import com.j256.simplemagic.ContentType;

import org.apache.logging.log4j.LogManager;
import org.apache.logging.log4j.Logger;
// import org.apache.commons.math3.analysis.function.Add;
import org.openqa.selenium.By;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.Select;
import org.openqa.selenium.support.ui.WebDriverWait;

import io.github.bonigarcia.wdm.WebDriverManager;
import net.bytebuddy.build.ToStringPlugin.Enhance.Prefix;

import javax.mail.Address;
//   import com.testautomationguru.utility.PDFUtil;
// import javax.mail.Flags;
import javax.mail.Folder;
import javax.mail.Message;/**models an email message */
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.Flags.Flag;
//   import javax.mail.search.FlagTerm;
// import javax.security.auth.Subject;

import com.google.common.base.Strings;

public class MailiMAP {
    public static void main(String[] args) throws Exception {
        String username = "1dc\\svc-tcoe-auto-tester";
        String password = "xjsE@k328";
        MailiMAP sha = new MailiMAP();
        
        WebDriverManager.chromedriver().setup();//initialize chrome

        WebDriver driver = new ChromeDriver();
        // driver.get("https://fiservtestservicepoint.fiservapps.com/navpage.do");
        driver.get("https://fiservservicepoint.fiservapps.com/navpage.do");
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
        // driver.findElement(By.id(id)
        driver.manage().window().maximize();
        driver.findElement(By.linkText("Fiserv Associates Click Here to Authenticate")).click();        
        driver.findElement(By.id("i0116")).sendKeys("adarsha.niranjanamurthy@fiserv.com");
        driver.manage().timeouts().implicitlyWait(10, TimeUnit.SECONDS);
       
        driver.findElement(By.id("idSIButton9")).click();
        driver.findElement(By.id("passwordInput")).sendKeys("Mmghjkl;'");
        driver.findElement(By.id("submitButton")).click();
        Thread.sleep(2000);
        driver.findElement(By.id("filter")).sendKeys("Cases");
        Thread.sleep(6000);
        driver.findElement(By.xpath("//*[@id='8f0ab177c342310015519f2974d3ae12']/div/div")).click();
        
        sha.OutlookEmail(username, password, driver);
        
    }
    
    public void OutlookEmail(String username, String password, WebDriver driver) throws Exception {
        
        Properties properties = System.getProperties();
        // Set important information to properties object
        properties.setProperty("mail.store.protocol", "imap");
        properties.setProperty("mail.imap.ssl.enable", "true");
        properties.setProperty("mail.imap.partialfetch", "false");
        properties.setProperty("mail.mime.base64.ignoreerrors", "true");
        properties.setProperty("mail.imap.starttls.enable", "true");
        properties.setProperty("mail.imap.port", "993");
        properties.setProperty("mail.debug", "false");
        
        Session mailSession = Session.getInstance(properties);
        mailSession.setDebug(true);
        Store store = mailSession.getStore("imap");
        store.connect("chamail.1dc.com", username, password);
        
        Folder folder = store.getFolder("Implementation");
        folder.open(Folder.READ_WRITE);
        
        System.out.println("Total Message:" + folder.getMessageCount());
        System.out.println("Unread Message:" + folder.getUnreadMessageCount());
        
        Message[] messages = folder.getMessages();
        
        MyLab lab = new MyLab();
        
        // for(int i=0;i<=1;i++) {
        for (Message mail : messages) {
        // if (!mail.isSet(Flags.Flag.SEEN)) {
            // Message mail = messages[0];
            
            Address from = mail.getFrom()[0];/**jathre printings come by default */
            Address to= mail.getAllRecipients()[0];
            String subject = mail.getSubject();
            // String body = getEmailBody(mail);
            String subject2 = subject.replaceAll("FW: ", "");
            String sub = subject2.replaceAll("\\s"," ");    
            // String subjj = sub.replaceAll("[?|:/*<>]", " ");
            // String subj=subjj.replace("\\", " ");
            String subFinal = lab.cardRemover(sub);

            System.out.println(sub);
        
            // String descr = mail.getContent().toString();
            // String emailbody = getEmailBody(mail);
            // String content = getEmailBody(mail).body().text();
            String body = lab.getTextFromMessage(mail);

            // Get company name
            String company = body;
            company = company.substring(company.indexOf("Company :") + 10);
            company = company.substring(0, company.indexOf("Client Deployment:"));
            // Get callback name
            String callbackName = body;
            callbackName = callbackName.substring(callbackName.indexOf("Call Back Name :") + 17);
            callbackName = callbackName.substring(0, callbackName.indexOf("Company :"));
            // Get senders email domain
            String emailDomain = body;
            emailDomain = emailDomain.substring(emailDomain.indexOf("@") + 0);
            emailDomain = emailDomain.substring(0, emailDomain.indexOf(">"));
            // Get company name
            String clientDeployment = body;
            clientDeployment = clientDeployment.substring(clientDeployment.indexOf("Client Deployment:") + 18);
            clientDeployment = clientDeployment.substring(0, clientDeployment.indexOf("From:"));
            clientDeployment = clientDeployment.replaceAll("\\s"," ");
            // Get from address
            // String compfromExcel = "";
            // int rowPicker = 0;
            // String[] data0 = lab.readExcel(rowPicker);
            // int rows = Integer.parseInt(data0[2]);
            // ExcelMatch:for(rowPicker = 0; rowPicker <= rows ;  rowPicker ++) {

            //     String[] data = lab.readExcel(rowPicker);
            //     if (data[0].equalsIgnoreCase(domain)) {
            //         compfromExcel = data[1];
            //         System.out.println(data[0] + " | " +data[1]);
            //         break ExcelMatch;
            //     }
            // }

            String bodyFinal = lab.cardRemover(body);

            int size = mail.getSize();
            // Date receivedDate = mail.getReceivedDate();

            String homeWindowHandleID = driver.getWindowHandle();

            driver.switchTo().frame("gsft_main");
                
            driver.findElement(By.id("sys_display.sn_customerservice_case.company")).sendKeys(company);
            Thread.sleep(4000);
            Actions acto = new Actions(driver);
            acto.sendKeys(Keys.DOWN).perform();
            acto.sendKeys(Keys.ENTER).perform();
            
            // List<WebElement> elements = driver.findElements(By.xpath("//ul[@class='element_reference_input']//li/span"));
            // System.out.println(elements.size());
            driver.findElement(By.id("sys_display.sn_customerservice_case.u_client_deployment_id")).sendKeys("Deployment not found"); 
            Thread.sleep(2000);
            Actions act = new Actions(driver);
            act.sendKeys(Keys.DOWN).perform();
            act.sendKeys(Keys.ENTER).perform();
            driver.findElement(By.id("status.sn_customerservice_case.u_client_deployment_id")).click(); 
            
            // driver.findElement(By.id("sys_display.sn_customerservice_case.contact")).sendKeys(fromBody);
            driver.findElement(By.id("sys_display.sn_customerservice_case.contact")).sendKeys("Contact the");
            Thread.sleep(2000);
            Actions act1 = new Actions(driver);
            act1.sendKeys(Keys.DOWN).perform();
            act1.sendKeys(Keys.ENTER).perform();
           
            driver.findElement(By.id("sn_customerservice_case.u_callback_name")).sendKeys(callbackName);
            driver.findElement(By.id("status.sn_customerservice_case.contact")).click();
            
            String characterFilter = "[^\\p{L}\\p{M}\\p{N}\\p{P}\\p{Z}\\p{Cf}\\p{Cs}\\s]";
            String emotionlessBody = bodyFinal.replaceAll(characterFilter,"");            
            WebElement inputField = driver.findElement(By.id("sn_customerservice_case.description"));
            String prefix = lab.prefixSelector(clientDeployment);
            // String prefix = "STAR: ";
            String shortDescription = prefix+subFinal;
            driver.findElement(By.id("sn_customerservice_case.short_description")).sendKeys(shortDescription);
            Clipboard clipboard = Toolkit.getDefaultToolkit().getSystemClipboard();
            StringSelection transferable = new StringSelection(emotionlessBody);
            clipboard.setContents(transferable, null);//copy short description to clipboard
            inputField.click();inputField.sendKeys(Keys.CONTROL , "v");
            
            // driver.findElement(By.id("sn_customerservice_case.description")).sendKeys(emotionlessBody);
            Select drpx = new Select(driver.findElement(By.name("sn_customerservice_case.u_case_reporting_category")));
            drpx.selectByVisibleText("Service");
            // driver.findElement(By.id("sys_display.sn_customerservice_case.contact")).click();
            // driver.findElement(By.id("status.sn_customerservice_case.contact")).click();
            // driver.switchTo().alert().accept();
            Thread.sleep(2000);

            Select drpm = new Select(driver.findElement(By.name("sn_customerservice_case.category")));
            drpm.selectByVisibleText("General Assistance");
            Thread.sleep(2000);
            Select drpn = new Select(driver.findElement(By.name("sn_customerservice_case.subcategory")));
            drpn.selectByVisibleText("General Inquiry");

            lab.assignmentGroupSelector(prefix, driver);

            Select drp = new Select(driver.findElement(By.id("sn_customerservice_case.contact_type")));
            drp.selectByVisibleText("Email");
            Select drp1 = new Select(driver.findElement(By.id("sn_customerservice_case.u_case_urgency")));
            drp1.selectByVisibleText("4 - Low");
            
            String caseNumber = driver.findElement(By.id("sn_customerservice_case.number")).getAttribute("value");
            lab.createAttachment(mail, caseNumber, subFinal);/**++subject for file name */
            // driver.get("https://products.aspose.app/email/conversion/eml-to-msg");
            // driver.findElement(By.xpath("//*[@id='UploadFileInput-1']")).sendKeys(D:\\"+caseNumber+"-"+subFinal+".eml);

            mail.setFlag(Flag.SEEN, false);/* Unread message */
            

            System.out.println(Strings.repeat("#", 140));            
            System.out.println("MESSAGE : " +"\n");
            System.out.println("From :" + from +"\n"+ "To   :" + to +"\n" + "Sub  :" + subject +"\n"+ "Size :" + size +"\n"+ "Date :");
            System.out.println(caseNumber);
            // System.out.println(clientDeployment);
            // System.out.println(emailDomain);
            

            // System.out.println(content);

            // System.out.println(emailbody);
            // System.out.println("Has Attachments: " + lab.hasAttachments(mail));
            
            // upload file
            driver.findElement(By.id("header_add_attachment")).click();
            // WebElement attachButton = driver.findElement(By.id("attachFile"));
            // attachButton.sendKeys("D:\\Users\\F3C3XRT\\Documents\\"+caseNumber+".doc");
            WebElement attachButton2 = driver.findElement(By.id("attachFile"));
            attachButton2.sendKeys("D:\\Users\\F3C3XRT\\Documents\\"+caseNumber+".zip");
            WebDriverWait wait = new WebDriverWait(driver, 60);
            wait.until(ExpectedConditions.visibilityOfElementLocated(By.xpath("//*[@id='attachment_table_body']")));
            // Thread.sleep(5000);
            // // driver.switchTo().alert().accept();
            driver.findElement(By.id("attachment_closemodal")).click();
            driver.findElement(By.id("sysverb_update_and_stay")).click();
            // // driver.close();
            driver.switchTo().window(homeWindowHandleID);

            driver.findElement(By.xpath("//*[@id='8f0ab177c342310015519f2974d3ae12']/div/div")).click();

            
            // System.out.println("Subject: " + mail.getSubject());
            // System.out.println("From: " + mail.getFrom()[0]);
            // System.out.println("To: " + mail.getAllRecipients()[0]);
            // System.out.println("Date: " + mail.getReceivedDate());
            // System.out.println("Size: " + mail.getSize());
            // System.out.println("Flags: " + mail.getFlags());
            // System.out.println("ContentType: " + mail.getContentType());
            // System.out.println("Body: \n" + getEmailBody(mail));
            
            // }
        }
    }
    }

