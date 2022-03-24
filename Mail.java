package Test;

import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.BodyPart;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.NoSuchProviderException;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

/**
 * Mail
 */
public class Mail {

  public static void main(String[] args) throws MessagingException {
    String username = "1dc\\svc-tcoe-auto-tester";
    String password = "xjsE@k328";
    
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
    
  }
}