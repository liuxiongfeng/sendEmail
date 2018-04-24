import com.mysql.jdbc.Messages;
import com.sun.xml.internal.messaging.saaj.packaging.mime.MessagingException;
import org.apache.poi.util.StringUtil;

import java.io.*;
import java.net.Authenticator;
import java.util.Properties;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.*;
import java.util.Properties;

public class EmailSender extends Authenticator {

    static String subject = "";
    static String text = "";
    static String filepath1 = "";
    static String filepath2 = "";
    static String picpath = "";

    public void setParams(String subject1, String text1, String filepath10,String filepath20,String picpath1){
        subject = subject1;
        text = text1;
        filepath1 = filepath10;
        filepath2 = filepath20;
        picpath = picpath1;
    }
    /**
     *
     *
     * @throws MessagingException
     */
    public void sendMail(final String from, String to) throws MessagingException {

        try {
            String host = "smtp." + from.split("@")[1];

            final String password = "1991100126";

            Properties props = new Properties();
            props.put("mail.smtp.auth", "true");
            props.put("mail.smtp.starttls.enable", "true");
            props.put("mail.smtp.host", host);
            props.put("mail.smtp.port", "25");

            Session session = Session.getInstance(props,
                    new javax.mail.Authenticator() {
                        protected PasswordAuthentication getPasswordAuthentication() {
                            return new PasswordAuthentication(from, password);
                        }
                    });

            // Create a default MimeMessage object.
            Message message = new MimeMessage(session);

            // Set From: header field of the header.
            //设置自定义发件人昵称
            String nick="";
            try {
                nick=javax.mail.internet.MimeUtility.encodeText("大帅哥");
            } catch (UnsupportedEncodingException e) {
                e.printStackTrace();
            }
            message.setFrom(new InternetAddress(nick+" <"+from+">"));

            // Set To: header field of the header.
            message.setRecipients(Message.RecipientType.TO,
                    InternetAddress.parse(to));

            // Set Subject: header field
            message.setSubject("群发邮件");

            // 向multipart对象中添加邮件的各个部分内容，包括文本内容和附件
            Multipart multipart = new MimeMultipart();

            // 添加邮件正文
            BodyPart contentPart = new MimeBodyPart();
            contentPart.setContent(text, "text/html;charset=UTF-8");
            multipart.addBodyPart(contentPart);

            // 添加附件的内容
            if (filepath1 != null) {
                BodyPart attachmentBodyPart = new MimeBodyPart();
                DataSource source = new FileDataSource(filepath1);
                attachmentBodyPart.setDataHandler(new DataHandler(source));

                // 网上流传的解决文件名乱码的方法，其实用MimeUtility.encodeWord就可以很方便的搞定
                // 这里很重要，通过下面的Base64编码的转换可以保证你的中文附件标题名在发送时不会变成乱码
                //sun.misc.BASE64Encoder enc = new sun.misc.BASE64Encoder();
                //messageBodyPart.setFileName("=?GBK?B?" + enc.encode(attachment.getName().getBytes()) + "?=");

                //MimeUtility.encodeWord可以避免文件名乱码
                String filesp = filepath1.trim();
                String temp[] = filesp.split("\\\\");
                attachmentBodyPart.setFileName(temp[temp.length-1]);
                multipart.addBodyPart(attachmentBodyPart);
            }

            if (filepath2 != null) {
                BodyPart attachmentBodyPart = new MimeBodyPart();
                DataSource source = new FileDataSource(filepath2);
                attachmentBodyPart.setDataHandler(new DataHandler(source));

                // 网上流传的解决文件名乱码的方法，其实用MimeUtility.encodeWord就可以很方便的搞定
                // 这里很重要，通过下面的Base64编码的转换可以保证你的中文附件标题名在发送时不会变成乱码
                //sun.misc.BASE64Encoder enc = new sun.misc.BASE64Encoder();
                //messageBodyPart.setFileName("=?GBK?B?" + enc.encode(attachment.getName().getBytes()) + "?=");

                //MimeUtility.encodeWord可以避免文件名乱码
                String filesp = filepath2.trim();
                String temp[] = filesp.split("\\\\");
                attachmentBodyPart.setFileName(temp[temp.length-1]);
                multipart.addBodyPart(attachmentBodyPart);
            }

            // 将multipart对象放到message中
            message.setContent(multipart);
            // 保存邮件
            message.saveChanges();

            //设置邮件的内容体
            String cd = "<!DOCTYPE html PUBLIC '-//W3C//DTD XHTML 1.0 Transitional//EN' 'http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd'>"
                    + "<html xmlns='http://www.w3.org/1999/xhtml'>"
                    + "<head>"
                    + "<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />"
                    + "<title>无标题文档</title>"
                    + "</head>"
                    + "<body>"
                    +"邮件发送成功"
                    + "<div data-spm='100' class='copyright-100 '>"
                    + "<div class='y-row copyright' data-spm='25'>"
                    + "<h4 align='center' class='big'>"
                    + "<a target='blank' href='http://www.wondersgroup.com/about-wd/201412310639133820.html'>关于我们</a>&nbsp;&nbsp;"
                    + "<a target='blank' href='http://www.wondersgroup.com/news'>服务公告</a>&nbsp;&nbsp;"
                    + "<a target='blank' href='http://www.wondersgroup.com/contact'>联系我们</a>&nbsp;&nbsp;"
                    + "</h4><h5> &copy; 1995 - 2015 WONDERS INFORMATION CO., LTD. ALL RIGHTS RESERVED. 万达信息股份有限公司版权所有 沪ICP备05002296号 </h5></div></div></body></html>";


           // message.setContent(cd,"text/html;charset=UTF-8");
            // Send message
            Transport.send(message);

            System.out.println("Sent message successfully");
        } catch (javax.mail.MessagingException e) {
            e.printStackTrace();
        }

    }

}
