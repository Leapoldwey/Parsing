package org.example.SendEmail;

import javax.activation.DataHandler;
import javax.activation.DataSource;
import javax.activation.FileDataSource;
import javax.mail.*;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;
import java.util.Properties;

public class SendEmail {
    public static void send() {
        final String from = "mihail_mihailow@list.ru";
        String to = "mihail_mihailow@list.ru";
        String host = "smtp.mail.ru";
        String smtpPort = "465";

        Properties properties = new Properties();
        properties.put("mail.smtp.host", host);
        properties.put("mail.smtp.port", smtpPort);
        properties.put("mail.smtp.ssl.enable", "true");
        properties.put("mail.smtp.auth", "true");

        Session session = Session.getInstance(properties, new Authenticator() {
            @Override
            protected PasswordAuthentication getPasswordAuthentication() {
                return new PasswordAuthentication(from, "1Gc2QAMGbjLZA7yLdfND");
            }
        });

        session.setDebug(true);

        try {
            MimeMessage m = new MimeMessage(session);
            m.setFrom(new InternetAddress(from));
            m.addRecipients(Message.RecipientType.TO, new InternetAddress(to).toString());
            m.setSubject("Курс валют");

            BodyPart messageBodyPart = new MimeBodyPart();
            messageBodyPart.setText("");

            Multipart multipart = new MimeMultipart();
            multipart.addBodyPart(messageBodyPart);

            messageBodyPart = new MimeBodyPart();

            String filename = "C:\\Users\\User\\IdeaProjects\\Parsing\\excel.xlsx";

            DataSource source = new FileDataSource(filename);

            messageBodyPart.setDataHandler(new DataHandler(source));
            messageBodyPart.setFileName(filename);

            multipart.addBodyPart(messageBodyPart);
            m.setContent(multipart);

            Transport.send(m);

        } catch (Exception e) {
            e.printStackTrace();
        }

    }
}
