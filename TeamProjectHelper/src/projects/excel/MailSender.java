/**
 * 
 * @author : Aslam
 * @Dependancies : ExcelHandler.java
 * @Details : This class is responsible for handling the email operations.
 * 		like creating messages and session and then finally sends the mail to
 * 		the recipients. The attachment(excel file) which have to be included in 
 * 		the mail is provided by ExcelHandler.java class
 *  
 */

package projects.excel;

import java.util.Properties;

import javax.mail.Authenticator;
import javax.mail.Message;
import javax.mail.Multipart;
import javax.mail.PasswordAuthentication;
import javax.mail.Session;
import javax.mail.Transport;
import javax.mail.internet.InternetAddress;
import javax.mail.internet.MimeBodyPart;
import javax.mail.internet.MimeMessage;
import javax.mail.internet.MimeMultipart;

public class MailSender {

	private final String fromEmail = "myEmail@gmail.com";
	private final String password = "myPassword";
	private String toEmail;
	private String attachment;

	private Session session;
	private Message message;

	/**
	 * Parameters : Accepts the sender's email and the attachment that has to be
	 * sent along with the email
	 * 
	 * Job : 
	 * 		1. initializes toEmail and attachment
	 * 		2. calls the neccessary methods in order to process the data
	 */
	public void sendMail(String toEmail, String attachment) {
		this.toEmail = toEmail;
		this.attachment = attachment;
		System.out.println(toEmail + " " + attachment);

		makeConnection();
		createMessage();
		send();

		System.out.println("--------------------------------------------------");
	}

	/**
	 *  sets up the connection with sever and creates a Session
	 */
	private void makeConnection() {

		Properties properties = new Properties();

		properties.put("mail.smtp.auth", "true");
		properties.put("mail.smtp.starttls.enable", "true");
		properties.put("mail.smtp.host", "smtp.gmail.com");
		properties.put("mail.smtp.port", "587");

		// runs in another thread
		session = Session.getInstance(properties, new Authenticator() {
			@Override
			protected PasswordAuthentication getPasswordAuthentication() {
				System.out.println("connecting to server ... ");
				return new PasswordAuthentication(fromEmail, password);
			}
		});
	}

	/**
	 *  creates the Message object to be sent via mail
	 */
	private void createMessage() {

		message = new MimeMessage(session);
		
		try {

			message.setFrom(new InternetAddress(fromEmail));
			message.setRecipient(Message.RecipientType.TO, new InternetAddress(toEmail));
			message.setSubject("Results Have been published");

			Multipart multipart = new MimeMultipart();

			// text body
			MimeBodyPart text1 = new MimeBodyPart();
			text1.setText(getTextBody());

			// attachment
			MimeBodyPart file = new MimeBodyPart();
			file.attachFile(attachment);

			multipart.addBodyPart(file, 0);
			multipart.addBodyPart(text1, 1);

			message.setContent(multipart);

		} catch (Exception e) {
			System.out.println("SOME ERROR OCCURED !!!");
			e.printStackTrace();
		}
	}

	/**
	 *  sends the message using Transport.send()
	 */
	private boolean send() {

		try {
			Transport.send(message);
			System.out.printf("MESSAGE SUCCESSFULLY SENT to <%s> !!!\n", toEmail);
			return true;
		} catch (Exception e) {

			System.out.println(e.getMessage());
			e.printStackTrace();
			return false;
		}
	}

	/**
	 * a helper method for createMessage() for creating the text body
	 */
	private String getTextBody() {

		// splits the email so that we get the person's name
		String[] splittedArray = toEmail.split("@");
		String name = splittedArray[0];

		String message = "Dear Mr. " + name + ",\n" + "\tGreetings to you. I hope you are at the best of your health. "
				+ "\nWelcome to my GitHub account - https://github.com/Aslam-web/Team-Project-Codes/tree/master/TeamProjectHelper/src/projects/excel"

				+ "\n\n\nThanks & Regards" + "\nMr M.N Aslam," + "\nJAVA developer Trainer,"
				+ "\nHaaris Infotech Institutions," + "\nEmail : aslam1qqqq@gmail.com," + "\nPhone: +91 63799 71782.";

		return message;
	}
}