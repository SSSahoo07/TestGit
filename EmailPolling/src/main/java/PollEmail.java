import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Properties;

import javax.mail.Folder;
import javax.mail.Message;
import javax.mail.MessagingException;
import javax.mail.Multipart;
import javax.mail.Session;
import javax.mail.Store;
import javax.mail.URLName;
import javax.mail.internet.MimeBodyPart;

import com.itextpdf.html2pdf.HtmlConverter;

public class PollEmail {

	private Session session;
	private Store store;
	private Folder folder;

	// hardcoding protocol and the folder
	// it can be parameterized and enhanced as required
	private String protocol = "imaps";
	private String file = "INBOX";

	public PollEmail() {

	}

	public boolean isLoggedIn() {
		return store.isConnected();
	}

	/**
	 * to login to the mail host server
	 */
	public void login(String host, String username, String password)
			throws Exception {
		URLName url = new URLName(protocol, host, 993, file, username, password);

		if (session == null) {
			Properties props = null;
			try {
				props = System.getProperties();
			} catch (SecurityException sex) {
				props = new Properties();
			}
			session = Session.getInstance(props, null);
		}
		store = session.getStore(url);
		store.connect();
		folder = store.getFolder(url);

		folder.open(Folder.READ_WRITE);
		Message[] messages=folder.getMessages();
		for(Message message:messages)
		{
			System.out.println(message.getSubject()+":: Message Type-->"+message.getContentType());
			Object content = message.getContent();
	        StringBuilder result = new StringBuilder();
	        String contentType=message.getContentType();
	        if (contentType.startsWith("text/plain")) {
	            result.append(content);
	            File messageFile=new File("D:\\Email_Polling\\"+message.getSubject().replaceAll(":", "_")+"\\");
		        if(!messageFile.exists())
		        	{
		        	System.out.println(messageFile.mkdirs());
		        	}
		        FileOutputStream fout=new FileOutputStream(messageFile+File.separator+message.getSubject().replaceAll(":", "_")+".txt");
		        fout.write(result.toString().getBytes());
		        fout.close();
	        } else if (content instanceof Multipart) {
	        	File messageFile=new File("D:\\Email_Polling\\"+message.getSubject().trim().replaceAll(":", "_")+"\\");
		        if(!messageFile.exists())
		        	{
		        	System.out.println(messageFile.mkdirs());
		        	}
	            Multipart parts = (Multipart) content;
	            for (int i = 0; i < parts.getCount(); i++) {
	                MimeBodyPart part = (MimeBodyPart) parts.getBodyPart(i);
	                System.out.println(part.getContentType());
	                if (part.getContentType().startsWith("multipart/")) {
	                	 Multipart body = (Multipart) part.getContent();
	                	 for(int j=0;j<body.getCount();j++)
						{
	                		 MimeBodyPart bodyPart = (MimeBodyPart) body.getBodyPart(j);
	                		 System.out.println("bodyPart  "+bodyPart.getContentType());
							if (bodyPart.getContentType().startsWith("text/plain")) {
								result.append(bodyPart.getContent());
								
								FileOutputStream fout = new FileOutputStream(messageFile + File.separator
										+ message.getSubject().trim().replaceAll(":", "_") + ".txt");
								fout.write(result.toString().getBytes());
								fout.close();
							}
							if (bodyPart.getContentType().startsWith("text/html")) {
								
								createHtml(messageFile,message,bodyPart);
								
							}
						}
	                }
	                else if(part.getContentType().startsWith("application/pdf"))
	                {
	                	part.saveFile(messageFile+"\\"+part.getFileName());
	                }
	                else if(part.getContentType().startsWith("text/html"))
	                {
	                	createHtml(messageFile,message,part);
	                	
	                }
	                else if(part.getContentType().startsWith("text/plain"))
	                {
	                	FileOutputStream fout = new FileOutputStream(messageFile + File.separator
								+ message.getSubject().trim().replaceAll(":", "_") + ".txt");
						fout.write(part.getContent().toString().getBytes());
						fout.close();
	                }
	            }
	            
		        
	        }
	        //System.out.println(result);
	      
	        System.out.println("*******************************************************************************************************");
		}
		
		System.out.println(folder.getUnreadMessageCount());
	}

	private void createHtml(File messageFile, Message message, MimeBodyPart bodyPart) throws MessagingException, IOException {
		FileOutputStream fout = new FileOutputStream(messageFile + File.separator
				+ message.getSubject().trim().replaceAll(":", "_") + ".html");
		fout.write(bodyPart.getContent().toString().getBytes());
		fout.close();
		String htmlFile=messageFile + File.separator+ message.getSubject().trim().replaceAll(":", "_") + ".html";
		String pdfFile=messageFile + File.separator+ message.getSubject().trim().replaceAll(":", "_") + ".pdf";
		HtmlConverter.convertToPdf(new File(htmlFile), new File(pdfFile));
		
	}

	/**
	 * to logout from the mail host server
	 */
	public void logout() throws MessagingException {
		folder.close(false);
		store.close();
		store = null;
		session = null;
	}

	
	public int getMessageCount() {
		int messageCount = 0;
		try {
			messageCount = folder.getMessageCount();
		} catch (MessagingException me) {
			me.printStackTrace();
		}
		return messageCount;
	}

	public Message[] getMessages() throws MessagingException {
		return folder.getMessages();
	}
	public static void main(String[] args) {
		
		PollEmail pollEmail=new PollEmail();
		try {
			pollEmail.login("imap-mail.outlook.com", "sonalissahoo@outlook.com", "Kahna@16jan");
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
	}

	

}
