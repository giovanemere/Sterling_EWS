package com.aossas.o365;


import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;

import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.WellKnownFolderName;
import microsoft.exchange.webservices.data.core.enumeration.search.FolderTraversal;
import microsoft.exchange.webservices.data.core.enumeration.search.SortDirection;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceLocalException;
import microsoft.exchange.webservices.data.core.exception.service.local.ServiceVersionException;
import microsoft.exchange.webservices.data.core.service.folder.Folder;
import microsoft.exchange.webservices.data.core.service.item.EmailMessage;
import microsoft.exchange.webservices.data.core.service.item.Item;
import microsoft.exchange.webservices.data.core.service.schema.EmailMessageSchema;
import microsoft.exchange.webservices.data.core.service.schema.FolderSchema;
import microsoft.exchange.webservices.data.core.service.schema.ItemSchema;
import microsoft.exchange.webservices.data.credential.ExchangeCredentials;
import microsoft.exchange.webservices.data.credential.WebCredentials;
import microsoft.exchange.webservices.data.property.complex.Attachment;
import microsoft.exchange.webservices.data.property.complex.AttachmentCollection;
import microsoft.exchange.webservices.data.property.complex.FileAttachment;
import microsoft.exchange.webservices.data.property.complex.FolderId;
//import microsoft.exchange.webservices.data.property.complex.ItemAttachment;
import microsoft.exchange.webservices.data.property.complex.ItemId;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.FolderView;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;

public class UtilMail {

	private ExchangeService service;
	private String path;
	private static final Log log = LogFactory.getLog(ReadInbox.class);
	 
	public UtilMail() {
		service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
	}
	
	public UtilMail(String user, String password) throws Exception  {
		this();
		ExchangeCredentials credentials = new WebCredentials(user, password);
		service.setCredentials(credentials);
		
			//service.setTraceEnabled(true);
			service.autodiscoverUrl(user, new RedirectionUrlCallback());
			//service.setUrl(new URI("https://outlook.office365.com/EWS/Exchange.asmx"));
			Folder inbox = Folder.bind(service, WellKnownFolderName.Inbox);
			System.out.println("Total Correos: " + inbox.getTotalCount());
		
	}
	
	public UtilMail(String user, String password, String path) throws Exception {
		this(user, password);
		this.setPath(path);
	}
	
	public String getPath() {
		return path;
	}

	public void setPath(String path) {
		this.path = path;
	}
	
	public static class RedirectionUrlCallback implements IAutodiscoverRedirectionUrl {
        public boolean autodiscoverRedirectionUrlValidationCallback(String redirectionUrl) {
          return redirectionUrl.toLowerCase().startsWith("https://");
        }
    }
	
	public ExchangeService getService() {
		return service;
	}
	
	private Folder getFolder(String folderPath) throws Exception {
		Folder folder = Folder.bind(service, WellKnownFolderName.Root);
		FolderView view = new FolderView(100);
		//PropertySet properties =  new PropertySet(BasePropertySet.IdOnly, FolderSchema.DisplayName);
		view.setTraversal(FolderTraversal.Deep);
		// Return only folders that contain items.
		
		String[] folderTree = folderPath.split("/");
		FolderId folderId = FolderId.getFolderIdFromWellKnownFolderName(WellKnownFolderName.Root);
		SearchFilter searchFilter;
		for(String folderName : folderTree) {			
			searchFilter = new SearchFilter.IsEqualTo(FolderSchema.DisplayName, folderName);
			FindFoldersResults findFolderResults = service.findFolders(folderId, searchFilter, view);
			//service.loadPropertiesForItems(findFolderResults.iterator(), properties);
			int resultCount = findFolderResults.getTotalCount();
			if(resultCount == 1) {
				folder = findFolderResults.getFolders().get(0);
				folderId = folder.getId();
			}else {
				throw new Exception(folderName + "not found!");
			}		
			
		}
		
		return folder;
	}
	


	public List<String> getMailList(String folderPath, String mailFrom) throws Exception {
		Folder folder = getFolder(folderPath);		
		return getMailList(folder, mailFrom);
	}
	
	private List<String> getMailList(Folder folder, String mailFrom) {
		List<String> mailList = new ArrayList<String>();
		try {
			
		    ItemView view = new ItemView(Integer.MAX_VALUE);
		    view.getOrderBy().add(ItemSchema.DateTimeReceived, SortDirection.Ascending);
		    
		    SearchFilter searchFilter = new SearchFilter.IsEqualTo(EmailMessageSchema.From, mailFrom); 
		    
		    FindItemsResults<Item> results = service.findItems(folder.getId(), searchFilter, view);
		    service.loadPropertiesForItems(results, new PropertySet(BasePropertySet.FirstClassProperties,EmailMessageSchema.Id));
		    for (Item item : results) {
		    	ItemId itemId= item.getId();
		    	mailList.add(itemId.getUniqueId());
		    }		   
		            
		} catch (Exception e) {
		    log.error("Error ", e);
			//e.printStackTrace();
		}	
		
		return mailList;
	
	}
	private Folder createFolderProcesados() throws Exception {
		Folder folder = new Folder(service);
		folder.setDisplayName("Procesados");
		folder.save(WellKnownFolderName.Inbox);
		
		return folder;
	}
	
	private Folder getFolderProcesados() throws Exception {
		Folder folder;
		try{
		  folder = getFolder("Procesados");
		} catch(Exception e) {
			folder =  createFolderProcesados();
		}		
		return folder;
	}
	
	public HashMap<String, HashMap<String, String>> getAttachmentsFromMail(String emailId) throws IllegalArgumentException, ServiceVersionException, ServiceLocalException, Exception {
		ItemId id = new ItemId(emailId);		
		return getAttachmentsFromMail(id);
	}
	
	private HashMap<String, HashMap<String, String>> getAttachmentsFromMail(ItemId mailId) throws Exception  {
		HashMap<String, HashMap<String, String>> attachments;	
		Folder destinationFolder = getFolderProcesados();
	    FolderId destinationFolderId = destinationFolder.getId();
	    
        //Item itm = Item.bind(service, mailId, new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments));
        EmailMessage emailMessage = EmailMessage.bind(service, mailId, new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments));
        //System.out.println(emailMessage.getBody().toString());
        attachments = getAttachments(emailMessage);
        
        log.info(emailMessage.getSubject() + " Files: " +  attachments.size());
        emailMessage.move(destinationFolderId);
		return attachments;
	}
	

	private HashMap<String, HashMap<String, String>> getAttachments(EmailMessage emailMessage) throws IllegalArgumentException, ServiceVersionException, ServiceLocalException, Exception {
		HashMap<String, HashMap<String, String>> attachments = new HashMap<String, HashMap<String, String>>();    
        
		if (emailMessage.getHasAttachments() || emailMessage.getAttachments().getCount() > 0) {
		        //get all the attachments
		        AttachmentCollection attachmentsCol = emailMessage.getAttachments();

		        log.info("File Count: " +attachmentsCol.getCount());
		        
		        //loop over the attachments
		        for (int i = 0; i < attachmentsCol.getCount(); i++) {
		            Attachment attachment = attachmentsCol.getPropertyAtIndex(i);
		            log.debug("Starting to process attachment "+ attachment.getName());

		               //FileAttachment - Represents a file that is attached to an email item
		                if (attachment instanceof FileAttachment || attachment.getIsInline()) {
		                	
		                    attachments.putAll(extractFileAttachments(attachment));

		                } 
		            }
		        
		    } else {
		        log.debug("Email message does not have any attachments.");
		    	//System.out.println("Email message does not have any attachments.");
		    }
       
        
        return attachments;
    }
	
	 private HashMap<String, HashMap<String, String>> extractFileAttachments(Attachment attachment) {
			//Extract File Attachments
	    	HashMap<String, HashMap<String, String>> attachments = new HashMap<String, HashMap<String, String>>();
	    	
	    	 HashMap<String, String> fileAttachments = new  HashMap<String, String>();
			try {
			    FileAttachment fileAttachment = (FileAttachment) attachment;
			    // if we don't call this, the Content property may be null.
			    fileAttachment.load();

			    //extract the attachment content, it's not base64 encoded.
			    byte[] attachmentContent;
			    attachmentContent = fileAttachment.getContent();

			    if (attachmentContent != null && attachmentContent.length > 0) {

			        //check the size
			        int attachmentSize = attachmentContent.length;

			        //check if the attachment is valid
			        //ValidateEmail.validateAttachment(fileAttachment, properties, emailIdentifier, attachmentSize);

			        fileAttachments.put(UtilConstants.ATTACHMENT_SIZE, String.valueOf(attachmentSize));

			        //get attachment name
			        String fileName = fileAttachment.getName() == null ? fileAttachment.getFileName() : fileAttachment.getName();
			        fileAttachments.put(UtilConstants.ATTACHMENT_NAME, fileName);

			        String mimeType = fileAttachment.getContentType();
			        fileAttachments.put(UtilConstants.ATTACHMENT_MIME_TYPE, mimeType);

			        log.info("File Name: " + fileName + "  File Size: " + attachmentSize);
			        //System.out.println("File Name: " + fileName + "  File Size: " + attachmentSize);


			        if (attachmentContent != null && attachmentContent.length > 0) {
			            //convert the content to base64 encoded string and add to the collection.
			        	if(path != null && !"".equals(path))
			        	  UtilFunctions.writeByte(path + fileName, attachmentContent);
			            String base64Encoded = new String(attachmentContent);//UtilFunctions.encodeToBase64(attachmentContent);
			            fileAttachments.put(UtilConstants.ATTACHMENT_CONTENT, base64Encoded);
			        }

			    }
			}catch (Exception e) {
				log.error("Error in File extraction", e);
				//e.printStackTrace();
			}
		    return attachments;
	    }
	 
	 public byte[] getAttachmentContent(String emailId, String fileAttachmentId) {
		 EmailMessage emailMessage;
		 byte[] content = null;
		try {
			emailMessage = EmailMessage.bind(service, new ItemId(emailId),new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments));
			content = getAttachmentContent(emailMessage, fileAttachmentId);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		 return content;
	 }
	 
	 private byte[] getAttachmentContent(EmailMessage emailMessage, String fileAttachmentId) {
		 byte[] content = null;
		 //PropertySet propertySet = new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments);
		 //Item item = Item.bind(service, fileAttachmentId, propertySet);
		 try {
			if (emailMessage.getHasAttachments() || emailMessage.getAttachments().getCount() > 0) {
				 AttachmentCollection attachmentsCol = emailMessage.getAttachments();
				 for(Attachment attachment : attachmentsCol)
					 if (fileAttachmentId.equals(attachment.getId()) && ( attachment instanceof FileAttachment || attachment.getIsInline())) {
						 FileAttachment fileAttachment = (FileAttachment) attachment;
						 fileAttachment.load();
						 content  = fileAttachment.getContent();
			         } 
			 }
		} catch (ServiceVersionException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (ServiceLocalException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		 return content;
	 }
			
		
	
}


