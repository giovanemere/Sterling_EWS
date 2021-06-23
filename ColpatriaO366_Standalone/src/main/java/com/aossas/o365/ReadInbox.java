package com.aossas.o365;

import java.util.HashMap;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;


import microsoft.exchange.webservices.data.autodiscover.IAutodiscoverRedirectionUrl;
import microsoft.exchange.webservices.data.core.ExchangeService;
import microsoft.exchange.webservices.data.core.PropertySet;
import microsoft.exchange.webservices.data.core.enumeration.misc.ExchangeVersion;
import microsoft.exchange.webservices.data.core.enumeration.property.BasePropertySet;
import microsoft.exchange.webservices.data.core.enumeration.property.BodyType;
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
import microsoft.exchange.webservices.data.property.complex.ItemAttachment;
import microsoft.exchange.webservices.data.search.FindFoldersResults;
import microsoft.exchange.webservices.data.search.FindItemsResults;
import microsoft.exchange.webservices.data.search.FolderView;
import microsoft.exchange.webservices.data.search.ItemView;
import microsoft.exchange.webservices.data.search.filter.SearchFilter;


public class ReadInbox {
	private ExchangeService service;
	private String path;
	 private static final Log log = LogFactory.getLog(ReadInbox.class);
	public ReadInbox() {
		service = new ExchangeService(ExchangeVersion.Exchange2010_SP2);
	}
	
	public ReadInbox(String user, String password) throws Exception  {
		this();
		ExchangeCredentials credentials = new WebCredentials(user, password);
		service.setCredentials(credentials);
		
			//service.setTraceEnabled(true);
			service.autodiscoverUrl(user, new RedirectionUrlCallback());
			//service.setUrl(new URI("https://outlook.office365.com/EWS/Exchange.asmx"));
			Folder inbox = Folder.bind(service, WellKnownFolderName.Inbox);
			System.out.println("Total Correos: " + inbox.getTotalCount());
		
	}
	
	public ReadInbox(String user, String password, String path) throws Exception {
		this(user, password);
		this.setPath(path);
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
	
	
	
	public void readMailInbox() throws Exception {
		Folder folder = Folder.bind(service, WellKnownFolderName.Inbox);
		readMailFolder(folder);
		
	}
	
	private void getAttachmentsFromMail(FindItemsResults<Item> results) throws Exception  {
		HashMap<String, HashMap<String, String>> attachments;	
		Folder destinationFolder = getFolderProcesados();
	    FolderId destinationFolderId = destinationFolder.getId();
	    
        for (Item item : results) {
	        Item itm = Item.bind(service, item.getId(), new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments));
	        EmailMessage emailMessage = EmailMessage.bind(service, itm.getId(), new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments));
	        //System.out.println(emailMessage.getBody().toString());
	        attachments = getAttachments(service, emailMessage);
	        log.info(emailMessage.getSubject() + " Files: " +  attachments.size());
	        emailMessage.move(destinationFolderId);
        }
		
	}
	

	public void readMailFoder(String folderPath, String mailFrom) throws Exception {
		Folder folder = getFolder(folderPath);		
		readMailFolder(folder, mailFrom);
	}
	
	public void readMailFolder(Folder folder, String mailFrom) {
		try {
			
		    ItemView view = new ItemView(Integer.MAX_VALUE);
		    view.getOrderBy().add(ItemSchema.DateTimeReceived, SortDirection.Ascending);
		    
		    SearchFilter searchFilter = new SearchFilter.IsEqualTo(EmailMessageSchema.From, mailFrom); 
		    
		    FindItemsResults<Item> results = service.findItems(folder.getId(), searchFilter, view);
		    service.loadPropertiesForItems(results, new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments));
		    
		    getAttachmentsFromMail(results);
		            
		} catch (Exception e) {
		    log.error("Error ", e);
			//e.printStackTrace();
		}	
	
	}
	
	public void readMailFolder(Folder folder) {
		try {
			
			
		    ItemView view = new ItemView(Integer.MAX_VALUE);
		    view.getOrderBy().add(ItemSchema.DateTimeReceived, SortDirection.Ascending);
		    
		    FindItemsResults<Item> results = service.findItems(folder.getId(),view);
		    service.loadPropertiesForItems(results, new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments));
		    
		    getAttachmentsFromMail(results);
		            
		} catch (Exception e) {
		    log.error("Error ", e);
			//e.printStackTrace();
		}	
	
	}
	
	public String getFileExtension(String name) {
		String[] tokens = name.split(".");
		return tokens.length > 1 ? tokens[tokens.length -1] : "";
	}

	private HashMap<String, HashMap<String, String>> getAttachments(ExchangeService service, EmailMessage emailMessage) throws IllegalArgumentException, ServiceVersionException, ServiceLocalException, Exception {
		HashMap<String, HashMap<String, String>> attachments = new HashMap<String, HashMap<String, String>>();    
        
		if (emailMessage.getHasAttachments() || emailMessage.getAttachments().getCount() > 0) {
			PropertySet properties =  new PropertySet(BasePropertySet.FirstClassProperties, EmailMessageSchema.Attachments);
		        //get all the attachments
		        AttachmentCollection attachmentsCol = emailMessage.getAttachments();

		        log.info("File Count: " +attachmentsCol.getCount());
		        
		        //loop over the attachments
		        for (int i = 0; i < attachmentsCol.getCount(); i++) {
		            Attachment attachment = attachmentsCol.getPropertyAtIndex(i);
		            log.debug("Starting to process attachment "+ attachment.getName());

		               //FileAttachment - Represents a file that is attached to an email item
		                if (attachment instanceof FileAttachment || attachment.getIsInline()) {
		                	
		                    attachments.putAll(extractFileAttachments(attachment,properties));

		                } else if (attachment instanceof ItemAttachment) { //ItemAttachment - Represents an Exchange item that is attached to another Exchange item.
		                	HashMap<String, Item[]> appendedBody = new HashMap<String, Item[]>();
		                    attachments.putAll(extractItemAttachments(service, attachment, properties, appendedBody));
		                }
		            }
		        
		    } else {
		        log.debug("Email message does not have any attachments.");
		    	//System.out.println("Email message does not have any attachments.");
		    }
       
        
        return attachments;
    }
	
	public HashMap<String, HashMap<String, String>> extractFileAttachments(Attachment attachment,PropertySet properties) {
		return extractFileAttachments(attachment,properties,attachment.getId());
	}
	
    private HashMap<String, HashMap<String, String>> extractFileAttachments(Attachment attachment,PropertySet properties, String emailIdentifier) {
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
		        ValidateEmail.validateAttachment(fileAttachment, properties,
		                emailIdentifier, attachmentSize);

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
		
	private HashMap<String, HashMap<String, String>> extractItemAttachments(ExchangeService service,Attachment attachment,PropertySet properties, HashMap<String, Item[]> appendedBody) throws Exception {
		return extractItemAttachments(service, attachment, properties, appendedBody, "");
	}
	
	private HashMap<String, HashMap<String, String>> extractItemAttachments(ExchangeService service,Attachment attachment,PropertySet properties, HashMap<String, Item[]> appendedBody, String emailIdentifier) throws Exception {
	HashMap<String, HashMap<String, String>> itemAttachments = new HashMap<String, HashMap<String, String>>();
		//Extract Item Attachment
		try {
		    ItemAttachment itemAttachment = (ItemAttachment) attachment;

		    PropertySet propertySet = new PropertySet(
		            BasePropertySet.FirstClassProperties, ItemSchema.Attachments, 
		            ItemSchema.Body, ItemSchema.Id, ItemSchema.DateTimeReceived,
		            EmailMessageSchema.DateTimeReceived, EmailMessageSchema.Body);

		    itemAttachment.load();
		    propertySet.setRequestedBodyType(BodyType.Text);

		    Item item = itemAttachment.getItem();
		    Item[] eBody;
		    eBody = appendItemBody(item, appendedBody.get(UtilConstants.BODY_CONTENT));

		    appendedBody.put(UtilConstants.BODY_CONTENT, eBody);

		    /*
		     * We need to check if Item attachment has further more
		     * attachments like .msg attachment, which is an outlook email
		     * as attachment. Yes, we can attach an email chain as
		     * attachment and that email chain can have multiple
		     * attachments.
		     */
		    AttachmentCollection childAttachments = item.getAttachments();
		    //check if not empty collection. move on
		    if (childAttachments != null && childAttachments.getCount() > 0) {

		        for (Attachment childAttachment : childAttachments) {

		            if (childAttachment instanceof FileAttachment) {

		                itemAttachments.putAll(extractFileAttachments(childAttachment, properties, emailIdentifier));

		            } else if (childAttachment instanceof ItemAttachment) {

		                itemAttachments = extractItemAttachments(service, childAttachment, properties, appendedBody, emailIdentifier);
		            }
		        }
		    }
		} catch (Exception e) {
		    throw new Exception("Exception while extracting Item Attachments: " + e.getMessage());
		}
		return itemAttachments;
		
	}

	private Item[] appendItemBody(Item item, Item[] eBody) {
		Item[] itemList = new Item[eBody.length +1];
		for(Item i : eBody) {
			itemList[itemList.length] = i;
		}
		itemList[itemList.length] = item;
		return itemList;
	}

	
	public String getPath() {
		return path;
	}

	public void setPath(String path) {
		this.path = path;
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

	
}
