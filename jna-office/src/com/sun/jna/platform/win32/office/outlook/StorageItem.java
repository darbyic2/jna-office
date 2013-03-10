package com.sun.jna.platform.win32.office.outlook;

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;

public class StorageItem extends BaseOutlookObject {

	public StorageItem(IDispatch iDisp) {
		super(iDisp);
	}
	
	public Attachments getAttachments() {
		
		return new Attachments(getAutomationProperty("Attachments"));
	}
	
	public String getBody() {
		
		return getStringProperty("Body");
	}
	
	public void setBody(String body) {
		
		setProperty("Body", body);
	}
	
	public Date getCreationTime() {
		
		return getDateProperty("CreationTime");
	}
	
	public String getCreator() {
		
		return getStringProperty("Creator");
	}
	
	public void setCreator(String progIdOfCreator) {
		
		setProperty("Creator", progIdOfCreator);
	}

	public void delete() {
		
		invokeNoReply("Delete");
	}
	
	public String getEntryID() {
		
		return getStringProperty("EntryID");
	}
	
	public Date getLastModificationTime() {
		
		return getDateProperty("LastModificationTime");
	}
	
	public PropertyAccessor getPropertyAccessor() {
		
		return new PropertyAccessor(getAutomationProperty("PropertyAccessor"));
	}
	
	public void save() {
		
		invokeNoReply("Save");
	}
	
	public int getSize() {
		
		return getIntProperty("Size");
	}
	
	public String getSubject() {
		
		return getStringProperty("Subject");
	}
	
	public void setSubject(String subject) {
		
		setProperty("Subject", subject);
	}

	public UserProperties getUserProperties() {
		
		return new UserProperties(getAutomationProperty("UserProperties"));
	}
	
}
