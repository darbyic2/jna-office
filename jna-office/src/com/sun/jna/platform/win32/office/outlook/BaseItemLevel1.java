package com.sun.jna.platform.win32.office.outlook;

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;

public class BaseItemLevel1 extends BaseOutlookObject {

	protected BaseItemLevel1(IDispatch iDispatch) {
		super(iDispatch);
	}
	
	public String getBody() {
		
		return getStringProperty("Body");
	}
	
	public void setBody(String body) {
		
		setProperty("Body", body);
	}
	
	public String getCategories() {
		
		return getStringProperty("Categories");
	}
	
	public void setCategories(String categoryList) {
		
		setProperty("Categories", categoryList);
	}
	
	public void close(InspectorCloseOption option) {
		
		invokeNoReply("Close", newVariant(option.value()));
	}
	
	public void delete() {
		
		invokeNoReply("Delete");
	}
	
	public void display() {
		
		display(false);
	}
	
	public void display(boolean modal) {
		
		invokeNoReply("Display", newVariant(modal));
	}
	
	public Date getCreationTime() {
		
		return getDateProperty("CreationTime");
	}
	
	public String getEntryID() {
		
		return getStringProperty("EntryID");
	}
	
	public Inspector getInspector() {
		
		return new Inspector(getAutomationProperty("GetInspector"));
	}
	
	public ItemProperties getItemProperties() {
		
		return new ItemProperties(getAutomationProperty("ItemProperties"));
	}
	
	public Date getLastModificationTime() {
		
		return getDateProperty("LastModificationTime");
	}
	
	public String getMessageClass() {
		
		return getStringProperty("MessageClass");
	}
	
	public void setMessageClass(String classOfMessage) {
		
		setProperty("MessageClass", classOfMessage);
	}
	
	public void move(Folder folder) {
		
		invokeNoReply("Move", newVariant(folder.getIDispatch()));
	}
	
	public void printOut() {
		
		invokeNoReply("PrintOut");
	}
	
	public PropertyAccessor getPropertyAccessor() {
		
		return new PropertyAccessor(getAutomationProperty("PropertyAccessor"));
	}
	
	public void save() {
		
		invokeNoReply("Save");
	}
	
	public void saveAs(String filePath) {
		
		invokeNoReply("SaveAs", newVariant(filePath));
	}
	
	public void saveAs(String filePath, SaveAsType typ) {
		
		invokeNoReply("SaveAs", newVariant(filePath), newVariant(typ.value()));
	}
	
	public boolean isSaved() {
		
		return getBooleanProperty("Saved");
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

}
