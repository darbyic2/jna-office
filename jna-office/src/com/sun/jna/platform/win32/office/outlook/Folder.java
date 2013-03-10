package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class Folder extends BaseOutlookObject {

	Folder(IDispatch iDispatch) {
		super(iDispatch);
	}
	
	public String getAddressBookName() {
		
		return getStringProperty("AddressBookName");
	}
	
	public void setAddressBookName(String name) {
		
		setProperty("AddressBookName", name);
	}
	
	public void addToPFFavourites() {
		
		invokeNoReply("AddToPFFavourites");
	}
	
	public Folder copyTo(Folder destination) {
		
		return new Folder((IDispatch) invoke("CopyTo", newVariant(destination.getIDispatch())).getValue());
	}
	
	public ItemType getDefaultItemType() {
		
		return ItemType.parse(getShortProperty("DefaultItemType"));
	}
	
	public void delete() {
		
		invokeNoReply("Delete");
	}
	
	public String getDescription() {
		
		return getStringProperty("Description");
	}
	
	public void setDescription(String desc) {
		
		setProperty("Description", desc);
	}
	
	public void display() {
		
		invokeNoReply("Display");
	}
	
	public String getEntryID() {
		
		return getStringProperty("EntryID");
	}
	
	public String getFolderPath() {
		
		return getStringProperty("FolderPath");
	}
	
	public Folders getFolders() {
		
		return new Folders(getAutomationProperty("Folders"));
	}
	
	public boolean isInAppFolderSyncObject() {
		
		return getBooleanProperty("InAppFolderSyncObject");
	}
	
	public void setInAppFolderSyncObject(boolean flag) {
		
		setProperty("InAppFolderSyncObject", flag);
	}
	
	public boolean isSharePointFolder() {
		
		return getBooleanProperty("IsSharePointFolder");
	}
	
	public Items getItems() {
		
		return new Items(getAutomationProperty("Items"));
	}
	
}
