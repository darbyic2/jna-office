package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class Namespace extends BaseOutlookObject {
	
	public final static String MAPI = "Mapi";
	

	Namespace(IDispatch iDispatch) {
		super(iDispatch);
	}
	
	public Folder getDefaultFolder(FolderType defaultType) {
		
		return new Folder((IDispatch) invoke("GetDefaultFolder", newVariant(defaultType.value())).getValue());
	}

	public boolean isOffline() {
		
		return getBooleanProperty("Offline");
	}
	
	public void logoff() {
		
		invokeNoReply("Logoff");
	}
	
	public Folder openSharedFolder(String urlOrFilePath) {
		
		return new Folder((IDispatch) invoke("OpenSharedFolder", newVariant(urlOrFilePath)).getValue());
	}
	
	public Folder pickFolder() {
		
		return new Folder((IDispatch) invoke("PickFolder").getValue());
	}
	
	/**
	 * Only good for removing .pst files from the Outlook user interface.
	 * You can not remove a store from the main mailbox on the server, or remove
	 * a folder from the users hard drive using the Outlook model.
	 * 
	 * @param folder
	 */
	public void removeStore(Folder folder) {
		
		invokeNoReply("RemoveStore", folder.toVariant());
	}
	
	/**
	 * Equivalent to Send/Receive All. This is a synchronous operation.
	 * Initiates immediate delivery of all undelivered messages submitted in the
	 * current session, and immediate receipt of mail for all accounts in the
	 * current profile.
	 * 
	 * @param showProgressDialog
	 */
	public void sendAndReceive(boolean showProgressDialog) {
		
		invokeNoReply("SendAndReceive", newVariant(showProgressDialog));
	}
	
	/**
	 * @return type of session. Only 'MAPI' is supported.
	 */
	public String getType() {
		
		return getStringProperty("Type");
	}
	
}
