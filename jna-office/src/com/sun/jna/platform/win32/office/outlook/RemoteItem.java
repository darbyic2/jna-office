package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class RemoteItem extends BaseItemLevel2 {
	
	public RemoteItem(IDispatch iDisp) {
		super(iDisp);
	}

	public Conversation getConversation() {
		
		return new Conversation((IDispatch) invoke("GetConversation").getValue());
	}
	
	public boolean hasAttachment() {
		
		return getBooleanProperty("HasAttachment");
	}
	
	public String getRemoteMessageClass() {
		
		return getStringProperty("RemoteMessageClass");
	}
	
	public int getTransferSize() {
		
		return getIntProperty("TransferSize");
	}
	
	public int getTransferTime() {
		
		return getIntProperty("TransferTime");
	}
	
}
