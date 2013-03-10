package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class SharingItem extends BaseItemLevel4 {

	public SharingItem(IDispatch iDisp) {
		super(iDisp);
	}

	public void allow() {
		
		invokeNoReply("Allow");
	}
	
	public boolean isAllowWriteAccess() {
		
		return getBooleanProperty("AllowWriteAccess");
	}
	
	public void setAllowWriteAccess(boolean flag) {
		
		setProperty("AllowWriteAccess", flag);
	}
	
	public SharingItem deny() {
		
		return new SharingItem((IDispatch) invoke("Deny").getValue());
	}
	
	public Folder openSharedFolder() {
		
		return new Folder((IDispatch) invoke("OpenSharedFolder").getValue());
	}
	
	public String getRemoteID() {
		
		return getStringProperty("RemoteID");
	}
	
	public String getRemoteName() {
		
		return getStringProperty("RemoteName");
	}
	
	public String getRemotePath() {
		
		return getStringProperty("RemotePath");
	}
	
	public DefaultFolder getRequestedFolder() {
		
		return DefaultFolder.parse(getShortProperty("RequestedFolder"));
	}
	
	public SharingProvider getSharingProvider() {
		
		return SharingProvider.parse(getShortProperty("SharingProvider"));
	}
	
	public String getSharingProviderGuid() {
		
		return getStringProperty("SharingProviderGuid");
	}
	
	public SharingMsgType getType() {
		
		return SharingMsgType.parse(getShortProperty("Type"));
	}
	
	public void setSharingProviderGuid(SharingMsgType typ) {
		
		setProperty("Type", typ.value());
	}
	
}
