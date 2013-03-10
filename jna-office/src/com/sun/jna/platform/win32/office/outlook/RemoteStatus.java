package com.sun.jna.platform.win32.office.outlook;

public class RemoteStatus extends AbstractEnum {
	
	public final static RemoteStatus	olRemoteStatusNone	= new RemoteStatus(0, "olRemoteStatusNone");	//No remote status has been set.
	public final static RemoteStatus	olUnMarked			= new RemoteStatus(1, "olUnMarked");			//Item is not marked.
	public final static RemoteStatus	olMarkedForDownload	= new RemoteStatus(2, "olMarkedForDownload");	//Item is marked for download.
	public final static RemoteStatus	olMarkedForCopy		= new RemoteStatus(3, "olMarkedForCopy");		//Item is marked to be copied.
	public final static RemoteStatus	olMarkedForDelete	= new RemoteStatus(4, "olMarkedForDelete");		//Item is marked for deletion.
	
	private RemoteStatus(int status, String name) {
		super((short) status, name);
	}

	public static RemoteStatus parse(short status) {
		
		switch(status) {
		
		case 1:
			return olUnMarked;
			
		case 2:
			return olMarkedForDownload;
			
		case 3:
			return olMarkedForCopy;
			
		case 4:
			return olMarkedForDelete;
		
		case 0:
		default:
			return olRemoteStatusNone;
		}
	}

}
