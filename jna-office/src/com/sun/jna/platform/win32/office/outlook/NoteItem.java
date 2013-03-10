package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class NoteItem extends BaseItemLevel1 {

	public NoteItem(IDispatch iDisp) {
		super(iDisp);
	}

	public boolean isAutoResolvedWinner() {
		
		return getBooleanProperty("AutoResolvedWinner");
	}
	
	public boolean isConflict() {
		
		return getBooleanProperty("IsConflict");
	}
	
	public Conflicts getConflicts() {
		
		return new Conflicts(getAutomationProperty("Conflicts"));
	}
	
	public ItemDownloadState getDownloadState() {
		
		return ItemDownloadState.parse(getShortProperty("DownloadState"));
	}
	
	public int getHeight() {
		
		return getIntProperty("Height");
	}
	
	public void setHeight(int height) {
		
		setProperty("Height", height);
	}
	
	public int getLeft() {
		
		return getIntProperty("Left");
	}
	
	public void setLeft(int left) {
		
		setProperty("Left", left);
	}
	
	public Links getLinks() {
		
		return new Links(getAutomationProperty("Links"));
	}
	
	public RemoteStatus getMarkForDownload() {
		
		return RemoteStatus.parse(getShortProperty("MarkForDownload"));
	}
	
	public void setMarkForDownload(RemoteStatus status) {
		
		setProperty("MarkForDownload", status.value());
	}
	
	public int getTop() {
		
		return getIntProperty("Top");
	}
	
	public void setTop(int top) {
		
		setProperty("Top", top);
	}
	
	public int getWidth() {
		
		return getIntProperty("Width");
	}
	
	public void setWidth(int width) {
		
		setProperty("Width", width);
	}
	
}
