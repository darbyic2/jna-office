package com.sun.jna.platform.win32.office.outlook;

import java.io.File;

import com.sun.jna.platform.win32.COM.IDispatch;

public class Attachment extends BaseOutlookObject {

	public Attachment(IDispatch iDisp) {
		super(iDisp);
	}
	
	public AttachmentBlockLevel getBlockLevel() {
		
		return AttachmentBlockLevel.parse(getShortProperty("BlockLevel"));
	}
	
	public void delete() {
		
		invokeNoReply("Delete");
	}
	
	public String getDisplayName() {
		
		return getStringProperty("DisplayName");
	}
	
	public void setDisplayName(String name) {
		
		setProperty("DisplayName", name);
	}
	
	public String getFileName() {
		
		return getStringProperty("FileName");
	}
	
	public File getTemporyFilePath() {
		
		return new File((invoke("GetTemporyFilePath").getValue().toString()));
	}
	
	public int getIndex() {
		
		return getIntProperty("Index");
	}
	
	public String getPathName() {
		
		return getStringProperty("PathName");
	}
	
	public int getPosition() {
		
		return getIntProperty("Position");
	}
	
	public void saveAsFile(File file) {
		
		invoke("SaveAsFile", newVariant(file.getAbsolutePath()));
	}
	
	public int getSize() {
		
		return getIntProperty("Size");
	}
	
	public AttachmentType getType() {
		
		return AttachmentType.parse(getShortProperty("Type"));
	}

}
