package com.sun.jna.platform.win32.office.outlook;

import java.io.File;

import com.sun.jna.platform.win32.COM.IDispatch;

public class Attachments extends BaseOutlookObject {

	Attachments(IDispatch iDisp) {
		super(iDisp);
	}
	
	public Attachment add(File src) {
		
		return new Attachment((IDispatch) invoke("Add", newVariant(src.getAbsolutePath())).getValue());
	}
	
	public int count() {
		
		return getIntProperty("Count");
	}
	
	public Attachment getItem(int index) {
		
		return new Attachment((IDispatch) invoke("Item", newVariant(index)).getValue());
	}
	
	public void remove(int index) {
		
		invokeNoReply("Remove", newVariant(index));
	}
	
}
