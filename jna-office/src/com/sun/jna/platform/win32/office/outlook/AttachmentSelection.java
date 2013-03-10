package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class AttachmentSelection extends BaseOutlookObject {

	AttachmentSelection(IDispatch iDisp) {
		super(iDisp);
	}
	
	public int count() {

		return getIntProperty("Count");
	}
	
	public Attachment getItem(int index) {
		
		return new Attachment((IDispatch) invoke("Item", newVariant(index)).getValue());
	}
	
	public Attachment getItem(String name) {
		
		return new Attachment((IDispatch) invoke("Item", newVariant(name)).getValue());
	}
	
	public SelectionLocation getLocation() {
		
		return SelectionLocation.parse(getShortProperty("Location"));
	}
}
