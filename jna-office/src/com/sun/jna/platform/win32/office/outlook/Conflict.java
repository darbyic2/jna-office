package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class Conflict extends BaseOutlookObject {

	Conflict(IDispatch iDisp) {
		super(iDisp);
	}
	
	public BaseOutlookObject getConflictingItem() {
		
		return new BaseOutlookObject(getAutomationProperty("Item"));
	}
	
	public String getName() {
		
		return getStringProperty("Name");
	}
	
	public int getConflictObjectClassEnum() {
		
		return getIntProperty("Type");
	}
}
