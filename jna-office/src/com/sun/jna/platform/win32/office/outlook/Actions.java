package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class Actions extends BaseOutlookObject {

	Actions(IDispatch iDisp) {
		super(iDisp);
	}
	
	public Action add() {
		
		return new Action((IDispatch) invoke("Add").getValue());
	}
	
	public int count() {
		
		return getIntProperty("Count");
	}
	
	public Action getItem(String actionName) {
		
		return new Action((IDispatch) invoke("Item", newVariant(actionName)).getValue());
	}
	
	public void remove(String actionName) {
		
		invokeNoReply("Display", newVariant(actionName));
	}
	
}
