package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class BaseTaskItem extends BaseItemLevel2 {

	public BaseTaskItem(IDispatch iDisp) {
		super(iDisp);
	}
	
	public Conversation getConversation() {
		
		return new Conversation((IDispatch) invoke("GetConversation").getValue());
	}
	
}
