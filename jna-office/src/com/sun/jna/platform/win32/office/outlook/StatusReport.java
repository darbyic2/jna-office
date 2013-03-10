package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class StatusReport extends BaseOutlookObject {
	
	public StatusReport(IDispatch iDisp) {
		super(iDisp);
	}

	public void send() {
		
		invokeNoReply("Send");
	}
}
