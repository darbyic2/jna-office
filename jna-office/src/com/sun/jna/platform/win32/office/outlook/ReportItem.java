package com.sun.jna.platform.win32.office.outlook;

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;

public class ReportItem extends BaseItemLevel2 {
	
	public ReportItem(IDispatch iDisp) {
		super(iDisp);
	}

	public Conversation getConversation() {
		
		return new Conversation((IDispatch) invoke("GetConversation").getValue());
	}
	
	public Date getRetentionExpirationDate() {
		
		return getDateProperty("RetentionExpirationDate");
	}
	
	public String getRetentionPolicyName() {
		
		return getStringProperty("RetentionPolicyName");
	}
	
}
