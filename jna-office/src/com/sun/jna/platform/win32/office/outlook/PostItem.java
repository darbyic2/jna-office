package com.sun.jna.platform.win32.office.outlook;

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;

public class PostItem extends BaseItemLevel3 {

	public PostItem(IDispatch iDisp) {
		super(iDisp);
	}

	public MailBodyFormat getBodyFormat() {
		
		return MailBodyFormat.parse(getIntProperty("BodyFormat"));
	}
	
	public void setBodyFormat(MailBodyFormat fmt) {
		
		setProperty("BodyFormat", fmt.value());
	}
	
	public void clearConversationIndex() {
		
		invokeNoReply("ClearConversationIndex");
	}
	
	public Date getExpiryTime() {
		
		return getDateProperty("ExpiryTime");
	}
	
	public void setExpiryTime(Date expireAt) {
		
		setProperty("ExpiryTime", expireAt);
	}
	
	public MailItem forward() {
		
		return new MailItem((IDispatch) invoke("Forward").getValue());
	}
	
	public String getHtmlBody() {
		
		return getStringProperty("HTMLBody");
	}
	
	public void setHtmlBody(String htmlBody) {
		
		setProperty("HTMLBody", htmlBody);
	}
	
	public int getInternetCodePage() {
		
		return getIntProperty("InternetCodePage");
	}
	
	public void setInternetCodePage(int codePage) {
		
		setProperty("InternetCodePage", codePage);
	}
	
	public void post() {
		
		invokeNoReply("Post");
	}
	
	public Date getReceivedTime() {
		
		return getDateProperty("ReceivedTime");
	}
	
	public MailItem reply() {
		
		return new MailItem((IDispatch) invoke("Reply").getValue());
	}
	
	public String getSenderEmailAddress() {
		
		return getStringProperty("SenderEmailAddress");
	}
	
	public String getSenderEmailType() {
		
		return getStringProperty("SenderEmailType");
	}
	
	public String getSenderName() {
		
		return getStringProperty("SenderName");
	}
	
	public Date getSentOn() {
		
		return getDateProperty("SentOn");
	}
	
}
