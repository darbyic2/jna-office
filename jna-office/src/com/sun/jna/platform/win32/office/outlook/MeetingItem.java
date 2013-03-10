package com.sun.jna.platform.win32.office.outlook;

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;

public class MeetingItem extends BaseItemLevel2 {

	public MeetingItem(IDispatch iDisp) {
		super(iDisp);
	}

	public AppointmentItem getAssociatedAppointment() {
		
		return new AppointmentItem((IDispatch) invoke("GetAppointmentItem").getValue());
	}
	
	public boolean isAutoForwarded() {
		
		return getBooleanProperty("AutoForwarded");
	}
	
	public void setAutoForwarded(boolean flag) {
		
		setProperty("AutoForwarded", flag);
	}
	
	public Conversation getConversation() {
		
		return new Conversation((IDispatch) invoke("GetConversation").getValue());
	}
	
	public Date getDeferredDeliveryTime() {
		
		return getDateProperty("DeferredDeliveryTime");
	}
	
	public void setDeferredDeliveryTime(Date deliverAt) {
		
		setProperty("DeferredDeliveryTime", deliverAt);
	}
	
	public boolean isDeleteAfterSubmit() {
		
		return getBooleanProperty("DeleteAfterSubmit");
	}
	
	public void setDeleteAfterSubmit(boolean flag) {
		
		setProperty("DeleteAfterSubmit", flag);
	}
	
	public Date getExpiryTime() {
		
		return getDateProperty("ExpiryTime");
	}
	
	public void setExpiryTime(Date expireAt) {
		
		setProperty("ExpiryTime", expireAt);
	}
	
	public MeetingItem forward() {
		
		return new MeetingItem((IDispatch) invoke("Forward").getValue());
	}
	
	public boolean isLatestVersion() {
		
		return getBooleanProperty("IsLatestVersion");
	}
	
	public String getMeetingWorkspaceURL() {
		
		return getStringProperty("MeetingWorkspaceURL");
	}
	
	public boolean isOriginatorDeliveryReportRequested() {
		
		return getBooleanProperty("OriginatorDeliveryReportRequested");
	}
	
	public void setOriginatorDeliveryReportRequested(boolean flag) {
		
		setProperty("OriginatorDeliveryReportRequested", flag);
	}
	
	public Recipients getRecipients() {
		
		return new Recipients(getAutomationProperty("Recipients"));
	}
	
	public Date getReceivedTime() {
		
		return getDateProperty("ReceivedTime");
	}
	
	public void setReceivedTime(Date rxAt) {
		
		setProperty("ReceivedTime", rxAt);
	}
	
	public boolean isReminderSet() {
		
		return getBooleanProperty("ReminderSet");
	}
	
	public void setReminder(boolean flag) {
		
		setProperty("ReminderSet", flag);
	}
	
	public Date getReminderTime() {
		
		return getDateProperty("ReminderTime");
	}
	
	public void setReminderTime(Date start) {
		
		setProperty("ReminderTime", start);
	}
	
	public MailItem reply() {
		
		return new MailItem((IDispatch) invoke("Reply").getValue());
	}
	
	public MailItem replyAll() {
		
		return new MailItem((IDispatch) invoke("ReplyAll").getValue());
	}
	
	public Recipients getReplyRecipients() {
		
		return new Recipients(getAutomationProperty("ReplyRecipients"));
	}
	
	public Date getRetentionExpirationDate() {
		
		return getDateProperty("RetentionExpirationDate");
	}
	
	public String getRetentionPolicyName() {
		
		return getStringProperty("RetentionPolicyName");
	}
	
//	/**
//	 * @return not working properly. RTFBody property returns a Byte array.
//	 */
//	public String getRTFBody() {
//		/* TODO Needs fixing. Should be getting a Byte array back. */
//		
//		return getStringProperty("RTFBody");
//	}
//	
//	/**
//	 * Not working properly. RTFBody property returns a Byte array.
//	 * 
//	 * @param text
//	 */
//	public void setRTFBody(String text) {
//		/* TODO Needs fixing. Should be getting a Byte array back. */
//		
//		setProperty("RTFBody", text);
//	}
	
	public Folder getSaveSentMessageFolder() {
		
		return new Folder(getAutomationProperty("SaveSentMessageFolder"));
	}
	
	public void setSaveSentMessageFolder(Folder folder) {
		
		setProperty("SaveSentMessageFolder", folder.getIDispatch());
	}
	
	public void send() {
		
		invokeNoReply("Send");
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
	
	public Account getSendUsingAccount() {
		
		return new Account(getAutomationProperty("SendUsingAccount"));
	}
	
	public void setSendUsingAccount(Account acct) {
		
		setProperty("SendUsingAccount", acct.getIDispatch());
	}
	
	public boolean isSent() {
		
		return getBooleanProperty("Sent");
	}
	
	public Date getSentOn() {
		
		return getDateProperty("SentOn");
	}
	
	public boolean isSubmitted() {
		
		return getBooleanProperty("Submitted");
	}
	
}
