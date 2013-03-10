package com.sun.jna.platform.win32.office.outlook;

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;

public class MobileItem extends BaseItemLevel1 {

	public MobileItem(IDispatch iDisp) {
		super(iDisp);
	}
	
	public Actions getActions() {
		
		return new Actions(getAutomationProperty("Actions"));
	}
	
	public Attachments getAttachments() {
		
		return new Attachments(getAutomationProperty("Attachments"));
	}
	
	public String getBillingInformation() {
		
		return getStringProperty("BillingInformation");
	}
	
	public void setBillingInformation(String billingInfo) {
		
		setProperty("BillingInformation", billingInfo);
	}
	
	public String getCompanies() {
		
		return getStringProperty("Companies");
	}
	
	public void setCompanies(String companies) {
		
		setProperty("Companies", companies);
	}
	
	public String getConversationIndex() {
		
		return getStringProperty("ConversationIndex");
	}
	
	public String getConversationTopic() {
		
		return getStringProperty("ConversationTopic");
	}
	
	public int getCount() {
		
		return getIntProperty("Count");
	}
	
	public FormDescription getFormDescription() {
		
		return new FormDescription(getAutomationProperty("FormDescription"));
	}
	
	public MobileItem forward() {
		
		return new MobileItem((IDispatch) invoke("Forward").getValue());
	}
	
	public String getHtmlBody() {
		
		return getStringProperty("HTMLBody");
	}
	
	public void setHtmlBody(String htmlBody) {
		
		setProperty("HTMLBody", htmlBody);
	}
	
	public Importance getImportance() {
		
		return Importance.parse(getShortProperty("Importance"));
	}
	
	public void setImportance(Importance level) {
		
		setProperty("Importance", level.value());
	}
	
	public String getMileage() {
		
		return getStringProperty("Mileage");
	}
	
	public void setMileage(String freeFormText) {
		
		setProperty("Mileage", freeFormText);
	}
	
	public MobileFormat getMobileFormat() {
		
		return MobileFormat.parse(getShortProperty("MobileFormat"));
	}
	
	public boolean isNoAging() {
		
		return getBooleanProperty("NoAging");
	}
	
	public void setNoAging(boolean flag) {
		
		setProperty("NoAging", flag);
	}
	
	public int getOutlookInternalVersion() {
		
		return getIntProperty("OutlookInternalVersion");
	}
	
	public String getOutlookVersion() {
		
		return getStringProperty("OutlookVersion");
	}
	
	public String getReceivedByEntryID() {
		
		return getStringProperty("ReceivedByEntryID");
	}
	
	public String getReceivedByName() {
		
		return getStringProperty("ReceivedByName");
	}
	
	public Date getReceivedTime() {
		
		return getDateProperty("ReceivedTime");
	}
	
	public Recipients getRecipients() {
		
		return new Recipients(getAutomationProperty("Recipients"));
	}
	
	public MobileItem reply() {
		
		return new MobileItem((IDispatch) invoke("Reply").getValue());
	}
	
	public MobileItem replyAll() {
		
		return new MobileItem((IDispatch) invoke("ReplyAll").getValue());
	}
	
	public String getReplyRecipientNames() {
		
		return getStringProperty("ReplyRecipientNames");
	}
	
	public Recipients getReplyRecipients() {
		
		return new Recipients(getAutomationProperty("ReplyRecipients"));
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
	
	public Sensitivity getSensitivity() {
		
		return Sensitivity.parse(getShortProperty("Sensitivity"));
	}
	
	public void setSensitivity(Sensitivity level) {
		
		setProperty("Sensitivity", level.value());
	}
	
	public boolean isSent() {
		
		return getBooleanProperty("Sent");
	}
	
	public Date getSentOn() {
		
		return getDateProperty("SentOn");
	}
	
	public String getSMILBody() {
		
		return getStringProperty("SMILBody");
	}
	
	public void setSMILBody(String bodyText) {
		
		setProperty("SMILBody", bodyText);
	}
	
	public boolean isSubmitted() {
		
		return getBooleanProperty("Submitted");
	}
	
	public String getTo() {
		
		return getStringProperty("To");
	}
	
	public void setTo(String to) {
		
		setProperty("To", to);
	}
	
	public boolean isUnRead() {
		
		return getBooleanProperty("UnRead");
	}
	
	public void setUnread(boolean flag) {
		
		setProperty("Unread", flag);
	}
	
	public UserProperties getUserProperties() {
		
		return new UserProperties(getAutomationProperty("UserProperties"));
	}
	
}
