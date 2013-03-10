package com.sun.jna.platform.win32.office.outlook;

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;

public class BaseItemLevel4 extends BaseItemLevel3 {

	protected BaseItemLevel4(IDispatch iDisp) {
		super(iDisp);
	}
	
	public void addBusinessCard(ContactItem contact) {
		
		invoke("AddBusinessCard", newVariant(contact.getIDispatch()));
	}

	public boolean isAlternateRecipientAllowed() {
		
		return getBooleanProperty("AlternateRecipientAllowed");
	}
	
	public void setAlternateRecipientAllowed(boolean canForwardItem) {
		
		setProperty("VotingResponse", canForwardItem);
	}
	
	public boolean isAutoForwarded() {
		
		return getBooleanProperty("AutoForwarded");
	}
	
	public void setAutoForwarded(boolean flag) {
		
		setProperty("AutoForwarded", flag);
	}
	
	public String getBcc() {
		
		return getStringProperty("BCC");
	}
	
	public void setBcc(String bcc) {
		
		setProperty("BCC", bcc);
	}
	
	public MailBodyFormat getBodyFormat() {
		
		return MailBodyFormat.parse(getIntProperty("BodyFormat"));
	}
	
	public void setBodyFormat(MailBodyFormat fmt) {
		
		setProperty("BodyFormat", fmt.value());
	}
	
	public String getCc() {
		
		return getStringProperty("CC");
	}
	
	public void setCc(String cc) {
		
		setProperty("CC", cc);
	}
	
	public void clearConversationIndex() {
		
		invokeNoReply("ClearConversationIndex");
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
	
	public String getFlagRequest() {
		
		return getStringProperty("FlagRequest");
	}
	
	public void setFlagRequest(String flagText) {
		
		setProperty("FlagRequest", flagText);
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
	
	public boolean isOriginatorDeliveryReportRequested() {
		
		return getBooleanProperty("OriginatorDeliveryReportRequested");
	}
	
	public void setOriginatorDeliveryReportRequested(boolean flag) {
		
		setProperty("OriginatorDeliveryReportRequested", flag);
	}
	
	public Permission getPermission() {
		
		return Permission.parse(getShortProperty("Permission"));
	}
	
	public void setPermission(Permission permission) {
		
		setProperty("Permission", permission.value());
	}
	
	public PermissionService getPermissionService() {
		
		return PermissionService.parse(getShortProperty("PermissionService"));
	}
	
	public void setPermissionService(PermissionService permission) {
		
		setProperty("PermissionService", permission.value());
	}
	
	public String getPermissionTemplateGuid() {
		
		return getStringProperty("PermissionTemplateGuid");
	}
	
	public void setPermissionTemplateGuid(String guid) {
		
		setProperty("PermissionTemplateGuid", guid);
	}
	
	public boolean isReadReceiptRequested() {
		
		return getBooleanProperty("ReadReceiptRequested");
	}
	
	public void setReadReceiptRequested(boolean flag) {
		
		setProperty("ReadReceiptRequested", flag);
	}
	
	public String getReceivedByEntryID() {
		
		return getStringProperty("ReceivedByEntryID");
	}
	
	public String getReceivedByName() {
		
		return getStringProperty("ReceivedByName");
	}
	
	public String getReceivedOnBehalfOfEntryID() {
		
		return getStringProperty("ReceivedOnBehalfOfEntryID");
	}
	
	public String getReceivedOnBehalfOfName() {
		
		return getStringProperty("ReceivedOnBehalfOfName");
	}
	
	public Date getReceivedTime() {
		
		return getDateProperty("ReceivedTime");
	}
	
	public boolean isRecipientReassignmentProhibited() {
		
		return getBooleanProperty("RecipientReassignmentProhibited");
	}
	
	public void setRecipientReassignmentProhibited(boolean flag) {
		
		setProperty("RecipientReassignmentProhibited", flag);
	}
	
	public Recipients getRecipients() {
		
		return new Recipients(getAutomationProperty("Recipients"));
	}
	
	public RemoteStatus getRemoteStatus() {
		
		return RemoteStatus.parse(getShortProperty("RemoteStatus"));
	}
	
	public void setRemoteStatus(RemoteStatus status) {
		
		setProperty("RemoteStatus", status.value());
	}
	
	public String getReplyRecipientNames() {
		
		return getStringProperty("ReplyRecipientNames");
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
	
	public String getSentOnBehalfOfName() {
		
		return getStringProperty("SentOnBehalfOfName");
	}
	
	public void setSentOnBehalfOfName(String name) {
		
		setProperty("SentOnBehalfOfName", name);
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
	
}
