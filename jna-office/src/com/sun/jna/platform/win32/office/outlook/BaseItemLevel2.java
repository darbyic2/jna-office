package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public abstract class BaseItemLevel2 extends BaseItemLevel1 {

	protected BaseItemLevel2(IDispatch iDisp) {
		super(iDisp);
	}
	
	public Actions getActions() {
		
		return new Actions(getAutomationProperty("Actions"));
	}
	
	public Attachments getAttachments() {
		
		return new Attachments(getAutomationProperty("Attachments"));
	}
	
	public boolean isAutoResolvedWinner() {
		
		return getBooleanProperty("AutoResolvedWinner");
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
	
	public boolean isConflict() {
		
		return getBooleanProperty("IsConflict");
	}
	
	public Conflicts getConflicts() {
		
		return new Conflicts(getAutomationProperty("Conflicts"));
	}
	
	public String getConversationID() {
		
		return getStringProperty("ConversationID");
	}
	
	public String getConversationIndex() {
		
		return getStringProperty("ConversationIndex");
	}
	
	public String getConversationTopic() {
		
		return getStringProperty("ConversationTopic");
	}
	
	public ItemDownloadState getDownloadState() {
		
		return ItemDownloadState.parse(getShortProperty("DownloadState"));
	}
	
	public FormDescription getFormDescription() {
		
		return new FormDescription(getAutomationProperty("FormDescription"));
	}
	
	public Importance getImportance() {
		
		return Importance.parse(getShortProperty("Importance"));
	}
	
	public void setImportance(Importance level) {
		
		setProperty("Importance", level.value());
	}
	
	public Links getLinks() {
		
		return new Links(getAutomationProperty("Links"));
	}
	
	public RemoteStatus getMarkForDownload() {
		
		return RemoteStatus.parse(getShortProperty("MarkForDownload"));
	}
	
	public void setMarkForDownload(RemoteStatus status) {
		
		setProperty("MarkForDownload", status.value());
	}
	
	public String getMileage() {
		
		return getStringProperty("Mileage");
	}
	
	public void setMileage(String freeFormText) {
		
		setProperty("Mileage", freeFormText);
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
	
	public Sensitivity getSensitivity() {
		
		return Sensitivity.parse(getShortProperty("Sensitivity"));
	}
	
	public void setSensitivity(Sensitivity level) {
		
		setProperty("Sensitivity", level.value());
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
	
	public void showCategoriesDialog() {
		
		invokeNoReply("ShowCategoriesDialog");
	}
	
}
