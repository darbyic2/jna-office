package com.sun.jna.platform.win32.office.outlook;

import java.io.File;

import com.sun.jna.platform.win32.COM.IDispatch;

public class FormDescription extends BaseOutlookObject {

	FormDescription(IDispatch iDisp) {
		super(iDisp);
	}
	
	public String getCategory() {
		
		return getStringProperty("Category");
	}
	
	public void setCategory(String cat) {
		
		setProperty("Category", cat);
	}
	
	public String getCategorySub() {
		
		return getStringProperty("CategorySub");
	}
	
	public void setCategorySub(String subCat) {
		
		setProperty("CategorySub", subCat);
	}
	
	public String getComment() {
		
		return getStringProperty("Comment");
	}
	
	public void setComment(String comment) {
		
		setProperty("Comment", comment);
	}
	
	public String getContactName() {
		
		return getStringProperty("ContactName");
	}
	
	public void setContactName(String name) {
		
		setProperty("ContactName", name);
	}
	
	public String getDisplayName() {
		
		return getStringProperty("DisplayName");
	}
	
	public void setDisplayName(String name) {
		
		setProperty("DisplayName", name);
	}
	
	public boolean isHidden() {
		
		return getBooleanProperty("Hidden");
	}
	
	public void setHidden(boolean flag) {
		
		setProperty("Hidden", flag);
	}
	
	public File getIcon() {
		
		return new File(getStringProperty("Icon"));
	}
	
	public void setIcon(File iconFilePath) {
		
		setProperty("Icon", iconFilePath.getAbsolutePath());
	}
	
	public boolean isLocked() {
		
		return getBooleanProperty("Locked");
	}
	
	public void setLocked(boolean flag) {
		
		setProperty("Locked", flag);
	}
	
	public String getMessageClass() {
		
		return getStringProperty("MessageClass");
	}
	
	public File getMiniIcon() {
		
		return new File(getStringProperty("MiniIcon"));
	}
	
	public void setMiniIcon(File iconFilePath) {
		
		setProperty("MiniIcon", iconFilePath.getAbsolutePath());
	}
	
	public String getName() {
		
		return getStringProperty("Name");
	}
	
	public void setName(String name) {
		
		setProperty("Name", name);
	}
	
	public String getNumber() {
		
		return getStringProperty("Number");
	}
	
	public void setNumber(String num) {
		
		setProperty("Number", num);
	}
	
	public boolean isOneOff() {
		
		return getBooleanProperty("OneOff");
	}
	
	public void setOneOff(boolean flag) {
		
		setProperty("OneOff", flag);
	}
	
	public void publishForm(FormRegistry registry) {
		
		if (registry == FormRegistry.olFolderRegistry)
			throw new RuntimeException("Publishing Form to Folder Registry is only valid with a Folder destination");
		
		invokeNoReply("PublishForm", newVariant(registry.value()));
	}
	
	public void publishForm(Folder folder) {
		
		invokeNoReply("PublishForm", newVariant(FormRegistry.olFolderRegistry.value()), newVariant(folder.getIDispatch()));
	}
	
	public String getScriptText() {
		
		return getStringProperty("ScriptText");
	}
	
	public String getTemplate() {
		
		return getStringProperty("Template");
	}
	
	public void setTemplate(String name) {
		
		setProperty("Template", name);
	}
	
	public boolean useWordMail() {
		
		return getBooleanProperty("UseWordMail");
	}
	
	public void setUseWordMail(boolean flag) {
		
		setProperty("UseWordMail", flag);
	}
	
	public String getVersion() {
		
		return getStringProperty("Version");
	}
	
	public void setVersionVersion(String vn) {
		
		setProperty("Version", vn);
	}
	
}
