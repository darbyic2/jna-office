package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.WinDef.BOOL;

public class Inspector extends BaseOutlookObject {

	Inspector(IDispatch iDisp) {
		super(iDisp);
	}
	
	public void activate() {
		
		invokeNoReply("Activate");
	}
	
	public AttachmentSelection getAttachmentSelection() {
		
		return new AttachmentSelection((IDispatch) invoke("AttachmentSelection").getValue());
	}
	
	public String getCaption() {
		
		return getStringProperty("Caption");
	}
	
	public void close(InspectorCloseOption saveMode) {
		
		invokeNoReply("Close", newVariant(saveMode.value()));
	}
	
	public BaseOutlookObject getCurrentItem() {
		
		return new BaseOutlookObject((IDispatch) invoke("CurrentItem").getValue());
	}
	
	public void display() {
		
		display(false);
	}
	
	public void display(boolean modal) {
		
		invokeNoReply("Display", newVariant(modal));
	}
	
	public EditorType getEditorType() {
		
		return EditorType.parse(getShortProperty("EditorType"));
	}
	
	public int getHeight() {
		
		return getIntProperty("Height");
	}
	
	public void setHeight(int pixels) {
		
		setProperty("Height", pixels);
	}
	
	public void hideFormPage(String pageName) {
		
		invokeNoReply("HideFormPage", newVariant(pageName));
	}
	
	public boolean isWordMail() {
		
		return (((BOOL) invoke("IsWordMail").getValue()).intValue() != 0);
	}
	
	public int getLeft() {
		
		return getIntProperty("Left");
	}
	
	public void setLeft(int pixels) {
		
		setProperty("Left", pixels);
	}
	
	public Pages getModifiedFormPages() {
		
		return new Pages(getAutomationProperty("ModifiedFormPages"));
	}
	
	public int getTop() {
		
		return getIntProperty("Top");
	}
	
	public void setTop(int pixels) {
		
		setProperty("Top", pixels);
	}
	
	public int getWidth() {
		
		return getIntProperty("Width");
	}
	
	public void setWidth(int pixels) {
		
		setProperty("Width", pixels);
	}
	
	public WindowState getWindowState() {
		
		return WindowState.parse(getShortProperty("WindowState"));
	}
	
	public void setWindowState(WindowState state) {
		
		setProperty("WindowState", state.value());
	}
	
}
