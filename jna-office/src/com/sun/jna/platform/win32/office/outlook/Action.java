package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.OleAuto;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant.VARIANT;
import com.sun.jna.platform.win32.WinDef.LONG;

public class Action extends BaseOutlookObject {

	public Action(IDispatch iDisp) {
		super(iDisp);
	}
	
	public ActionCopyLike getCopyLike() {
		
		return ActionCopyLike.parse(getShortProperty("CopyLike"));
	}
	
	public void setCopyLike(ActionCopyLike actionType) {
		
		setProperty("CopyLike", actionType.value());
	}

	public void delete() {
		
		invokeNoReply("Delete");
	}
	
	public boolean isEnabled() {
		
		return getBooleanProperty("Enabled");
	}
	
	public void setEnabled(boolean flag) {
		
		setProperty("Enabled", flag);
	}
	
	private BaseItemLevel1 wrappedObject(IDispatch iDisp) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_PROPERTYGET, result, iDisp, "Class");
		
		int classId = ((LONG) result.getValue()).intValue();
		
		switch (classId) {
		
		case ClassEnum.olMail:
			return new MailItem(iDisp);
			
		case ClassEnum.olAppointment:
			return new AppointmentItem(iDisp);
			
		case ClassEnum.olJournal:
			return new JournalItem(iDisp);
			
		case ClassEnum.olPost:
			return new PostItem(iDisp);
			
		case ClassEnum.olTask:
			return new TaskItem(iDisp);
		
		case ClassEnum.olContact:
			return new ContactItem(iDisp);
			
			//The following are all currently unsupported item types
		case ClassEnum.olDistributionList:
		case ClassEnum.olDocument:
		case ClassEnum.olMeetingCancellation:
		case ClassEnum.olMeetingForwardNotification:
		case ClassEnum.olMeetingRequest:
		case ClassEnum.olMeetingResponseNegative:
		case ClassEnum.olMeetingResponsePositive:
		case ClassEnum.olMeetingResponseTentative:
		case ClassEnum.olMobile:
		case ClassEnum.olRemote:
		case ClassEnum.olReport:
		case ClassEnum.olSharing:
		case ClassEnum.olTaskRequest:
		case ClassEnum.olTaskRequestAccept:
		case ClassEnum.olTaskRequestDecline:
		case ClassEnum.olTaskRequestUpdate:
		default:
			throw new RuntimeException("Action.execute() not yet implemented for class ID: " + classId);
		}
	}
	
	public BaseItemLevel1 execute() {
		
		return wrappedObject((IDispatch) invoke("Execute").getValue());
	}
	
	public String getMessageClass() {
		
		return getStringProperty("MessageClass");
	}
	
	public void setMessageClass(String messageClass) {
		
		setProperty("MessageClass", messageClass);
	}

	public String getName() {
		
		return getStringProperty("Name");
	}
	
	public void setName(String name) {
		
		setProperty("Name", name);
	}

	public String getPrefix() {
		
		return getStringProperty("Prefix");
	}
	
	public void setPrefix(String prefix) {
		
		setProperty("Prefix", prefix);
	}

	public ActionReplyStyle getReplyStyle() {
		
		return ActionReplyStyle.parse(getShortProperty("ReplyStyle"));
	}
	
	public void setReplyStyle(ActionReplyStyle style) {
		
		setProperty("ReplyStyle", style.value());
	}

	public ActionResponseStyle getResponseStyle() {
		
		return ActionResponseStyle.parse(getShortProperty("ResponseStyle"));
	}
	
	public void setResponseStyle(ActionResponseStyle style) {
		
		setProperty("ResponseStyle", style.value());
	}

	public ActionShowOn getShowOn() {
		
		return ActionShowOn.parse(getShortProperty("ShowOn"));
	}
	
	public void setShowOn(ActionShowOn style) {
		
		setProperty("ShowOn", style.value());
	}

}
