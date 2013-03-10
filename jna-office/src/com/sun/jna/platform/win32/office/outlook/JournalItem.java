package com.sun.jna.platform.win32.office.outlook;

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;

public class JournalItem extends BaseItemLevel2 {

	public JournalItem(IDispatch iDisp) {
		super(iDisp);
	}

	public Conversation getConversation() {
		
		return new Conversation((IDispatch) invoke("GetConversation").getValue());
	}
	
	public String getContactNames() {
		
		return getStringProperty("ContactNames");
	}
	
	public void setContactNames(String names) {
		
		setProperty("ContactNames", names);
	}
	
	public boolean isDocPosted() {
		
		return getBooleanProperty("DocPosted");
	}
	
	public void setDocPosted(boolean flag) {
		
		setProperty("DocPosted", flag);
	}
	
	public boolean isDocPrinted() {
		
		return getBooleanProperty("DocPrinted");
	}
	
	public void setDocPrinted(boolean flag) {
		
		setProperty("DocPrinted", flag);
	}
	
	public boolean isDocRouted() {
		
		return getBooleanProperty("DocRouted");
	}
	
	public void setDocRouted(boolean flag) {
		
		setProperty("DocRouted", flag);
	}
	
	public boolean isDocSaved() {
		
		return getBooleanProperty("DocSaved");
	}
	
	public void setDocSaved(boolean flag) {
		
		setProperty("DocSaved", flag);
	}
	
	public int getDuration() {
		
		return getIntProperty("Duration");
	}
	
	public void setDuration(int minutes) {
		
		setProperty("Duration", minutes);
	}
	
	public Date getEndTime() {
		
		return getDateProperty("End");
	}
	
	public void setEndTime(Date tim) {
		
		setProperty("End", tim);
	}
	
	public MailItem forward() {
		
		return new MailItem((IDispatch) invoke("Forward").getValue());
	}
	
	public Recipients getRecipients() {
		
		return new Recipients(getAutomationProperty("Recipients"));
	}
	
	public MailItem reply() {
		
		return new MailItem((IDispatch) invoke("Reply").getValue());
	}
	
	public MailItem replyAll() {
		
		return new MailItem((IDispatch) invoke("ReplyAll").getValue());
	}
	
	public Date getStartime() {
		
		return getDateProperty("Start");
	}
	
	public void setStartTime(Date tim) {
		
		setProperty("Start", tim);
	}
	
	public void startTimer() {
		
		invokeNoReply("StartTimer");
	}
	
	public void stopTimer() {
		
		invokeNoReply("StopTimer");
	}
	
	public String getType() {
		
		return getStringProperty("Type");
	}
	
	public void setType(String typ) {
		
		setProperty("Type", typ);
	}
	
}
