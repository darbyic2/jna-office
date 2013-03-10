package com.sun.jna.platform.win32.office.outlook;

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;

public class BaseItemLevel3 extends BaseItemLevel2 {

	protected BaseItemLevel3(IDispatch iDisp) {
		super(iDisp);
	}
	
	public void clearTaskFlag() {
		
		invokeNoReply("ClearTaskFlag");
	}
	
	public Conversation getConversation() {
		
		return new Conversation((IDispatch) invoke("GetConversation").getValue());
	}
	
	public boolean isMarkedAsTask() {
		
		return getBooleanProperty("IsMarkedAsTask");
	}
	
	public void markAsTask(MarkInterval interval) {
		
		invokeNoReply("MarkAsTask", newVariant(interval.value()));
	}
	
	public boolean isReminderSet() {
		
		return getBooleanProperty("ReminderSet");
	}
	
	public void setReminder(boolean flag) {
		
		setProperty("ReminderSet", flag);
	}
	
	public boolean getReminderOverrideDefault() {
		
		return getBooleanProperty("ReminderOverrideDefault");
	}
	
	public void setReminderOverrideDefault(boolean flag) {
		
		setProperty("ReminderOverrideDefault", flag);
	}
	
	public boolean isReminderPlaySound() {
		
		return getBooleanProperty("ReminderPlaySound");
	}
	
	public void setReminderPlaySound(boolean flag) {
		
		setProperty("ReminderPlaySound", flag);
	}
	
	public String getReminderSoundFile() {
		
		return getStringProperty("ReminderSoundFile");
	}
	
	public void setReminderSoundFile(String filePath) {
		
		setProperty("ReminderSoundFile", filePath);
	}
	
	public Date getReminderTime() {
		
		return getDateProperty("ReminderTime");
	}
	
	public void setReminderTime(Date start) {
		
		setProperty("ReminderTime", start);
	}
	
//	/**
//	 * @return not working properly. RTFBody property returns a Byte array.
//	 */
//	public String getRTFBody() {
//		/* TODO Needs fixing. Should be getting a Byte array back */
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
//		/* TODO Needs fixing. Should be getting a Byte array back */
//		
//		setProperty("RTFBody", text);
//	}
	
	public Date getTaskCompletedDate() {
		
		return getDateProperty("TaskCompletedDate");
	}
	
	public void setTaskCompletedDate(Date dat) {
		
		setProperty("TaskCompletedDate", dat);
	}
	
	public Date getTaskDueDate() {
		
		return getDateProperty("TaskDueDate");
	}
	
	public void setTaskDueDate(Date dat) {
		
		setProperty("TaskDueDate", dat);
	}
	
	public Date getTaskStartDate() {
		
		return getDateProperty("TaskStartDate");
	}
	
	public void setTaskStartDate(Date dat) {
		
		setProperty("TaskStartDate", dat);
	}
	
	public String getTaskSubject() {
		
		return getStringProperty("TaskSubject");
	}
	
	public void setTaskSubject(String subject) {
		
		setProperty("TaskSubject", subject);
	}
	
	public Date getToDoTaskOrdinal() {
		
		return getDateProperty("ToDoTaskOrdinal");
	}
	
	public void setToDoTaskOrdinal(Date dat) {
		
		setProperty("ToDoTaskOrdinal", dat);
	}
	
}
