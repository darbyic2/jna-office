package com.sun.jna.platform.win32.office.outlook;

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.OaIdl.VARIANT_BOOL;

public class TaskItem extends BaseTaskItem {

	public TaskItem(IDispatch auto) {
		super(auto);
	}

	public int getActualWork() {
		
		return getIntProperty("ActualWork");
	}
	
	public void setActualWork(int minutes) {
		
		setProperty("ActualWork", minutes);
	}
	
	public TaskItem assign() {
		
		return new TaskItem((IDispatch) invoke("Assign").getValue());
	}
	
	public void cancelResponseState() {
		
		invokeNoReply("CancelResponseState");
	}
	
	public String getCardData() {
		
		return getStringProperty("CardData");
	}
	
	public void setCardData(String text) {
		
		setProperty("CardData", text);
	}
	
	public void clearRecurrencePattern() {
		
		invokeNoReply("ClearRecurrencePattern");
	}
	
	public boolean isComplete() {
		
		return getBooleanProperty("Complete");
	}
	
	public void setComplete(boolean flag) {
		
		setProperty("Complete", flag);
	}
	
	public String getContactNames() {
		
		return getStringProperty("ContactNames");
	}
	
	public void setContactNames(String names) {
		
		setProperty("ContactNames", names);
	}
	
	public Date getDateCompleted() {
		
		return getDateProperty("DateCompleted");
	}
	
	public void setDateCompleted(Date dat) {
		
		setProperty("DateCompleted", dat);
	}
	
	public TaskDelegationState getDelegationState() {
		
		return TaskDelegationState.parse(getShortProperty("DelegationState"));
	}
	
	public String getDelegator() {
		
		return getStringProperty("Delegator");
	}
	
	public Date getDueDate() {
		
		return getDateProperty("DueDate");
	}
	
	public void setDueDate(Date dat) {
		
		setProperty("DueDate", dat);
	}
	
	public int getInternetCodePage() {
		
		return getIntProperty("InternetCodePage");
	}
	
	public void setInternetCodePage(int codePage) {
		
		setProperty("InternetCodePage", codePage);
	}
	
	public void markComplete() {
		
		invokeNoReply("MarkComplete");
	}
	
	public int getOrdinal() {
		
		return getIntProperty("Ordinal");
	}
	
	public void setOrdinal(int val) {
		
		setProperty("Ordinal", val);
	}
	
	public String getOwner() {
		
		return getStringProperty("Owner");
	}
	
	public void setOwner(String val) {
		
		setProperty("Owner", val);
	}
	
	public TaskOwnership getOwnership() {
		
		return TaskOwnership.parse(getShortProperty("Ownership"));
	}
	
	public int getPercentComplete() {
		
		return getIntProperty("PercentComplete");
	}
	
	public void setPercentComplete(int pc) {
		
		setProperty("PercentComplete", pc);
	}
	
	public Recipients getRecipients() {
		
		return new Recipients(getAutomationProperty("Recipients"));
	}
	
	public RecurrencePattern getRecurrencePattern() {
		
		return new RecurrencePattern((IDispatch) invoke("GetRecurrencePattern").getValue());
	}
	
	public boolean isRecurring() {
		
		return getBooleanProperty("IsRecurring");
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
	
	public TaskItem respond(TaskResponse rsp, boolean noUI, boolean promptForComments) {
		
		return new TaskItem((IDispatch) invoke("Respond", newVariant(noUI), newVariant(promptForComments)).getValue());
	}
	
	public TaskResponse getResponseState() {
		
		return TaskResponse.parse(getShortProperty("ResponseState"));
	}
	
	public String getRole() {
		
		return getStringProperty("Role");
	}
	
	public void setRole(String role) {
		
		setProperty("Role", role);
	}
	
	public String getSchedulePlusPriority() {
		
		return getStringProperty("SchedulePlusPriority");
	}
	
	public void setSchedulePlusPriority(String val) {
		
		setProperty("SchedulePlusPriority", val);
	}
	
	public void send() {
		
		invokeNoReply("Send");
	}
	
	public Account getSendUsingAccount() {
		
		return new Account(getAutomationProperty("SendUsingAccount"));
	}
	
	public void setSendUsingAccount(Account acct) {
		
		setProperty("SendUsingAccount", acct.getIDispatch());
	}
	
	public boolean skipRecurrence() {
		
		return (((VARIANT_BOOL) invoke("SkipRecurrence").getValue()).intValue() != 0);
	}
	
	public Date getStartDate() {
		
		return getDateProperty("StartDate");
	}
	
	public void setStartDate(Date dat) {
		
		setProperty("StartDate", dat);
	}
	
	public TaskStatus getStatus() {
		
		return TaskStatus.parse(getShortProperty("Status"));
	}
	
	public void setStatus(TaskStatus status) {
		
		setProperty("Status", status.value());
	}
	
	public String getStatusOnCompletionRecipients() {
		
		return getStringProperty("StatusOnCompletionRecipients");
	}
	
	public void setStatusOnCompletionRecipients(String val) {
		
		setProperty("StatusOnCompletionRecipients", val);
	}
	
	public StatusReport getStatusReport() {
		
		return new StatusReport((IDispatch) invoke("SkipRecurrence").getValue());
	}
	
	public String getStatusUpdateRecipients() {
		
		return getStringProperty("StatusUpdateRecipients");
	}
	
	public void setStatusUpdateRecipients(String val) {
		
		setProperty("StatusUpdateRecipients", val);
	}
	
	public boolean isTeamTask() {
		
		return getBooleanProperty("TeamTask");
	}
	
	public void setTeamTask(boolean flag) {
		
		setProperty("TeamTask", flag);
	}
	
	public Date getToDoTaskOrdinal() {
		
		return getDateProperty("ToDoTaskOrdinal");
	}
	
	public void setToDoTaskOrdinal(Date dat) {
		
		setProperty("ToDoTaskOrdinal", dat);
	}
	
	public int getTotalWork() {
		
		return getIntProperty("TotalWork");
	}
	
	public void setTotalWork(int minutes) {
		
		setProperty("TotalWork", minutes);
	}
	
}
