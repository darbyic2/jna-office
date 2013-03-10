package com.sun.jna.platform.win32.office.outlook;

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;

public class AppointmentItem extends BaseItemLevel2 {

	public AppointmentItem(IDispatch iDisp) {
		super(iDisp);
	}

	public boolean isAllDayEvent() {
		
		return getBooleanProperty("AllDayEvent");
	}
	
	public void setAllDayEvent(boolean flag) {
		
		setProperty("AllDayEvent", flag);
	}
	
	public BusyStatus getBusyStatus() {
		
		return BusyStatus.parse(getShortProperty("BusyStatus"));
	}
	
	public void setBusyStatus(BusyStatus status) {
		
		setProperty("BusyStatus", status.value());
	}
	
	public void clearRecurrencePattern() {
		
		invokeNoReply("ClearRecurrencePattern");
	}
	
	public void close(InspectorCloseOption option) {
		
		invokeNoReply("Close", newVariant(option.value()));
	}
	
	public Conversation getConversation() {
		
		return new Conversation((IDispatch) invoke("GetConversation").getValue());
	}
	
	public AppointmentItem copy() {
		
		return new AppointmentItem((IDispatch) invoke("Copy").getValue());
	}
	
	public AppointmentItem copyTo(Folder folder, AppointmentCopyOption option) {
		
		return new AppointmentItem((IDispatch) invoke("CopyTo", newVariant(folder.getIDispatch()), newVariant(option.value())).getValue());
	}
	
	public void display() {
		
		display(false);
	}
	
	public void display(boolean modal) {
		
		invokeNoReply("Display", newVariant(modal));
	}
	
	public int getDuration() {
		
		return getIntProperty("Duration");
	}
	
	public void setDuration(int durn) {
		
		setProperty("Duration", durn);
	}
	
	public Date getEndTime() {
		
		return getDateProperty("End");
	}
	
	public void setEndTime(Date tim) {
		
		setProperty("End", tim);
	}
	
	public Date getEndInEndTimeZone() {
		
		return getDateProperty("EndInEndTimeZone");
	}
	
	public void setEndInEndTimeZone(Date tim) {
		
		setProperty("EndInEndTimeZone", tim);
	}
	
	public TimeZone getEndTimeZone() {
		
		return new TimeZone(getAutomationProperty("EndTimeZone"));
	}
	
	public void setEndTimeZone(TimeZone tim) {
		
		setProperty("EndTimeZone", tim.getIDispatch());
	}
	
	public Date getEndUTC() {
		
		return getDateProperty("EndUTC");
	}
	
	public void setEndUTC(Date tim) {
		
		setProperty("EndUTC", tim);
	}
	
	public boolean isForceUpdateToAllAttendees() {
		
		return getBooleanProperty("ForceUpdateToAllAttendees");
	}
	
	public void setForceUpdateToAllAttendees(boolean flag) {
		
		setProperty("ForceUpdateToAllAttendees", flag);
	}
	
	public MailItem forwardAsVcal() {
		
		return new MailItem((IDispatch) invoke("ForwardAsVcal").getValue());
	}
	
	public String getGlobalAppointmentID() {
		
		return getStringProperty("GetGlobalAppointmentID");
	}
	
	public int getInternetCodePage() {
		
		return getIntProperty("InternetCodePage");
	}
	
	public void setInternetCodePage(int codePage) {
		
		setProperty("InternetCodePage", codePage);
	}
	
	public String getLocation() {
		
		return getStringProperty("Location");
	}
	
	public void setLocation(String location) {
		
		setProperty("Location", location);
	}
	
	public MeetingStatus getMeetingStatus() {
		
		return MeetingStatus.parse(getShortProperty("MeetingStatus"));
	}
	
	public void setMeetingStatus(MeetingStatus fmt) {
		
		setProperty("MeetingStatus", fmt.value());
	}
	
	public String getMeetingWorkspaceURL() {
		
		return getStringProperty("MeetingWorkspaceURL");
	}
	
	public String getOptionalAttendees() {
		
		return getStringProperty("OptionalAttendees");
	}
	
	public void setOptionalAttendees(String attendees) {
		
		setProperty("OptionalAttendees", attendees);
	}
	
	public AddressEntry getOrganizer() {
		
		return new AddressEntry((IDispatch) invoke("GetOrganizer").getValue());
	}
	
	public String getOrganizerName() {
		
		return getStringProperty("Organizer");
	}
	
	public Recipients getRecipients() {
		
		return new Recipients(getAutomationProperty("Recipients"));
	}
	
	public RecurrencePattern getRecurrencePattern() {
		
		return new RecurrencePattern((IDispatch) invoke("GetRecurrencePattern").getValue());
	}
	
	public RecurrenceState getRecurrenceState() {
		
		return RecurrenceState.parse(getShortProperty("RecurrenceState"));
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
	
	public int getReminderMinutesBeforeStart() {
		
		return getIntProperty("ReminderMinutesBeforeStart");
	}
	
	public void setReminderMinutesBeforeStart(int durn) {
		
		setProperty("ReminderMinutesBeforeStart", durn);
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
	
	public Date getReplyTime() {
		
		return getDateProperty("ReplyTime");
	}
	
	public void setReplyTime(Date tim) {
		
		setProperty("ReplyTime", tim);
	}
	
	public String getRequiredAttendees() {
		
		return getStringProperty("RequiredAttendees");
	}
	
	public void setRequiredAttendees(String attendees) {
		
		setProperty("RequiredAttendees", attendees);
	}
	
	public String getResources() {
		
		return getStringProperty("Resources");
	}
	
	public void setgetResources(String resources) {
		
		setProperty("Resources", resources);
	}
	
	public MeetingItem respond(MeetingResponse rsp) {
		
		return respond(rsp, false, true);
	}
	
	public MeetingItem respond(MeetingResponse rsp, boolean inhibitUI) {
		
		return respond(rsp, inhibitUI, true);
	}
	
	public MeetingItem respond(MeetingResponse rsp, boolean inhibitUI, boolean promptForInput) {
		
		return new MeetingItem((IDispatch) invoke("Respond", newVariant(inhibitUI), newVariant(promptForInput)).getValue());
	}
	
	public boolean isResponseRequested() {
		
		return getBooleanProperty("ResponseRequested");
	}
	
	public void setResponseRequested(boolean flag) {
		
		setProperty("ResponseRequested", flag);
	}
	
	public ResponseStatus getResponseStatus() {
		
		return ResponseStatus.parse(getShortProperty("ResponseStatus"));
	}
	
//	/**
//	 * @return not working properly. RTFBody property returns a Byte array.
//	 */
//	public String getRTFBody() {
//		/* TODO Needs fixing. Should be getting a Byte array back. 
//		 */
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
//		/* TODO Needs fixing. Should be getting a Byte array back. 
//		 */
//		
//		setProperty("RTFBody", text);
//	}
	
	public void send() {
		
		invokeNoReply("Send");
	}
	
	public Account getSendUsingAccount() {
		
		return new Account(getAutomationProperty("SendUsingAccount"));
	}
	
	public void setSendUsingAccount(Account acct) {
		
		setProperty("SendUsingAccount", acct.getIDispatch());
	}
	
	public Date getStartime() {
		
		return getDateProperty("Start");
	}
	
	public void setStartTime(Date tim) {
		
		setProperty("Start", tim);
	}
	
	public Date getStartInEndTimeZone() {
		
		return getDateProperty("StartInEndTimeZone");
	}
	
	public void setStartInEndTimeZone(Date tim) {
		
		setProperty("StartInEndTimeZone", tim);
	}
	
	public TimeZone getStartTimeZone() {
		
		return new TimeZone(getAutomationProperty("StartTimeZone"));
	}
	
	public void StartTimeZone(TimeZone tim) {
		
		setProperty("StartTimeZone", tim.getIDispatch());
	}
	
	public Date getStartUTC() {
		
		return getDateProperty("StartUTC");
	}
	
	public void setStartUTC(Date tim) {
		
		setProperty("StartUTC", tim);
	}
	
}
