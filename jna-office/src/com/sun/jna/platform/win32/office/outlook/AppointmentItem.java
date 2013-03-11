/* Copyright (c) 2013 Ian Darby, All Rights Reserved
 * 
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 * 
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.  
 */

package com.sun.jna.platform.win32.office.outlook;

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;

/**
 * Represents a meeting, a one-time appointment, or a recurring appointment or
 * meeting in the Calendar folder
 * <p>
 * Use the {@link Outlook#createAppointmentItem()} method to create an
 * AppointmentItem object that represents a new appointment.
 * </p>
 * <p>
 * Use getItems(index), where index is the index number of an appointment or a
 * value used to match the default property of an appointment, to return a
 * single AppointmentItem object from a Calendar folder.
 * </p>
 * <p>
 * You can also return an AppointmentItem object from a MeetingItem object by
 * using the {@link MeetingItem#getAssociatedAppointment()} method.
 * </p>
 * 
 * @author Ian Darby
 * 
 */
public class AppointmentItem extends BaseItemLevel2 {

	/**
	 * Constructor scope is restricted to package as it should not be used
	 * directly by user applications. It is only intended to be used from within
	 * factory methods and properties of the Outlook object model itself. It may
	 * also be called from unit tests which may supply a mock version of the
	 * IDispatch object.
	 * 
	 * @param iDisp
	 *            the IDispatch object which is the underlying Actions object
	 *            within the Outlook object model. All methods and properties of
	 *            this wrapper class ultimately delegate to IDispatch.
	 */
	AppointmentItem(IDispatch iDisp) {
		super(iDisp);
	}

	/**
	 * Returns true if the appointment is an all-day event (as opposed to a
	 * specified time). Read/write.
	 * 
	 * @return true if the appointment is an all-day event (as opposed to a
	 *         specified time).
	 */
	public boolean isAllDayEvent() {
		
		return getBooleanProperty("AllDayEvent");
	}
	
	/**
	 * Set the all-day event property to true/false (as opposed to a specified
	 * time).
	 * 
	 * @param flag
	 *            true == all-day appointment; false == specified time.
	 */
	public void setAllDayEvent(boolean flag) {
		
		setProperty("AllDayEvent", flag);
	}
	
	/**
	 * Returns a BusyStatus constant indicating the busy status of the user for
	 * the appointment. Read/write.
	 * 
	 * @return a BusyStatus constant indicating the busy status of the user for
	 *         the appointment.
	 */
	public BusyStatus getBusyStatus() {
		
		return BusyStatus.parse(getShortProperty("BusyStatus"));
	}
	
	/**
	 * Sets an OlBusyStatus constant indicating the busy status of the user for
	 * the appointment. Read/write.
	 * 
	 * @param status
	 *            OlBusyStatus constant indicating the busy status of the user
	 *            for the appointment.
	 */
	public void setBusyStatus(BusyStatus status) {
		
		setProperty("BusyStatus", status.value());
	}
	
	/**
	 * Removes the recurrence settings and restores the single-occurrence state
	 * for an appointment or task.
	 */
	public void clearRecurrencePattern() {
		
		invokeNoReply("ClearRecurrencePattern");
	}
	
	/**
	 * Creates another instance of an AppointmentItem.
	 * 
	 * @return another instance of this AppointmentItem.
	 */
	public AppointmentItem copy() {
		
		return new AppointmentItem((IDispatch) invoke("Copy").getValue());
	}
	
	/**
	 * Copies the AppointmentItem to the folder that is specified by the
	 * DestinationFolder parameter and returns an object that represents the
	 * item created in the destination folder by the copy operation.
	 * 
	 * @param destinationFolder
	 *            Specifies the folder to which the AppointmentItem object is
	 *            copied.
	 * 
	 * @param option
	 *            Specifies the user experience of the copy operation.
	 * 
	 * @return an AppointmentItem that represents the object created in the
	 *         destination folder as a result of the copy operation.
	 */
	public AppointmentItem copyTo(Folder destinationFolder, AppointmentCopyOptions option) {
		
		return new AppointmentItem((IDispatch) invoke("CopyTo", newVariant(destinationFolder.getIDispatch()), newVariant(option.value())).getValue());
	}
	
	/**
	 * Returns an int indicating the duration (in minutes) of the
	 * AppointmentItem. Read/write.
	 * 
	 * @return an int indicating the duration (in minutes) of the
	 *         AppointmentItem.
	 */
	public int getDuration() {
		
		return getIntProperty("Duration");
	}
	
	/**
	 * Sets an int indicating the duration (in minutes) of the AppointmentItem.
	 * Read/write.
	 * 
	 * @param durn
	 *            the duration (in minutes) of the AppointmentItem.
	 */
	public void setDuration(int durn) {
		
		setProperty("Duration", durn);
	}
	
	/**
	 * Returns a Date indicating the end date and time of an AppointmentItem.
	 * Read/write.
	 * <p>
	 * Corresponds to the &quot;End&quot; property of the actual Outlook
	 * AppointmentItem.
	 * </p>
	 * 
	 * @return a Date indicating the end date and time of an AppointmentItem.
	 */
	public Date getEndTime() {
		
		return getDateProperty("End");
	}
	
	/**
	 * Sets a Date indicating the end date and time of an AppointmentItem.
	 * Read/write.
	 * <p>
	 * Corresponds to the &quot;End&quot; property of the actual Outlook
	 * AppointmentItem.
	 * </p>
	 * 
	 * @param tim
	 *            the end date and time of an AppointmentItem.
	 */
	public void setEndTime(Date tim) {
		
		setProperty("End", tim);
	}
	
	/**
	 * Returns a Date value that represents the end date and time of the
	 * appointment expressed in the AppointmentItem.EndTimeZone. Read/write.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a Date value that represents the end date and time of the
	 *         appointment expressed in the AppointmentItem.EndTimeZone.
	 */
	public Date getEndInEndTimeZone() {
		
		return getDateProperty("EndInEndTimeZone");
	}
	
	/**
	 * Sets a Date value that represents the end date and time of the
	 * appointment expressed in the AppointmentItem.EndTimeZone. Read/write.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param tim
	 *            Date value that represents the end date and time of the
	 *            appointment expressed in the AppointmentItem.EndTimeZone.
	 */
	public void setEndInEndTimeZone(Date tim) {
		
		setProperty("EndInEndTimeZone", tim);
	}
	
	/**
	 * Returns a TimeZone value that corresponds to the end time of the
	 * appointment. Read/write.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a TimeZone value that corresponds to the end time of the
	 *         appointment.
	 */
	public TimeZone getEndTimeZone() {
		
		return new TimeZone(getAutomationProperty("EndTimeZone"));
	}
	
	/**
	 * Sets a TimeZone value that corresponds to the end time of the
	 * appointment. Read/write.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param timZon
	 *            TimeZone value that corresponds to the end time of the
	 *            appointment.
	 */
	public void setEndTimeZone(TimeZone timZon) {
		
		setProperty("EndTimeZone", timZon.getIDispatch());
	}
	
	/**
	 * Returns a Date value that represents the end date and time of the
	 * appointment expressed in the Coordinated Univeral Time (UTC) standard.
	 * Read/write.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a Date value that represents the end date and time of the
	 *         appointment expressed in the Coordinated Univeral Time (UTC)
	 *         standard.
	 */
	public Date getEndUTC() {
		
		return getDateProperty("EndUTC");
	}
	
	/**
	 * Sets a Date value that represents the end date and time of the
	 * appointment expressed in the Coordinated Univeral Time (UTC) standard.
	 * Read/write.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param tim
	 *            a Date value that represents the end date and time of the
	 *            appointment expressed in the Coordinated Univeral Time (UTC)
	 *            standard.
	 */
	public void setEndUTC(Date tim) {
		
		setProperty("EndUTC", tim);
	}
	
	/**
	 * Returns a boolean value that indicates whether updates to the
	 * AppointmentItem object should be sent to all attendees. Read/write.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a boolean value that indicates whether updates to the
	 *         AppointmentItem object should be sent to all attendees.
	 */
	public boolean isForceUpdateToAllAttendees() {
		
		return getBooleanProperty("ForceUpdateToAllAttendees");
	}
	
	/**
	 * Sets a boolean value that indicates whether updates to the
	 * AppointmentItem object should be sent to all attendees. Read/write.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param flag
	 *            a boolean value that indicates whether updates to the
	 *            AppointmentItem object should be sent to all attendees.
	 */
	public void setForceUpdateToAllAttendees(boolean flag) {
		
		setProperty("ForceUpdateToAllAttendees", flag);
	}
	
	/**
	 * Forwards the AppointmentItem as a vCal; virtual calendar item. The
	 * ForwardAsVcal method returns a MailItem with the vCal file attached.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return A MailItem object that represents the new mail item to which the
	 *         calendar information is attached.
	 */
	public MailItem forwardAsVcal() {
		
		return new MailItem((IDispatch) invoke("ForwardAsVcal").getValue());
	}
	
	/**
	 * Returns a String value that represents a unique global identifier for the
	 * AppointmentItem object. Read-only.
	 * <p>
	 * There are situations where the entry ID of AppointmentItem objects may
	 * change, such as when an item is moved to a different folder or to a
	 * different store. Entry IDs can also change when a user performs certain
	 * functions in Outlook, such as exporting and then re-importing data.
	 * </p>
	 * <p>
	 * Therefore, each Outlook appointment item is assigned a Global Object ID,
	 * a unique global identifier which does not change during those situations.
	 * The Global Object ID is a MAPI property that Outlook uses to correlate
	 * meeting updates and responses with a particular meeting on the calendar.
	 * The Global Object ID is the same across all copies of the item.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a String value that represents a unique global identifier for the
	 *         AppointmentItem object.
	 */
	public String getGlobalAppointmentID() {
		
		return getStringProperty("GetGlobalAppointmentID");
	}
	
	/**
	 * Returns an int that determines the Internet code page used by the item.
	 * Read/write.
	 * 
	 * @return an int that determines the Internet code page used by the item.
	 */
	public int getInternetCodePage() {
		
		return getIntProperty("InternetCodePage");
	}
	
	/**
	 * Sets an int that determines the Internet code page used by the item.
	 * Read/write.
	 * 
	 * @param codePage
	 *            an int that determines the Internet code page used by the
	 *            item.
	 */
	public void setInternetCodePage(int codePage) {
		
		setProperty("InternetCodePage", codePage);
	}
	
	/**
	 * Returns a String representing the specific office location (for example,
	 * Building 1 Room 1 or Suite 123) for the appointment. Read/write.
	 * 
	 * @return a String representing the specific office location.
	 */
	public String getLocation() {
		
		return getStringProperty("Location");
	}
	
	/**
	 * Sets a String representing the specific office location (for example,
	 * Building 1 Room 1 or Suite 123) for the appointment. Read/write.
	 * 
	 * @param location
	 *            a String representing the specific office location.
	 */
	public void setLocation(String location) {
		
		setProperty("Location", location);
	}
	
	/**
	 * Returns a MeetingStatus constant specifying the meeting status of the
	 * appointment. Read/write.
	 * <p>
	 * Use this property to make a MeetingItem object available for the
	 * appointment.
	 * </p>
	 * 
	 * @return a MeetingStatus constant specifying the meeting status of the
	 *         appointment.
	 */
	public MeetingStatus getMeetingStatus() {
		
		return MeetingStatus.parse(getShortProperty("MeetingStatus"));
	}
	
	/**
	 * Sets a MeetingStatus constant specifying the meeting status of the
	 * appointment. Read/write.
	 * <p>
	 * Use this property to make a MeetingItem object available for the
	 * appointment.
	 * </p>
	 * 
	 * @param fmt
	 *            constant specifying the meeting status of the appointment.
	 */
	public void setMeetingStatus(MeetingStatus fmt) {
		
		setProperty("MeetingStatus", fmt.value());
	}
	
	/**
	 * Returns the URL for the Meeting Workspace that the appointment item is
	 * linked to. Read-only.
	 * <p>
	 * A Meeting Workspace is a shared Web site for planning the meeting and
	 * tracking the results. Typically a SharePoint site.
	 * </p>
	 * 
	 * @return the URL for the Meeting Workspace that the appointment item is
	 *         linked to.
	 */
	public String getMeetingWorkspaceURL() {
		
		return getStringProperty("MeetingWorkspaceURL");
	}
	
	/**
	 * Returns a String representing the display string of optional attendees
	 * names for the appointment. Read/write.
	 * <p>
	 * This property corresponds to the MAPI property PidTagDisplayCc.
	 * </p>
	 * 
	 * @return a String representing the display string of optional attendees
	 *         names for the appointment.
	 */
	public String getOptionalAttendees() {
		
		return getStringProperty("OptionalAttendees");
	}
	
	/**
	 * Sets a String representing the display string of optional attendees names
	 * for the appointment. Read/write.
	 * <p>
	 * This property corresponds to the MAPI property PidTagDisplayCc.
	 * </p>
	 * 
	 * @param attendees
	 *            a String representing the display string of optional attendees
	 *            names for the appointment.
	 */
	public void setOptionalAttendees(String attendees) {
		
		setProperty("OptionalAttendees", attendees);
	}
	
	/**
	 * Returns a String representing the name of the organizer of the
	 * appointment. Read-only.
	 * 
	 * @return a String representing the name of the organizer of the
	 *         appointment.
	 */
	public AddressEntry getOrganizer() {
		
		return new AddressEntry((IDispatch) invoke("GetOrganizer").getValue());
	}
	
	/**
	 * Returns a Recipients collection that represents all the recipients for
	 * the Outlook item (item: An item is the basic element that holds
	 * information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.). Read-only.
	 * 
	 * @return a Recipients collection that represents all the recipients for
	 *         the Outlook item (item: An item is the basic element that holds
	 *         information in Outlook (similar to a file in other programs).
	 *         Items include e-mail messages, appointments, contacts, tasks,
	 *         journal entries, notes, posted items, and documents.).
	 */
	public Recipients getRecipients() {
		
		return new Recipients(getAutomationProperty("Recipients"));
	}
	
	/**
	 * Returns a RecurrencePattern object that represents the recurrence
	 * attributes of an appointment.
	 * 
	 * @return a RecurrencePattern object that represents the recurrence
	 *         attributes of an appointment.
	 */
	public RecurrencePattern getRecurrencePattern() {
		
		return new RecurrencePattern((IDispatch) invoke("GetRecurrencePattern").getValue());
	}
	
	/**
	 * Returns a RecurrenceState constant indicating the recurrence property of
	 * the specified object. Read-only.
	 * 
	 * @return a RecurrenceState constant indicating the recurrence property of
	 *         the specified object.
	 */
	public RecurrenceState getRecurrenceState() {
		
		return RecurrenceState.parse(getShortProperty("RecurrenceState"));
	}
	
	/**
	 * Returns a boolean value that is true if the appointment is a recurring
	 * appointment. Read-only.
	 * <p>
	 * When the GetRecurrencePattern method is used with an AppointmentItem
	 * object, this property is set to true.
	 * </p>
	 * 
	 * @return a boolean value that is true if the appointment is a recurring
	 *         appointment.
	 */
	public boolean isRecurring() {
		
		return getBooleanProperty("IsRecurring");
	}
	
	/**
	 * Returns a boolean value that is true if a reminder has been set for this
	 * item. Read/write.
	 * 
	 * @return a boolean value that is true if a reminder has been set for this
	 *         item.
	 */
	public boolean isReminderSet() {
		
		return getBooleanProperty("ReminderSet");
	}
	
	/**
	 * Sets a boolean value that is true if a reminder has been set for this
	 * item. Read/write.
	 * 
	 * @param flag
	 *            boolean value that is true if a reminder has been set for this
	 *            item.
	 */
	public void setReminder(boolean flag) {
		
		setProperty("ReminderSet", flag);
	}
	
	/**
	 * Returns an int indicating the number of minutes the reminder should occur
	 * prior to the start of the appointment. Read/write.
	 * 
	 * @return an int indicating the number of minutes the reminder should occur
	 *         prior to the start of the appointment.
	 */
	public int getReminderMinutesBeforeStart() {
		
		return getIntProperty("ReminderMinutesBeforeStart");
	}
	
	/**
	 * Sets an int indicating the number of minutes the reminder should occur
	 * prior to the start of the appointment. Read/write.
	 * 
	 * @param durn
	 *            an int indicating the number of minutes the reminder should
	 *            occur prior to the start of the appointment.
	 */
	public void setReminderMinutesBeforeStart(int durn) {
		
		setProperty("ReminderMinutesBeforeStart", durn);
	}
	
	/**
	 * Returns a boolean value that is true if the reminder overrides the
	 * default reminder behaviour for the item. Read/write.
	 * 
	 * @return a boolean value that is true if the reminder overrides the
	 *         default reminder behaviour for the item.
	 */
	public boolean getReminderOverrideDefault() {
		
		return getBooleanProperty("ReminderOverrideDefault");
	}
	
	/**
	 * Sets a boolean value that is true if the reminder overrides the default
	 * reminder behaviour for the item. Read/write.
	 * 
	 * @param flag
	 *            a boolean value that is true if the reminder overrides the
	 *            default reminder behaviour for the item.
	 */
	public void setReminderOverrideDefault(boolean flag) {
		
		setProperty("ReminderOverrideDefault", flag);
	}
	
	/**
	 * Returns a boolean value that is true if the reminder should play a sound
	 * when it occurs for this item. Read/write.
	 * 
	 * @return a boolean value that is true if the reminder should play a sound
	 *         when it occurs for this item.
	 */
	public boolean isReminderPlaySound() {
		
		return getBooleanProperty("ReminderPlaySound");
	}
	
	/**
	 * Sets a boolean value that is true if the reminder should play a sound
	 * when it occurs for this item. Read/write.
	 * 
	 * @param flag
	 *            a boolean value that is true if the reminder should play a
	 *            sound when it occurs for this item.
	 */
	public void setReminderPlaySound(boolean flag) {
		
		setProperty("ReminderPlaySound", flag);
	}
	
	/**
	 * Returns a String indicating the path and file name of the sound file to
	 * play when the reminder occurs for the Outlook item. Read/write.
	 * <p>
	 * This property is only valid if the ReminderOverrideDefault and
	 * ReminderPlaySound properties are set to true.
	 * </p>
	 * 
	 * @return a String indicating the path and file name of the sound file to
	 *         play when the reminder occurs for the Outlook item.
	 */
	public String getReminderSoundFile() {
		
		return getStringProperty("ReminderSoundFile");
	}
	
	/**
	 * Returns or sets a String indicating the path and file name of the sound
	 * file to play when the reminder occurs for the Outlook item. Read/write.
	 * <p>
	 * This property is only valid if the ReminderOverrideDefault and
	 * ReminderPlaySound properties are set to true.
	 * </p>
	 * 
	 * @param filePath
	 *            a String indicating the path and file name of the sound file
	 *            to play when the reminder occurs for the Outlook item.
	 */
	public void setReminderSoundFile(String filePath) {
		
		setProperty("ReminderSoundFile", filePath);
	}
	
	/**
	 * Returns a Date indicating the reply time for the appointment. Read/write.
	 * 
	 * @return a Date indicating the reply time for the appointment.
	 */
	public Date getReplyTime() {
		
		return getDateProperty("ReplyTime");
	}
	
	/**
	 * Sets a Date indicating the reply time for the appointment. Read/write.
	 * 
	 * @param tim
	 *            a Date indicating the reply time for the appointment.
	 */
	public void setReplyTime(Date tim) {
		
		setProperty("ReplyTime", tim);
	}
	
	/**
	 * Returns a semicolon-delimited String of required attendee names for the
	 * meeting appointment. Read/write.
	 * <p>
	 * This property only contains the display names for the required attendees.
	 * The attendee list should be set by using the Recipients collection.
	 * </p>
	 * 
	 * @return a semicolon-delimited String of required attendee names for the
	 *         meeting appointment. Read/write.
	 */
	public String getRequiredAttendees() {
		
		return getStringProperty("RequiredAttendees");
	}
	
	/**
	 * Sets a semicolon-delimited String of required attendee names for the
	 * meeting appointment. Read/write.
	 * <p>
	 * This property only contains the display names for the required attendees.
	 * The attendee list should be set by using the Recipients collection.
	 * </p>
	 * 
	 * @param attendees
	 *            a semicolon-delimited String of required attendee names for
	 *            the meeting appointment.
	 */
	public void setRequiredAttendees(String attendees) {
		
		setProperty("RequiredAttendees", attendees);
	}
	
	/**
	 * Returns a semicolon-delimited String of resource names for the meeting.
	 * Read/write.
	 * <p>
	 * This property contains the display names only. The Recipients collection
	 * should be used to modify the resource recipients. Resources are added as
	 * BCC recipients to the collection.
	 * </p>
	 * 
	 * @return a semicolon-delimited String of resource names for the meeting.
	 */
	public String getResources() {
		
		return getStringProperty("Resources");
	}
	
	/**
	 * Sets a semicolon-delimited String of resource names for the meeting.
	 * Read/write.
	 * <p>
	 * This property contains the display names only. The Recipients collection
	 * should be used to modify the resource recipients. Resources are added as
	 * BCC recipients to the collection.
	 * </p>
	 * 
	 * @param resources
	 *            a semicolon-delimited String of resource names for the
	 *            meeting.
	 */
	public void setResources(String resources) {
		
		setProperty("Resources", resources);
	}
	
	/**
	 * Responds to a meeting request.
	 * <p>
	 * When you call the Respond method with the olMeetingAccepted or
	 * olMeetingTentative parameter, Outlook will create a new appointment item
	 * that duplicates the original appointment item. The new item will have a
	 * different Entry ID. Outlook will then remove the original item. You
	 * should no longer use the Entry ID of the original item, but instead call
	 * the EntryID property to obtain the Entry ID for the new item for any
	 * subsequent needs. This is to ensure that this appointment item will be
	 * properly synchronised on your calendar if more than one client computer
	 * accesses your calendar but may be offline using the cache mode
	 * occasionally.
	 * </p>
	 * 
	 * @param rsp
	 *            The response to the request.
	 * 
	 * @return A MeetingItem object that represents the response to the meeting
	 *         request.
	 */
	public MeetingItem respond(MeetingResponse rsp) {
		
		return respond(rsp, false, true);
	}
	
	/**
	 * Responds to a meeting request.
	 * <p>
	 * When you call the Respond method with the olMeetingAccepted or
	 * olMeetingTentative parameter, Outlook will create a new appointment item
	 * that duplicates the original appointment item. The new item will have a
	 * different Entry ID. Outlook will then remove the original item. You
	 * should no longer use the Entry ID of the original item, but instead call
	 * the EntryID property to obtain the Entry ID for the new item for any
	 * subsequent needs. This is to ensure that this appointment item will be
	 * properly synchronised on your calendar if more than one client computer
	 * accesses your calendar but may be offline using the cache mode
	 * occasionally.
	 * </p>
	 * 
	 * @param rsp
	 *            The response to the request.
	 * 
	 * @param inhibitUI
	 *            True to not display a dialog box; the response is sent
	 *            automatically. False to display the dialog box for responding.
	 * 
	 * @return A MeetingItem object that represents the response to the meeting
	 *         request.
	 */
	public MeetingItem respond(MeetingResponse rsp, boolean inhibitUI) {
		
		return respond(rsp, inhibitUI, true);
	}
	
	/**
	 * Responds to a meeting request.
	 * <p>
	 * When you call the Respond method with the olMeetingAccepted or
	 * olMeetingTentative parameter, Outlook will create a new appointment item
	 * that duplicates the original appointment item. The new item will have a
	 * different Entry ID. Outlook will then remove the original item. You
	 * should no longer use the Entry ID of the original item, but instead call
	 * the EntryID property to obtain the Entry ID for the new item for any
	 * subsequent needs. This is to ensure that this appointment item will be
	 * properly synchronised on your calendar if more than one client computer
	 * accesses your calendar but may be offline using the cache mode
	 * occasionally.
	 * </p>
	 * 
	 * @param rsp
	 *            The response to the request.
	 * 
	 * @param inhibitUI
	 *            True to not display a dialog box; the response is sent
	 *            automatically. False to display the dialog box for responding.
	 * 
	 * @param promptForInput
	 *            False to not prompt the user for input; the response is
	 *            displayed in the inspector for editing. True to prompt the
	 *            user to either send or send with comments. This argument is
	 *            valid only if fNoUI is False.
	 * 
	 * @return A MeetingItem object that represents the response to the meeting
	 *         request.
	 */
	public MeetingItem respond(MeetingResponse rsp, boolean inhibitUI, boolean promptForInput) {
		
		return new MeetingItem((IDispatch) invoke("Respond", newVariant(inhibitUI), newVariant(promptForInput)).getValue());
	}
	
	/**
	 * Returns a boolean that indicates True if the sender would like a response
	 * to the meeting request for the appointment. Read/write.
	 * 
	 * @return a boolean that indicates True if the sender would like a response
	 *         to the meeting request for the appointment.
	 */
	public boolean isResponseRequested() {
		
		return getBooleanProperty("ResponseRequested");
	}
	
	/**
	 * Returns a Boolean that indicates True if the sender would like a response
	 * to the meeting request for the appointment. Read/write.
	 * 
	 * @param flag
	 *            a boolean that indicates True if the sender would like a
	 *            response to the meeting request for the appointment.
	 */
	public void setResponseRequested(boolean flag) {
		
		setProperty("ResponseRequested", flag);
	}
	
	/**
	 * Returns a ResponseStatus constant indicating the overall status of the
	 * meeting for the current user for the appointment. Read-only.
	 * 
	 * @return a ResponseStatus constant indicating the overall status of the
	 *         meeting for the current user for the appointment. Read-only.
	 */
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
	
	/**
	 * Sends the appointment.
	 */
	public void send() {
		
		invokeNoReply("Send");
	}
	
	/**
	 * Returns an Account object that represents the account under which the
	 * AppointmentItem is to be sent. Read/write.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return an Account object that represents the account under which the
	 *         AppointmentItem is to be sent.
	 */
	public Account getSendUsingAccount() {
		
		return new Account(getAutomationProperty("SendUsingAccount"));
	}
	
	/**
	 * Sets an Account object that represents the account under which the
	 * AppointmentItem is to be sent. Read/write.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param acct
	 *            an Account object that represents the account under which the
	 *            AppointmentItem is to be sent. Read/write.
	 */
	public void setSendUsingAccount(Account acct) {
		
		setProperty("SendUsingAccount", acct.getIDispatch());
	}
	
	/**
	 * Returns a Date indicating the starting date and time for the Outlook
	 * item. Read/write.
	 * <p>
	 * Maps to the &quot;Start&quot; property of the underlying AppointmentItem
	 * object.
	 * </p>
	 * 
	 * @return a Date indicating the starting date and time for the Outlook
	 *         item.
	 */
	public Date getStartime() {
		
		return getDateProperty("Start");
	}
	
	/**
	 * Sets a Date indicating the starting date and time for the Outlook item.
	 * Read/write.
	 * <p>
	 * Maps to the &quot;Start&quot; property of the underlying AppointmentItem
	 * object.
	 * </p>
	 * 
	 * @param tim
	 *            a Date indicating the starting date and time for the Outlook
	 *            item.
	 */
	public void setStartTime(Date tim) {
		
		setProperty("Start", tim);
	}
	
	/**
	 * Returns a Date value that represents the start date and time of the
	 * appointment expressed in the AppointmentItem.StartTimeZone. Read/write.
	 * <p>
	 * This is the value displayed as Start time in the appointment inspector
	 * user interface.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a Date value that represents the start date and time of the
	 *         appointment expressed in the AppointmentItem.StartTimeZone.
	 */
	public Date getStartInStartTimeZone() {
		
		return getDateProperty("StartInStartTimeZone");
	}
	
	/**
	 * Sets a Date value that represents the start date and time of the
	 * appointment expressed in the AppointmentItem.StartTimeZone. Read/write.
	 * <p>
	 * This is the value displayed as Start time in the appointment inspector
	 * user interface.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param tim
	 *            a Date value that represents the start date and time of the
	 *            appointment expressed in the AppointmentItem.StartTimeZone.
	 */
	public void setStartInStartTimeZone(Date tim) {
		
		setProperty("StartInStartTimeZone", tim);
	}
	
	/**
	 * Returns a TimeZone value that corresponds to the time zone for the start
	 * time of the appointment. Read/write.
	 * <p>
	 * The time zone information is used to map the appointment to the correct
	 * UTC time when the appointment is saved, and into the correct local time
	 * when the item is displayed in the calendar.
	 * </p>
	 * <p>
	 * Changing StartTimeZone affects the value of AppointmentItem.Start which
	 * is always represented in the local time zone,
	 * Application.TimeZones.CurrentTimeZone.
	 * </p>
	 * <p>
	 * Depending on the circumstances, changing the StartTimeZone may or may not
	 * cause Outlook to recalculate and update the
	 * AppointmentItem.StartInStartTimeZone.
	 * </p>
	 * <p>
	 * As an example, in the appointment inspector, if you are the organiser of
	 * an appointment with a start time at 1 P.M. PST and end time at 3 P.M.
	 * PST, changing the appointment to have an StartTimeZone of EST will result
	 * in an appointment lasting from 1 P.M. EST to 3 P.M. PST, with the
	 * StartInStartTimeZone remaining as 1 P.M. However, if you are not the
	 * organiser, then changing the StartTimeZone from PST to EST will cause
	 * Outlook to recalculate and update the StartInStartTimeZone, and the
	 * appointment will last from 4 P.M. EST to 3 P.M. PST.
	 * </p>
	 * <p>
	 * Another example is changing the StartTimeZone resulting in an appointment
	 * end time that occurs before a previously set appointment start time, in
	 * which case Outlook will recalculate and update the StartInStartTimeZone.
	 * For example, an appointment with a start time at 1 P.M. EST and end time
	 * at 3 P.M. EST has its StartTimeZone changed to PST. If Outlook did not
	 * recalculate the StartInStartTimeZone, the appointment would have a start
	 * time at 1 P.M. PST, which is equivalent to 4 P.M. EST, and which would
	 * occur before the end time of 3 P.M. EST. In practice, however, changing
	 * the StartTimeZone would result in Outlook recalculating and updating the
	 * StartInStartTimeZone to 10 A.M. (in the StartTimeZone PST).
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a TimeZone value that corresponds to the time zone for the start
	 *         time of the appointment.
	 */
	public TimeZone getStartTimeZone() {
		
		return new TimeZone(getAutomationProperty("StartTimeZone"));
	}
	
	/**
	 * Sets a TimeZone value that corresponds to the time zone for the start
	 * time of the appointment. Read/write.
	 * <p>
	 * The time zone information is used to map the appointment to the correct
	 * UTC time when the appointment is saved, and into the correct local time
	 * when the item is displayed in the calendar.
	 * </p>
	 * <p>
	 * Changing StartTimeZone affects the value of AppointmentItem.Start which
	 * is always represented in the local time zone,
	 * Application.TimeZones.CurrentTimeZone.
	 * </p>
	 * <p>
	 * Depending on the circumstances, changing the StartTimeZone may or may not
	 * cause Outlook to recalculate and update the
	 * AppointmentItem.StartInStartTimeZone.
	 * </p>
	 * <p>
	 * As an example, in the appointment inspector, if you are the organiser of
	 * an appointment with a start time at 1 P.M. PST and end time at 3 P.M.
	 * PST, changing the appointment to have an StartTimeZone of EST will result
	 * in an appointment lasting from 1 P.M. EST to 3 P.M. PST, with the
	 * StartInStartTimeZone remaining as 1 P.M. However, if you are not the
	 * organiser, then changing the StartTimeZone from PST to EST will cause
	 * Outlook to recalculate and update the StartInStartTimeZone, and the
	 * appointment will last from 4 P.M. EST to 3 P.M. PST.
	 * </p>
	 * <p>
	 * Another example is changing the StartTimeZone resulting in an appointment
	 * end time that occurs before a previously set appointment start time, in
	 * which case Outlook will recalculate and update the StartInStartTimeZone.
	 * For example, an appointment with a start time at 1 P.M. EST and end time
	 * at 3 P.M. EST has its StartTimeZone changed to PST. If Outlook did not
	 * recalculate the StartInStartTimeZone, the appointment would have a start
	 * time at 1 P.M. PST, which is equivalent to 4 P.M. EST, and which would
	 * occur before the end time of 3 P.M. EST. In practice, however, changing
	 * the StartTimeZone would result in Outlook recalculating and updating the
	 * StartInStartTimeZone to 10 A.M. (in the StartTimeZone PST).
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param tim
	 *            a TimeZone value that corresponds to the time zone for the
	 *            start time of the appointment.
	 */
	public void setStartTimeZone(TimeZone tim) {
		
		setProperty("StartTimeZone", tim.getIDispatch());
	}
	
	/**
	 * Returns a Date value that represents the start date and time of the
	 * appointment expressed in the Coordinated Univeral Time (UTC) standard.
	 * Read/write.
	 * <p>
	 * Changing the value for the AppointmentItem.Start property or the
	 * AppointmentItem.StartTimeZone property will cause Outlook to recalculate
	 * the value of StartUTC.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a Date value that represents the start date and time of the
	 *         appointment expressed in the Coordinated Univeral Time (UTC)
	 *         standard.
	 */
	public Date getStartUTC() {
		
		return getDateProperty("StartUTC");
	}
	
	/**
	 * Sets a Date value that represents the start date and time of the
	 * appointment expressed in the Coordinated Univeral Time (UTC) standard.
	 * Read/write.
	 * <p>
	 * Changing the value for the AppointmentItem.Start property or the
	 * AppointmentItem.StartTimeZone property will cause Outlook to recalculate
	 * the value of StartUTC.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param tim
	 *            a Date value that represents the start date and time of the
	 *            appointment expressed in the Coordinated Univeral Time (UTC)
	 *            standard.
	 */
	public void setStartUTC(Date tim) {
		
		setProperty("StartUTC", tim);
	}
	
}
