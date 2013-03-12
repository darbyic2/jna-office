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

import com.sun.jna.platform.win32.Variant;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant.VARIANT;

/**
 * There are an inordinate amount of
 * methods and properties that are shared between multiple Item types. However,
 * the commonality model is complex and there is a fine line between reducing
 * duplication and having a ridiculous depth to the inheritance model. The
 * compromise struck here was to have four levels of base class. Item classes
 * inherit from the most appropriate level.
 * 
 * @author Ian Darby
 * 
 * @see BaseItemLevel2
 * @see BaseItemLevel4
 * @see ContactItem
 * @see DistributionListItem
 * @see PostItem
 */
public class BaseItemLevel3 extends BaseItemLevel2 {

	/**
	 * Constructor scope is restricted to inheritance and package as it should
	 * not be used directly by user applications. It is only intended to be used
	 * from within factory methods and properties of the Outlook object model
	 * itself. It may also be called from unit tests which may supply a mock
	 * version of the IDispatch object.
	 * 
	 * @param iDisp
	 *            the IDispatch object which is the underlying Actions object
	 *            within the Outlook object model. All methods and properties of
	 *            this wrapper class ultimately delegate to IDispatch.
	 */
	protected BaseItemLevel3(IDispatch iDisp) {
		super(iDisp);
	}
	
	/**
	 * Clears the ContactItem object as a task.
	 * <p>
	 * Calling this method sets the {@link #isMarkedAsTask()} property to False.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 */
	public void clearTaskFlag() {
		
		invokeNoReply("ClearTaskFlag");
	}
	
	/**
	 * Obtains a Conversation object that represents the conversation to which
	 * this item belongs.
	 * <p>
	 * GetConversation returns Null (Nothing in Visual Basic) if no conversation
	 * exists for the item. No conversation exists for an item in the following
	 * scenarios:
	 * <ul>
	 * <li>The item has not been saved. An item can be saved programmatically,
	 * by user action, or by auto-save.</li>
	 * 
	 * <li>For an item that can be sent (for example, a mail item, appointment
	 * item, or contact item), the item has not been sent.</li>
	 * 
	 * <li>Conversations have been disabled through the Windows registry.</li>
	 * 
	 * <li>The store does not support Conversation view (for example, Outlook is
	 * running in classic online mode against a version of Microsoft Exchange
	 * earlier than Microsoft Exchange Server 2010). Use the
	 * IsConversationEnabled property of the Store object to determine whether
	 * the store supports Conversation view.</li>
	 * </ul>
	 * </p>
	 * 
	 * @return a Conversation object that represents the conversation to which
	 *         this item belongs.
	 */
	public Conversation getConversation() {

		VARIANT result = invoke("GetConversation");

		if (result == null
				|| result.getVarType().intValue() == Variant.VT_EMPTY
				|| result.getVarType().intValue() == Variant.VT_NULL
				|| result.getVarType().intValue() != Variant.VT_DISPATCH) {
			
			return null;

		} else {
			return new Conversation((IDispatch) result.getValue());
		}
	}
	
	/**
	 * Returns a boolean value that indicates whether the Item is marked as a
	 * task. Read-only.
	 * <p>
	 * Calling this method sets the value of several other properties, depending
	 * on the value provided in MarkInterval. For more information about the
	 * properties set by specifying MarkInterval, see MarkInterval Enumeration.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a boolean value that indicates whether the Item is marked as a
	 *         task.
	 */
	public boolean isMarkedAsTask() {
		
		return getBooleanProperty("IsMarkedAsTask");
	}
	
	/**
	 * Marks an Item object as a task and assigns a task interval for the
	 * object.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param interval
	 *            MarkInterval The task interval for the Item.
	 */
	public void markAsTask(MarkInterval interval) {
		
		invokeNoReply("MarkAsTask", newVariant(interval.value()));
	}
	
	/**
	 * Returns or sets a boolean value that is true if a reminder has been set for this item. Read/write.
	 * 
	 * @return a boolean value that is true if a reminder has been set for this item.
	 */
	public boolean isReminderSet() {
		
		return getBooleanProperty("ReminderSet");
	}
	
	/**
	 * Returns or sets a boolean value that is true if a reminder has been set
	 * for this item. Read/write.
	 * 
	 * @param flag
	 *            a boolean value that is true if a reminder has been set for
	 *            this item.
	 */
	public void setReminder(boolean flag) {
		
		setProperty("ReminderSet", flag);
	}
	
	/**
	 * Returns or sets a boolean value that is true if the reminder overrides
	 * the default reminder behaviour for the item. Read/write.
	 * <p>
	 * You must set the ReminderOverrideDefault property to validate the
	 * ReminderPlaySound and the ReminderSoundFile properties.
	 * </p>
	 * 
	 * @return a boolean value that is true if the reminder overrides the
	 *         default reminder behaviour for the item.
	 */
	public boolean getReminderOverrideDefault() {
		
		return getBooleanProperty("ReminderOverrideDefault");
	}
	
	/**
	 * Returns or sets a boolean value that is true if the reminder overrides
	 * the default reminder behaviour for the item. Read/write.
	 * <p>
	 * You must set the ReminderOverrideDefault property to validate the
	 * ReminderPlaySound and the ReminderSoundFile properties.
	 * </p>
	 * 
	 * @param flag
	 *            a boolean value that is true if the reminder overrides the
	 *            default reminder behaviour for the item.
	 */
	public void setReminderOverrideDefault(boolean flag) {
		
		setProperty("ReminderOverrideDefault", flag);
	}
	
	/**
	 * Returns or sets a boolean value that is true if the reminder should play
	 * a sound when it occurs for this item. Read/write.
	 * <p>
	 * The ReminderPlaySound property must be set in order to validate the
	 * ReminderSoundFile property.
	 * </p>
	 * 
	 * @return a boolean value that is true if the reminder should play a sound
	 *         when it occurs for this item.
	 */
	public boolean isReminderPlaySound() {
		
		return getBooleanProperty("ReminderPlaySound");
	}
	
	/**
	 * Returns or sets a boolean value that is true if the reminder should play
	 * a sound when it occurs for this item. Read/write.
	 * <p>
	 * The ReminderPlaySound property must be set in order to validate the
	 * ReminderSoundFile property.
	 * </p>
	 * 
	 * @param flag
	 *            a boolean value that is true if the reminder should play a
	 *            sound when it occurs for this item.
	 */
	public void setReminderPlaySound(boolean flag) {
		
		setProperty("ReminderPlaySound", flag);
	}
	
	/**
	 * Returns or sets a String indicating the path and file name of the sound
	 * file to play when the reminder occurs for the Outlook item. Read/write.
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
	 * Returns or sets a Date indicating the date and time at which the reminder
	 * should occur for the specified item (item: An item is the basic element
	 * that holds information in Outlook (similar to a file in other programs).
	 * Items include e-mail messages, appointments, contacts, tasks, journal
	 * entries, notes, posted items, and documents.). Read/write.
	 * 
	 * @return a Date indicating the date and time at which the reminder should
	 *         occur for the specified item.
	 */
	public Date getReminderTime() {
		
		return getDateProperty("ReminderTime");
	}
	
	/**
	 * Returns or sets a Date indicating the date and time at which the reminder
	 * should occur for the specified item (item: An item is the basic element
	 * that holds information in Outlook (similar to a file in other programs).
	 * Items include e-mail messages, appointments, contacts, tasks, journal
	 * entries, notes, posted items, and documents.). Read/write.
	 * 
	 * @param start
	 *            a Date indicating the date and time at which the reminder
	 *            should occur for the specified item.
	 */
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
	
	/**
	 * Returns or sets a Date value that represents the completion date of the
	 * task for this Item. Read/write.
	 * <p>
	 * This property returns Null (Nothing in Visual Basic) if the
	 * IsMarkedAsTask property is set to false.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a Date value that represents the completion date of the task for
	 *         this Item.
	 */
	public Date getTaskCompletedDate() {
		
		return getDateProperty("TaskCompletedDate");
	}
	
	/**
	 * Returns or sets a Date value that represents the completion date of the
	 * task for this Item. Read/write.
	 * <p>
	 * This property returns Null (Nothing in Visual Basic) if the
	 * IsMarkedAsTask property is set to false.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param dat
	 *            a Date value that represents the completion date of the task
	 *            for this Item.
	 */
	public void setTaskCompletedDate(Date dat) {
		
		setProperty("TaskCompletedDate", dat);
	}
	
	/**
	 * Returns or sets a Date value that represents the due date of the task for
	 * this Item. Read/write.
	 * <p>
	 * This property returns Null (Nothing in Visual Basic) if the
	 * IsMarkedAsTask property is set to false.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a Date value that represents the due date of the task for this
	 *         Item.
	 */
	public Date getTaskDueDate() {
		
		return getDateProperty("TaskDueDate");
	}
	
	/**
	 * Returns or sets a Date value that represents the due date of the task for
	 * this Item. Read/write.
	 * <p>
	 * This property returns Null (Nothing in Visual Basic) if the
	 * IsMarkedAsTask property is set to false.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param dat
	 *            a Date value that represents the due date of the task for this
	 *            Item.
	 */
	public void setTaskDueDate(Date dat) {
		
		setProperty("TaskDueDate", dat);
	}
	
	/**
	 * Returns or sets a Date value that represents the start date of the task
	 * for this Item object. Read/write.
	 * <p>
	 * This property returns Null (Nothing in Visual Basic) if the
	 * IsMarkedAsTask property is set to false.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a Date value that represents the start date of the task for this
	 *         Item object.
	 */
	public Date getTaskStartDate() {
		
		return getDateProperty("TaskStartDate");
	}
	
	/**
	 * Returns or sets a Date value that represents the start date of the task
	 * for this Item object. Read/write.
	 * <p>
	 * This property returns Null (Nothing in Visual Basic) if the
	 * IsMarkedAsTask property is set to false.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param dat
	 *            a Date value that represents the start date of the task for
	 *            this Item object.
	 */
	public void setTaskStartDate(Date dat) {
		
		setProperty("TaskStartDate", dat);
	}
	
	/**
	 * Returns or sets a String value that represents the subject of the task
	 * for the Item object. Read/write.
	 * <p>
	 * This property returns the value of the Subject property if the
	 * IsMarkedAsTask property is set to false.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a String value that represents the subject of the task for the
	 *         Item object.
	 */
	public String getTaskSubject() {
		
		return getStringProperty("TaskSubject");
	}
	
	/**
	 * Returns or sets a String value that represents the subject of the task
	 * for the Item object. Read/write.
	 * <p>
	 * This property returns the value of the Subject property if the
	 * IsMarkedAsTask property is set to false.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param subject
	 *            a String value that represents the subject of the task for the
	 *            Item object.
	 */
	public void setTaskSubject(String subject) {
		
		setProperty("TaskSubject", subject);
	}
	
	/**
	 * Returns or sets a Date value that represents the ordinal value of the
	 * task for the Item. Read/write.
	 * <p>
	 * This property returns Null (Nothing in Visual Basic) if the
	 * IsMarkedAsTask property is set to false.
	 * </p>
	 * <p>
	 * This property is used to indicate how the task should be ordered within
	 * the parent groups, such as the Today group or the Tomorrow group, of the
	 * To-Do Bar. The value used in this property does not have any relation to
	 * the values of the TaskStartDate, TaskDueDate, or TaskCompletedDate
	 * properties.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a Date value that represents the ordinal value of the task for
	 *         the Item.
	 */
	public Date getToDoTaskOrdinal() {
		
		return getDateProperty("ToDoTaskOrdinal");
	}
	
	/**
	 * Returns or sets a Date value that represents the ordinal value of the
	 * task for the Item. Read/write.
	 * <p>
	 * This property returns Null (Nothing in Visual Basic) if the
	 * IsMarkedAsTask property is set to false.
	 * </p>
	 * <p>
	 * This property is used to indicate how the task should be ordered within
	 * the parent groups, such as the Today group or the Tomorrow group, of the
	 * To-Do Bar. The value used in this property does not have any relation to
	 * the values of the TaskStartDate, TaskDueDate, or TaskCompletedDate
	 * properties.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param dat
	 *            a Date value that represents the ordinal value of the task for
	 *            the Item.
	 */
	public void setToDoTaskOrdinal(Date dat) {
		
		setProperty("ToDoTaskOrdinal", dat);
	}
	
}
