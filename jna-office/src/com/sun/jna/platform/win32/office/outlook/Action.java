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

import com.sun.jna.platform.win32.OleAuto;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant.VARIANT;
import com.sun.jna.platform.win32.WinDef.LONG;

/**
 * Represents a specialised action (for example, the voting options response)
 * that can be executed on an Outlook item (item: An item is the basic element
 * that holds information in Outlook (similar to a file in other programs).
 * Items include e-mail messages, appointments, contacts, tasks, journal
 * entries, notes, posted items, and documents.).
 * <p>
 * The Action object is a member of the Actions collection.
 * </p>
 * <p>
 * Use getActions(index), where index is the name of an available action, to
 * return a single Action object from the Actions collection object of an
 * Outlook item, such as a MailItem.
 * </p>
 * 
 * @author Ian Darby
 * 
 * @see {@link BaseOutlookObject}
 */
public class Action extends BaseOutlookObject {

	/**
	 * Constructor scope is restricted to package as it should not be used
	 * directly by user applications. It is only intended to be used from within
	 * factory methods and properties of the Outlook object model itself. It may
	 * also be called from unit tests which may supply a mock version of the
	 * IDispatch object.
	 * 
	 * @param iDisp
	 *            the IDispatch object which is the underlying Action object
	 *            within the Outlook object model. All methods and properties of
	 *            this wrapper class ultimately delegate to IDispatch.
	 */
	public Action(IDispatch iDisp) {
		super(iDisp);
	}

	/**
	 * @return an ActionCopyLike constant indicating the property inheritance
	 *         style to use for the action. Read/write.
	 * 
	 * @see #setCopyLike(ActionCopyLike)
	 */
	public ActionCopyLike getCopyLike() {

		return ActionCopyLike.parse(getShortProperty("CopyLike"));
	}

	/**
	 * Sets an ActionCopyLike constant indicating the property inheritance style
	 * to use for the action. Read/write.
	 * 
	 * @param actionType
	 *            an ActionCopyLike constant.
	 * 
	 * @see #getCopyLike()
	 */
	public void setCopyLike(ActionCopyLike actionType) {

		setProperty("CopyLike", actionType.value());
	}

	/**
	 * Deletes itself from the collection.
	 */
	public void delete() {

		invokeNoReply("Delete");
	}

	/**
	 * @return a Boolean that is True if the action is enabled in the
	 *         application. Read/write.
	 * 
	 * @see #setEnabled(boolean)
	 */
	public boolean isEnabled() {

		return getBooleanProperty("Enabled");
	}

	/**
	 * Sets a Boolean that is True if the action is enabled in the application.
	 * Read/write.
	 * 
	 * @param flag
	 *            value to set the enabled property to.
	 * 
	 * @see #isEnabled()
	 */
	public void setEnabled(boolean flag) {

		setProperty("Enabled", flag);
	}

	/**
	 * Helper factory method to generate a wrapped instance of the given
	 * IDispatch object.
	 * 
	 * @param iDisp
	 *            IDispatch object that is to be wrapped in a Java class
	 *            representation of the represented Outlook object model object.
	 * 
	 * @return a wrapped instance of the given IDispatch object.
	 * 
	 * @see #execute()
	 */
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

		case ClassEnum.olDistributionList:
			return new DistributionListItem(iDisp);
			
		case ClassEnum.olDocument:
			return new DocumentItem(iDisp);
			
		case ClassEnum.olMobile:
			return new MobileItem(iDisp);
			
		case ClassEnum.olNote:
			return new NoteItem(iDisp);
			
		case ClassEnum.olRemote:
			return new RemoteItem(iDisp);
			
		case ClassEnum.olSharing:
			return new SharingItem(iDisp);
			
		case ClassEnum.olTaskRequest:
			return new TaskRequestItem(iDisp);
			
		case ClassEnum.olTaskRequestAccept:
			return new TaskRequestAcceptItem(iDisp);
			
		case ClassEnum.olTaskRequestDecline:
			return new TaskRequestDecllinedItem(iDisp);
			
		case ClassEnum.olTaskRequestUpdate:
			return new TaskRequestUpdatedItem(iDisp);

		//The following are currently not supported options.
		case ClassEnum.olMeetingCancellation:
		case ClassEnum.olMeetingForwardNotification:
		case ClassEnum.olMeetingRequest:
		case ClassEnum.olMeetingResponseNegative:
		case ClassEnum.olMeetingResponsePositive:
		case ClassEnum.olMeetingResponseTentative:
		case ClassEnum.olReport:
		default:
			throw new RuntimeException(
					"Action.execute() not yet implemented for class ID: "
							+ classId);
		}
	}

	/**
	 * Executes the action for the specified item (item: An item is the basic
	 * element that holds information in Outlook (similar to a file in other
	 * programs). Items include e-mail messages, appointments, contacts, tasks,
	 * journal entries, notes, posted items, and documents.).
	 * 
	 * @return an Object value that represents the Outlook item created by the
	 *         action upon execution.
	 */
	public BaseItemLevel1 execute() {

		return wrappedObject((IDispatch) invoke("Execute").getValue());
	}

	/**
	 * Returns a String representing the message class for the Action.
	 * Read/write.
	 * <p>
	 * This property corresponds to the MAPI property PidTagMessageClass. The
	 * MessageClass property links the item (item: An item is the basic element
	 * that holds information in Outlook (similar to a file in other programs).
	 * Items include e-mail messages, appointments, contacts, tasks, journal
	 * entries, notes, posted items, and documents.) to the form on which it is
	 * based. When an item is selected, Outlook uses the message class to locate
	 * the form and expose its properties, such as Reply commands.
	 * </p>
	 * 
	 * @return a String representing the message class for the Action.
	 *         Read/write.
	 */
	public String getMessageClass() {

		return getStringProperty("MessageClass");
	}

	/**
	 * Sets a String representing the message class for the Action. Read/write.
	 * <p>
	 * This property corresponds to the MAPI property PidTagMessageClass. The
	 * MessageClass property links the item (item: An item is the basic element
	 * that holds information in Outlook (similar to a file in other programs).
	 * Items include e-mail messages, appointments, contacts, tasks, journal
	 * entries, notes, posted items, and documents.) to the form on which it is
	 * based. When an item is selected, Outlook uses the message class to locate
	 * the form and expose its properties, such as Reply commands.
	 * </p>
	 * 
	 * @param messageClass
	 *            String representing the message class for the Action.
	 */
	public void setMessageClass(String messageClass) {

		setProperty("MessageClass", messageClass);
	}

	/**
	 * @return the display name for the object. Read/write.
	 */
	public String getName() {

		return getStringProperty("Name");
	}

	/**
	 * Sets the display name for the object. Read/write.
	 * 
	 * @param name
	 *            the display name to use for the object.
	 */
	public void setName(String name) {

		setProperty("Name", name);
	}

	/**
	 * Returns a String specifying the prefix (for example, "Re") to use with
	 * the subject of the item when the action is executed. Read/write.
	 * 
	 * @return a String specifying the prefix (for example, "Re") to use with
	 *         the subject of the item when the action is executed. Read/write.
	 */
	public String getPrefix() {

		return getStringProperty("Prefix");
	}

	/**
	 * Sets a String specifying the prefix (for example, "Re") to use with the
	 * subject of the item when the action is executed. Read/write.
	 * 
	 * @param prefix
	 *            string prefix to use for the subject.
	 */
	public void setPrefix(String prefix) {

		setProperty("Prefix", prefix);
	}

	/**
	 * Returns an ActionReplyStyle constant indicating the text formatting reply
	 * style for the specified action. Read/write.
	 * 
	 * @return an ActionReplyStyle constant indicating the text formatting reply
	 *         style for the specified action. Read/write.
	 */
	public ActionReplyStyle getReplyStyle() {

		return ActionReplyStyle.parse(getShortProperty("ReplyStyle"));
	}

	/**
	 * Sets an ActionReplyStyle constant indicating the text formatting reply
	 * style for the specified action. Read/write.
	 * 
	 * @param style
	 *            ActionReplyStyle constant indicating the text formatting reply
	 *            style for the specified action.
	 */
	public void setReplyStyle(ActionReplyStyle style) {

		setProperty("ReplyStyle", style.value());
	}

	/**
	 * Returns an ActionResponseStyle constant indicating the response style
	 * used when the specified action is executed. Read/write.
	 * 
	 * @return an ActionResponseStyle constant indicating the response style
	 *         used when the specified action is executed. Read/write.
	 */
	public ActionResponseStyle getResponseStyle() {

		return ActionResponseStyle.parse(getShortProperty("ResponseStyle"));
	}

	/**
	 * Sets an ActionResponseStyle constant indicating the response style used
	 * when the specified action is executed. Read/write.
	 * 
	 * @param style
	 *            ActionResponseStyle constant indicating the response style
	 *            used when the specified action is executed.
	 */
	public void setResponseStyle(ActionResponseStyle style) {

		setProperty("ResponseStyle", style.value());
	}

	/**
	 * Returns an ActionShowOn constant representing the location where the
	 * action will be shown. Read/write.
	 * 
	 * @return an ActionShowOn constant representing the location where the
	 *         action will be shown. Read/write.
	 */
	public ActionShowOn getShowOn() {

		return ActionShowOn.parse(getShortProperty("ShowOn"));
	}

	/**
	 * Sets an ActionShowOn constant representing the location where the action
	 * will be shown. Read/write.
	 * 
	 * @param style
	 *            ActionShowOn constant representing the location where the
	 *            action will be shown.
	 */
	public void setShowOn(ActionShowOn style) {

		setProperty("ShowOn", style.value());
	}

}
