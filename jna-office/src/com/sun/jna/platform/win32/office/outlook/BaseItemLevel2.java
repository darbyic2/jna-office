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

import com.sun.jna.platform.win32.COM.IDispatch;

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
 * @see BaseItemLevel1
 * @see BaseItemLevel3
 * @see DocumentItem
 * @see RemoteItem
 * @see ReportItem
 * @see JournalItem
 * @see MeetingItem
 * @see AppointmentItem
 * @see BaseTaskItem
 */
public abstract class BaseItemLevel2 extends BaseItemLevel1 {

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
	protected BaseItemLevel2(IDispatch iDisp) {
		super(iDisp);
	}
	
	/**
	 * Returns an Actions collection that represents all the available actions
	 * for the item (item: An item is the basic element that holds information
	 * in Outlook (similar to a file in other programs). Items include e-mail
	 * messages, appointments, contacts, tasks, journal entries, notes, posted
	 * items, and documents.). Read-only.
	 * 
	 * @return an Actions collection that represents all the available actions
	 *         for the item.
	 */
	public Actions getActions() {
		
		return new Actions(getAutomationProperty("Actions"));
	}
	
	/**
	 * Returns an Attachments object that represents all the attachments for the
	 * specified item (item: An item is the basic element that holds information
	 * in Outlook (similar to a file in other programs). Items include e-mail
	 * messages, appointments, contacts, tasks, journal entries, notes, posted
	 * items, and documents.). Read-only.
	 * 
	 * @return an Attachments object that represents all the attachments for the
	 *         specified item.
	 */
	public Attachments getAttachments() {
		
		return new Attachments(getAutomationProperty("Attachments"));
	}
	
	/**
	 * Returns a boolean that determines if the item is a winner of an automatic
	 * conflict resolution. Read-only.
	 * <p>
	 * A value of False does not necessarily indicate that the item is a loser
	 * of an automatic conflict resolution. The item could be in conflict with
	 * another item.
	 * </p>
	 * <p>
	 * If an item has Conflicts.Count of its DocumentItem.Conflicts property
	 * greater than zero and if its AutoResolvedWinner property is true, it is a
	 * winner of an automatic conflict resolution. On the other hand, if the
	 * item is in conflict and has its AutoResolvedWinner property as false, it
	 * is a loser in an automatic conflict resolution.
	 * </p>
	 * 
	 * @return a boolean that determines if the item is a winner of an automatic
	 *         conflict resolution.
	 */
	public boolean isAutoResolvedWinner() {
		
		return getBooleanProperty("AutoResolvedWinner");
	}
	
	/**
	 * Returns a String representing the billing information associated
	 * with the Outlook item (item: An item is the basic element that holds
	 * information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.). Read/write.
	 * <p>
	 * This is a free-form text field.
	 * </p>
	 * 
	 * @return a String representing the billing information associated with the
	 *         Outlook item.
	 */
	public String getBillingInformation() {
		
		return getStringProperty("BillingInformation");
	}
	
	/**
	 * Sets a String representing the billing information associated
	 * with the Outlook item (item: An item is the basic element that holds
	 * information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.). Read/write.
	 * <p>
	 * This is a free-form text field.
	 * </p>
	 * 
	 * @param billingInfo
	 *            a String representing the billing information associated with
	 *            the Outlook item.
	 */
	public void setBillingInformation(String billingInfo) {
		
		setProperty("BillingInformation", billingInfo);
	}
	
	/**
	 * Returns or sets a String representing the names of the companies
	 * associated with the Outlook item (item: An item is the basic element that
	 * holds information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.). Read/write.
	 * <p>
	 * This is a free-form text field.
	 * </p>
	 * 
	 * @return a String representing the names of the companies associated with
	 *         the Outlook item.
	 */
	public String getCompanies() {
		
		return getStringProperty("Companies");
	}
	
	/**
	 * Returns or sets a String representing the names of the companies
	 * associated with the Outlook item (item: An item is the basic element that
	 * holds information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.). Read/write.
	 * <p>
	 * This is a free-form text field.
	 * </p>
	 * 
	 * @param companies
	 *            a String representing the names of the companies associated
	 *            with the Outlook item.
	 */
	public void setCompanies(String companies) {
		
		setProperty("Companies", companies);
	}
	
	/**
	 * Returns a boolean that determines if the item is in conflict. Read-only.
	 * <p>
	 * Whether or not an item is in conflict is determined by the state of the
	 * application. For example, when a user is offline and tries to access an
	 * online folder the action will fail. In this scenario, the IsConflict
	 * property will return true.
	 * </p>
	 * <p>
	 * If True, the specified item is in conflict.
	 * </p>
	 * 
	 * @return a boolean that determines if the item is in conflict.
	 */
	public boolean isConflict() {
		
		return getBooleanProperty("IsConflict");
	}
	
	/**
	 * Return the Conflicts object that represents the items that are in
	 * conflict for any Outlook item object. Read-only.
	 * 
	 * @return Conflicts object that represents the items that are in conflict
	 *         for any Outlook item object.
	 */
	public Conflicts getConflicts() {
		
		return new Conflicts(getAutomationProperty("Conflicts"));
	}
	
	/**
	 * Returns a String that uniquely identifies a Conversation object that the
	 * Item object belongs to. Read-only.
	 * <p>
	 * <strong> Does not apply to the DocumentItem object. </strong>
	 * </p>
	 * <p>
	 * This property associates items with a conversation. These items and the
	 * conversation all have the same value in their ConversationID property.
	 * </p>
	 * <p>
	 * This property corresponds with the MAPI property PidTagConversationId.
	 * </p>
	 * <p>
	 * If the Item object is created in a version of Microsoft Outlook
	 * earlier than Microsoft Outlook 2010, or if Outlook is running in online
	 * mode against a version of Microsoft Exchange Server earlier than
	 * Microsoft Exchange Server 2010, this property returns the same value as
	 * the ConversationTopic property.
	 * </p>
	 * 
	 * @return a String that uniquely identifies a Conversation object that the
	 *         Item object belongs to.
	 */
	public String getConversationID() {
		
		return getStringProperty("ConversationID");
	}
	
	/**
	 * Returns a String that indicates the relative position of the item (item:
	 * An item is the basic element that holds information in Outlook (similar
	 * to a file in other programs). Items include e-mail messages,
	 * appointments, contacts, tasks, journal entries, notes, posted items, and
	 * documents.) within the conversation thread. Read-only.
	 * <p>
	 * This property corresponds to the MAPI property PidTagConversationIndex.
	 * </p>
	 * 
	 * @return a String that indicates the relative position of the item.
	 */
	public String getConversationIndex() {
		
		return getStringProperty("ConversationIndex");
	}
	
	/**
	 * Returns a String representing the topic of the conversation thread of the
	 * Outlook item (item: An item is the basic element that holds information
	 * in Outlook (similar to a file in other programs). Items include e-mail
	 * messages, appointments, contacts, tasks, journal entries, notes, posted
	 * items, and documents.). Read-only.
	 * 
	 * @return a String representing the topic of the conversation thread of the
	 *         Outlook item.
	 */
	public String getConversationTopic() {
		
		return getStringProperty("ConversationTopic");
	}
	
	/**
	 * Returns a constant that belongs to the DownloadState enumeration
	 * indicating the download state of the item. Read-only.
	 * 
	 * @return a constant that belongs to the DownloadState enumeration
	 *         indicating the download state of the item.
	 */
	public ItemDownloadState getDownloadState() {
		
		return ItemDownloadState.parse(getShortProperty("DownloadState"));
	}
	
	/**
	 * Returns the FormDescription object that represents the form description
	 * for the specified Outlook item (item: An item is the basic element that
	 * holds information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.). Read-only.
	 * 
	 * @return the FormDescription object that represents the form description
	 *         for the specified Outlook item.
	 */
	public FormDescription getFormDescription() {
		
		return new FormDescription(getAutomationProperty("FormDescription"));
	}
	
	/**
	 * Returns an Importance constant indicating the relative importance level
	 * for the Outlook item (item: An item is the basic element that holds
	 * information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.). Read/write.
	 * <p>
	 * This property corresponds to the MAPI property PidTagImportance.
	 * </p>
	 * 
	 * @return an Importance constant indicating the relative importance level
	 *         for the Outlook item.
	 */
	public Importance getImportance() {
		
		return Importance.parse(getShortProperty("Importance"));
	}
	
	/**
	 * Sets an Importance constant indicating the relative importance level for
	 * the Outlook item (item: An item is the basic element that holds
	 * information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.). Read/write.
	 * <p>
	 * This property corresponds to the MAPI property PidTagImportance.
	 * </p>
	 * 
	 * @param level
	 *            an Importance constant indicating the relative importance
	 *            level for the Outlook item.
	 */
	public void setImportance(Importance level) {
		
		setProperty("Importance", level.value());
	}
	
	/**
	 * Returns a Links collection that represents the contacts to which the item
	 * is linked. Read-only.
	 * 
	 * @return a Links collection that represents the contacts to which the item
	 *         is linked.
	 */
	public Links getLinks() {
		
		return new Links(getAutomationProperty("Links"));
	}
	
	/**
	 * Returns a RemoteStatus constant that determines the status of an
	 * item once it is received by a remote user. Read/write.
	 * <p>
	 * This property gives remote users with less-than-ideal data-transfer
	 * capabilities increased messaging flexibility.
	 * </p>
	 * 
	 * @return a RemoteStatus constant that determines the status of an item
	 *         once it is received by a remote user.
	 */
	public RemoteStatus getMarkForDownload() {
		
		return RemoteStatus.parse(getShortProperty("MarkForDownload"));
	}
	
	/**
	 * Sets a RemoteStatus constant that determines the status of an
	 * item once it is received by a remote user. Read/write.
	 * <p>
	 * This property gives remote users with less-than-ideal data-transfer
	 * capabilities increased messaging flexibility.
	 * </p>
	 * 
	 * @param status
	 *            a RemoteStatus constant that determines the status of an item
	 *            once it is received by a remote user.
	 */
	public void setMarkForDownload(RemoteStatus status) {
		
		setProperty("MarkForDownload", status.value());
	}
	
	/**
	 * Returns a String representing the mileage for an item (item: An item is
	 * the basic element that holds information in Outlook (similar to a file in
	 * other programs). Items include e-mail messages, appointments, contacts,
	 * tasks, journal entries, notes, posted items, and documents.). Read/write.
	 * <p>
	 * This is a free-form string field and can be used to store mileage
	 * information associated with the item (for example, 100 miles documented
	 * for an appointment, contact, or task) for purposes of reimbursement.
	 * </p>
	 * 
	 * @return a String representing the mileage for an item.
	 */
	public String getMileage() {
		
		return getStringProperty("Mileage");
	}
	
	/**
	 * Sets a String representing the mileage for an item (item: An item is the
	 * basic element that holds information in Outlook (similar to a file in
	 * other programs). Items include e-mail messages, appointments, contacts,
	 * tasks, journal entries, notes, posted items, and documents.). Read/write.
	 * <p>
	 * This is a free-form string field and can be used to store mileage
	 * information associated with the item (for example, 100 miles documented
	 * for an appointment, contact, or task) for purposes of reimbursement.
	 * </p>
	 * 
	 * @param freeFormText
	 *            a String representing the mileage for an item.
	 */
	public void setMileage(String freeFormText) {
		
		setProperty("Mileage", freeFormText);
	}
	
	/**
	 * Returns or sets a boolean value that is True to not age the Outlook item
	 * (item: An item is the basic element that holds information in Outlook
	 * (similar to a file in other programs). Items include e-mail messages,
	 * appointments, contacts, tasks, journal entries, notes, posted items, and
	 * documents.). Read/write.
	 * 
	 * @return a boolean value that is True to not age the Outlook item.
	 */
	public boolean isNoAging() {
		
		return getBooleanProperty("NoAging");
	}
	
	/**
	 * Returns or sets a boolean value that is True to not age the Outlook item
	 * (item: An item is the basic element that holds information in Outlook
	 * (similar to a file in other programs). Items include e-mail messages,
	 * appointments, contacts, tasks, journal entries, notes, posted items, and
	 * documents.). Read/write.
	 * 
	 * @param flag
	 *            a boolean value that is True to not age the Outlook item.
	 */
	public void setNoAging(boolean flag) {
		
		setProperty("NoAging", flag);
	}
	
	/**
	 * Returns an int representing the build number of the Outlook application
	 * for an Outlook item (item: An item is the basic element that holds
	 * information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.). Read-only.
	 * 
	 * @return an int representing the build number of the Outlook application
	 *         for an Outlook item.
	 */
	public int getOutlookInternalVersion() {
		
		return getIntProperty("OutlookInternalVersion");
	}
	
	/**
	 * Returns a String indicating the major and minor version number of the
	 * Outlook application for an Outlook item (item: An item is the basic
	 * element that holds information in Outlook (similar to a file in other
	 * programs). Items include e-mail messages, appointments, contacts, tasks,
	 * journal entries, notes, posted items, and documents.). Read-only.
	 * 
	 * @return a String indicating the major and minor version number of the
	 *         Outlook application for an Outlook item.
	 */
	public String getOutlookVersion() {
		
		return getStringProperty("OutlookVersion");
	}
	
	/**
	 * Returns a constant in the Sensitivity enumeration indicating the
	 * sensitivity for the Outlook item. Read/write.
	 * <p>
	 * This property corresponds to the MAPI property PidTagSensitivity.
	 * </p>
	 * 
	 * @return a constant in the Sensitivity enumeration indicating the
	 *         sensitivity for the Outlook item.
	 */
	public Sensitivity getSensitivity() {
		
		return Sensitivity.parse(getShortProperty("Sensitivity"));
	}
	
	/**
	 * Sets a constant in the Sensitivity enumeration indicating the sensitivity
	 * for the Outlook item. Read/write.
	 * <p>
	 * This property corresponds to the MAPI property PidTagSensitivity.
	 * </p>
	 * 
	 * @param level
	 *            a constant in the Sensitivity enumeration indicating the
	 *            sensitivity for the Outlook item.
	 */
	public void setSensitivity(Sensitivity level) {
		
		setProperty("Sensitivity", level.value());
	}
	
	/**
	 * Displays the Show Categories dialog box, which allows you to select
	 * categories that correspond to the subject of the item.
	 */
	public void showCategoriesDialog() {
		
		invokeNoReply("ShowCategoriesDialog");
	}
	
	/**
	 * Returns a boolean value that is True if the Outlook item has not been
	 * opened (read). Read/write.
	 * 
	 * @return a boolean value that is True if the Outlook item has not been
	 *         opened (read).
	 */
	public boolean isUnRead() {
		
		return getBooleanProperty("UnRead");
	}
	
	/**
	 * Sets a boolean value that is True if the Outlook item has not been opened
	 * (read). Read/write.
	 * 
	 * @param flag
	 *            a boolean value that is True if the Outlook item has not been
	 *            opened (read).
	 */
	public void setUnread(boolean flag) {
		
		setProperty("Unread", flag);
	}
	
	/**
	 * Returns the UserProperties collection that represents all the user
	 * properties for the Outlook item. Read-only.
	 * <p>
	 * Even though olWordDocumentItem is a valid constant in the OlItemType
	 * enumeration, user-defined fields cannot to be added to a DocumentItem
	 * object and you will receive an error when you try to programmatically add
	 * a user-defined field to a DocumentItem object.
	 * </p>
	 * 
	 * @return the UserProperties collection that represents all the user
	 *         properties for the Outlook item.
	 */
	public UserProperties getUserProperties() {
		
		return new UserProperties(getAutomationProperty("UserProperties"));
	}
	
}
