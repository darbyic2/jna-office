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
 * All Item classes derive from this class. There are an inordinate amount of
 * methods and properties that are shared between multiple Item types. However,
 * the commonality model is complex and there is a fine line between reducing
 * duplication and having a ridiculous depth to the inheritance model. The
 * compromise struck here was to have four levels of base class. Item classes
 * inherit from the most appropriate level.
 * 
 * @author Ian Darby
 * 
 * @see BaseOutlookObject
 * @see BaseItemLevel2
 * @see NoteItem
 * @see MobileItem
 */
public class BaseItemLevel1 extends BaseOutlookObject {

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
	protected BaseItemLevel1(IDispatch iDispatch) {
		super(iDispatch);
	}

	/**
	 * Returns or sets a String representing the clear-text body of the Outlook
	 * item (item: An item is the basic element that holds information in
	 * Outlook (similar to a file in other programs). Items include e-mail
	 * messages, appointments, contacts, tasks, journal entries, notes, posted
	 * items, and documents.). Read/write.
	 * 
	 * @return a String representing the clear-text body of the Outlook item.
	 */
	public String getBody() {

		return getStringProperty("Body");
	}

	/**
	 * Returns or sets a String representing the clear-text body of the Outlook
	 * item (item: An item is the basic element that holds information in
	 * Outlook (similar to a file in other programs). Items include e-mail
	 * messages, appointments, contacts, tasks, journal entries, notes, posted
	 * items, and documents.). Read/write.
	 * 
	 * @param body
	 *            a String representing the clear-text body of the Outlook item.
	 */
	public void setBody(String body) {

		setProperty("Body", body);
	}

	/**
	 * Returns or sets a String representing the categories assigned to the
	 * Outlook item (item: An item is the basic element that holds information
	 * in Outlook (similar to a file in other programs). Items include e-mail
	 * messages, appointments, contacts, tasks, journal entries, notes, posted
	 * items, and documents.). Read/write.
	 * <p>
	 * Categories is a delimited string of category names that have been
	 * assigned to an Outlook item. This property uses the character specified
	 * in the value name, sList, under HKEY_CURRENT_USER\Control
	 * Panel\International in the Windows registry, as the delimiter for
	 * multiple categories.
	 * </p>
	 * 
	 * @return a String representing the categories assigned to the Outlook
	 *         item.
	 */
	public String getCategories() {

		return getStringProperty("Categories");
	}

	/**
	 * Returns or sets a String representing the categories assigned to the
	 * Outlook item (item: An item is the basic element that holds information
	 * in Outlook (similar to a file in other programs). Items include e-mail
	 * messages, appointments, contacts, tasks, journal entries, notes, posted
	 * items, and documents.). Read/write.
	 * <p>
	 * Categories is a delimited string of category names that have been
	 * assigned to an Outlook item. This property uses the character specified
	 * in the value name, sList, under HKEY_CURRENT_USER\Control
	 * Panel\International in the Windows registry, as the delimiter for
	 * multiple categories.
	 * </p>
	 * 
	 * @param categoryList
	 *            a String representing the categories assigned to the Outlook
	 *            item.
	 */
	public void setCategories(String categoryList) {

		setProperty("Categories", categoryList);
	}

	/**
	 * Closes and optionally saves changes to the Outlook item.
	 * 
	 * @param option
	 *            The close behaviour. If the item displayed within the inspector
	 *            has not been changed, this argument has no effect.
	 */
	public void close(InspectorCloseOption option) {

		invokeNoReply("Close", newVariant(option.value()));
	}

	/**
	 * Returns a Date indicating the creation time for the Outlook item (item:
	 * An item is the basic element that holds information in Outlook (similar
	 * to a file in other programs). Items include e-mail messages,
	 * appointments, contacts, tasks, journal entries, notes, posted items, and
	 * documents.). Read-only.
	 * <p>
	 * This property corresponds to the MAPI property PidTagCreationTime.
	 * </p>
	 * 
	 * @return a {@link java.util.Date} indicating the creation time for the
	 *         Outlook item.
	 */
	public Date getCreationTime() {

		return getDateProperty("CreationTime");
	}

	/**
	 * Deletes an object from the collection that it is in.
	 * <p>
	 * The Delete mothod deletes a single item in a collection. To delete all
	 * items in the Items collection of a folder, you must delete each item
	 * starting with the last item in the folder. For example, in the items
	 * collection of a folder, AllItems, if there are n number of items in the
	 * folder, start deleting the item at AllItems.Item(n), decrementing the
	 * index each time until you delete AllItems.Item(1).
	 * </p>
	 */
	public void delete() {

		invokeNoReply("Delete");
	}

	/**
	 * Displays a new Inspector object for the item.
	 * <p>
	 * The Display method is supported for explorer and inspector windows for
	 * the sake of backward compatibility. To activate an explorer or inspector
	 * window, use the Activate method.
	 * </p>
	 * <p>
	 * If you attempt to open an "unsafe" file system object (or "freedoc" file)
	 * by using the Microsoft Outlook object model, you receive the E_FAIL
	 * return code in the C or C++ programming languages. In Outlook 2000 and
	 * earlier, you could open an "unsafe" file system object by using the
	 * Display method.
	 * </p>
	 * <p>
	 * Functionally equivalent to calling display(false);
	 * </p>
	 */
	public void display() {

		display(false);
	}

	/**
	 * Displays a new Inspector object for the item.
	 * <p>
	 * The Display method is supported for explorer and inspector windows for
	 * the sake of backward compatibility. To activate an explorer or inspector
	 * window, use the Activate method.
	 * </p>
	 * <p>
	 * If you attempt to open an "unsafe" file system object (or "freedoc" file)
	 * by using the Microsoft Outlook object model, you receive the E_FAIL
	 * return code in the C or C++ programming languages. In Outlook 2000 and
	 * earlier, you could open an "unsafe" file system object by using the
	 * Display method.
	 * </p>
	 * 
	 * @param modal
	 *            True to make the window modal. The default value is False.
	 */
	public void display(boolean modal) {

		invokeNoReply("Display", newVariant(modal));
	}

	/**
	 * Returns a String representing the unique Entry ID of the object.
	 * Read-only.
	 * <p>
	 * This property corresponds to the MAPI property PidTagEntryId.
	 * </p>
	 * <p>
	 * A MAPI store provider assigns a unique ID string when an item is created
	 * in its store. Therefore, the EntryID property is not set for an Outlook
	 * item until it is saved or sent. The Entry ID changes when an item is
	 * moved into another store, for example, from your Inbox to a Microsoft
	 * Exchange Server public folder, or from one Personal Folders (.pst) file
	 * to another .pst file. Solutions should not depend on the EntryID property
	 * to be unique unless items will not be moved. The EntryID property returns
	 * a MAPI long-term Entry ID. For more information about long- and
	 * short-term EntryIDs, search http://msdn.microsoft.com for PidTagEntryId.
	 * </p>
	 * 
	 * @return a String representing the unique Entry ID of the object.
	 */
	public String getEntryID() {

		return getStringProperty("EntryID");
	}

	/**
	 * Returns an Inspector object that represents an inspector initialised to
	 * contain the specified item (item: An item is the basic element that holds
	 * information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.). Read-only.
	 * <p>
	 * This property is useful for returning an Inspector object in which to
	 * display the item, as opposed to using the Application.ActiveInspector
	 * method and setting the Inspector.CurrentItem property. If an Inspector
	 * object already exists for the item, the GetInspector property will return
	 * that Inspector object instead of creating a new one.
	 * </p>
	 * 
	 * @return an Inspector object that represents an inspector initialised to
	 *         contain the specified item.
	 */
	public Inspector getInspector() {

		return new Inspector(getAutomationProperty("GetInspector"));
	}

	/**
	 * Returns an ItemProperties collection that represents all standard and
	 * user-defined properties associated with the Outlook item. Read-only.
	 * <p>
	 * The ItemProperties collection is a zero-based collection, meaning that
	 * the first object in the collection is referenced by the index 0.
	 * </p>
	 * 
	 * @return an ItemProperties collection that represents all standard and
	 *         user-defined properties associated with the Outlook item.
	 */
	public ItemProperties getItemProperties() {

		return new ItemProperties(getAutomationProperty("ItemProperties"));
	}

	/**
	 * Returns a Date specifying the date and time that the Outlook item (item:
	 * An item is the basic element that holds information in Outlook (similar
	 * to a file in other programs). Items include e-mail messages,
	 * appointments, contacts, tasks, journal entries, notes, posted items, and
	 * documents.) was last modified. Read-only.
	 * <p>
	 * This property corresponds to the MAPI property
	 * PidTagLastModificationTime.
	 * </p>
	 * 
	 * @return a {@link java.util.Date} specifying the date and time that the
	 *         Outlook item was last modified.
	 */
	public Date getLastModificationTime() {

		return getDateProperty("LastModificationTime");
	}

	/**
	 * Returns or sets a String representing the message class for the Outlook
	 * item. Read/write.
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
	 * @return a String representing the message class for the Outlook item.
	 */
	public String getMessageClass() {

		return getStringProperty("MessageClass");
	}

	/**
	 * Returns or sets a String representing the message class for the Outlook
	 * item. Read/write.
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
	 * @param classOfMessage
	 *            a String representing the message class for the Outlook item.
	 */
	public void setMessageClass(String classOfMessage) {

		setProperty("MessageClass", classOfMessage);
	}

	/**
	 * Moves a Microsoft Outlookitem (item: An item is the basic element that
	 * holds information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.) to a new folder.
	 * 
	 * @param destFolder
	 *            the Outlook folder to which the item should be moved.
	 * 
	 * @return the item that has been moved. In practice this is a reference to
	 *         itself.
	 */
	public BaseOutlookObject move(Folder destFolder) {

		invokeNoReply("Move", newVariant(destFolder.getIDispatch()));
		return this;
	}

	/**
	 * Prints the Outlook item (item: An item is the basic element that holds
	 * information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.) using all default settings.The
	 * PrintOut method is the only Outlook method that can be used for printing.
	 */
	public void printOut() {

		invokeNoReply("PrintOut");
	}

	/**
	 * Returns a PropertyAccessor object that supports creating, getting,
	 * setting, and deleting properties of the parent object. Read-only.
	 * 
	 * @return a PropertyAccessor object that supports creating, getting,
	 *         setting, and deleting properties of the parent object.
	 */
	public PropertyAccessor getPropertyAccessor() {

		return new PropertyAccessor(getAutomationProperty("PropertyAccessor"));
	}

	/**
	 * Saves the Microsoft Outlookitem (item: An item is the basic element that
	 * holds information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.) to the current folder or, if this is
	 * a new item, to the Outlook default folder for the item type.
	 */
	public void save() {

		invokeNoReply("Save");
	}

	/**
	 * Saves the Microsoft Outlookitem (item: An item is the basic element that
	 * holds information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.) to the specified path and in the
	 * format of the specified file type. If the file type is not specified, the
	 * MSG format (.msg) is used.
	 * <p>
	 * Also note that even though olDoc is a valid OlSaveAsType constant,
	 * messages in HTML format cannot be saved in Document format, and the olDoc
	 * constant works only if Microsoft Word is set up as the default email
	 * editor.
	 * </p>
	 * 
	 * @param filePath
	 *            The path in which to save the item.
	 */
	public void saveAs(String filePath) {

		invokeNoReply("SaveAs", newVariant(filePath));
	}

	/**
	 * Saves the Microsoft Outlookitem (item: An item is the basic element that
	 * holds information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.) to the specified path and in the
	 * format of the specified file type. If the file type is not specified, the
	 * MSG format (.msg) is used.
	 * <p>
	 * Also note that even though olDoc is a valid OlSaveAsType constant,
	 * messages in HTML format cannot be saved in Document format, and the olDoc
	 * constant works only if Microsoft Word is set up as the default email
	 * editor.
	 * </p>
	 * 
	 * @param filePath
	 *            The path in which to save the item.
	 * 
	 * @param typ
	 *            The file type to save. Can be one of the following
	 *            OlSaveAsType constants: olHTML, olMSG, olRTF, olTemplate,
	 *            olDoc, olTXT, olVCal, olVCard, olICal, or olMSGUnicode.
	 */
	public void saveAs(String filePath, SaveAsType typ) {

		invokeNoReply("SaveAs", newVariant(filePath), newVariant(typ.value()));
	}

	/**
	 * Returns a boolean value that is true if the Outlook item has not been
	 * modified since the last save. Read-only.
	 * 
	 * @return a boolean value that is true if the Outlook item has not been
	 *         modified since the last save.
	 */
	public boolean isSaved() {

		return getBooleanProperty("Saved");
	}

	/**
	 * Returns an int indicating the size (in bytes) of the Outlook item.
	 * Read-only.
	 * 
	 * @return an int indicating the size (in bytes) of the Outlook item.
	 */
	public int getSize() {

		return getIntProperty("Size");
	}

	/**
	 * Returns a String indicating the subject for the Outlook item. Read/write
	 * (Read-only in the case of a NoteItem).
	 * <p>
	 * This property corresponds to the MAPI property PidTagSubject. The Subject
	 * property is the default property for Outlook items.
	 * </p>
	 * <p>
	 * In the case of a NoteItem the Subject property is a String that is
	 * calculated from the body text of the note.
	 * </p>
	 * 
	 * @return a String indicating the subject for the Outlook item.
	 */
	public String getSubject() {

		return getStringProperty("Subject");
	}

	/**
	 * Returns a String indicating the subject for the Outlook item. Read/write
	 * (Read-only in the case of a NoteItem).
	 * <p>
	 * This property corresponds to the MAPI property PidTagSubject. The Subject
	 * property is the default property for Outlook items.
	 * </p>
	 * <p>
	 * In the case of a NoteItem the Subject property is a String that is
	 * calculated from the body text of the note.
	 * </p>
	 * 
	 * @param subject
	 *            a String indicating the subject for the Outlook item.
	 */
	public void setSubject(String subject) {

		setProperty("Subject", subject);
	}

}
