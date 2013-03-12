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

import java.io.File;

import com.sun.jna.platform.win32.COM.IDispatch;

/**
 * Represents a document or link to a document contained in an Outlook item
 * (item: An item is the basic element that holds information in Outlook
 * (similar to a file in other programs). Items include e-mail messages,
 * appointments, contacts, tasks, journal entries, notes, posted items, and
 * documents.).
 * <p>
 * Use Attachments.getItem(index), where index is the index number, to return a single
 * Attachment object.
 * </p>
 * <p>
 * Use the Attachments.add method to add an attachment to an item.
 * </p>
 * 
 * @author Ian Darby
 * 
 */
public class Attachment extends BaseOutlookObject {

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
	public Attachment(IDispatch iDisp) {
		super(iDisp);
	}
	
	/**
	 * Returns an OlAttachmentBlockLevel constant that specifies if there is any
	 * restriction on the attachment based on its file extension. Read-only.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return an OlAttachmentBlockLevel constant that specifies if there is any
	 *         restriction on the attachment based on its file extension.
	 */
	public AttachmentBlockLevel getBlockLevel() {
		
		return AttachmentBlockLevel.parse(getShortProperty("BlockLevel"));
	}
	
	/**
	 * Deletes itself from the collection.
	 */
	public void delete() {
		
		invokeNoReply("Delete");
	}
	
	/**
	 * Returns a String representing the name, which does not need to be the
	 * actual file name, displayed below the icon representing the embedded
	 * attachment. Read/write.
	 * <p>
	 * This property corresponds to the MAPI property PidTagDisplayName.
	 * </p>
	 * 
	 * @return a String representing the name, which does not need to be the
	 *         actual file name, displayed below the icon representing the
	 *         embedded attachment.
	 */
	public String getDisplayName() {
		
		return getStringProperty("DisplayName");
	}
	
	/**
	 * Sets a String representing the name, which does not need to be the actual
	 * file name, displayed below the icon representing the embedded attachment.
	 * Read/write.
	 * 
	 * @param name
	 *            a String representing the name, which does not need to be the
	 *            actual file name, displayed below the icon representing the
	 *            embedded attachment.
	 */
	public void setDisplayName(String name) {
		
		setProperty("DisplayName", name);
	}
	
	/**
	 * Returns a String representing the file name of the attachment. Read-only.
	 * 
	 * @return a String representing the file name of the attachment.
	 */
	public String getFileName() {
		
		return getStringProperty("FileName");
	}
	
	/**
	 * Returns the full path to the attached file that is in a temporary-files
	 * folder. Read- only.
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @return the full path to the attached file that is in a temporary-files
	 *         folder.
	 */
	public File getTemporyFilePath() {
		
		return new File((invoke("GetTemporyFilePath").getValue().toString()));
	}
	
	/**
	 * Returns an int indicating the position of the object within the
	 * collection. Read-only.
	 * <p>
	 * The Index property is only valid during the current session (session: A
	 * sequence of operations performed by the Access database engine that
	 * begins when a user logs on and ends when the user logs off. All
	 * operations during a session form one transaction scope and are subject to
	 * the user's logon permissions.) and can change as objects are added to and
	 * deleted from the collection. The first object in the collection has an
	 * Index value of 1.
	 * </p>
	 * 
	 * @return an int indicating the position of the object within the
	 *         collection.
	 */
	public int getIndex() {
		
		return getIntProperty("Index");
	}
	
	/**
	 * Returns a String representing the full path to the linked attached file.
	 * Read-only.
	 * <p>
	 * This property is only valid for linked files.
	 * </p>
	 * 
	 * @return a String representing the full path to the linked attached file.
	 */
	public String getPathName() {
		
		return getStringProperty("PathName");
	}

	/**
	 * Returns an int indicating the position of the attachment within the body
	 * of the item (item: An item is the basic element that holds information in
	 * Outlook (similar to a file in other programs). Items include e-mail
	 * messages, appointments, contacts, tasks, journal entries, notes, posted
	 * items, and documents.). Read/write.
	 * 
	 * @return an int indicating the position of the attachment within the body
	 *         of the item.
	 */
	public int getPosition() {
		
		return getIntProperty("Position");
	}
	
	/**
	 * Sets an int indicating the position of the attachment within the body of
	 * the item (item: An item is the basic element that holds information in
	 * Outlook (similar to a file in other programs). Items include e-mail
	 * messages, appointments, contacts, tasks, journal entries, notes, posted
	 * items, and documents.). Read/write.
	 * 
	 * @param pos
	 *            an int indicating the position of the attachment within the
	 *            body of the item
	 */
	public void setPosition(int pos) {
		
		setProperty("Position", pos);
	}
	
	/**
	 * Returns a PropertyAccessor object that supports creating, getting,
	 * setting, and deleting properties of the parent Attachment object.
	 * Read-only.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a PropertyAccessor object that supports creating, getting,
	 *         setting, and deleting properties of the parent Attachment object.
	 */
	public PropertyAccessor getPropertyAccessor() {
		
		return new PropertyAccessor(getAutomationProperty("PropertyAccessor"));
	}
	
	/**
	 * Saves the attachment to the specified path.
	 * 
	 * @param file
	 *            path and name to save the attachment to.
	 */
	public void saveAsFile(File file) {
		
		invoke("SaveAsFile", newVariant(file.getAbsolutePath()));
	}
	
	/**
	 * Returns an int indicating the size (in bytes) of the attachment.
	 * Read-only.
	 * 
	 * @return an int indicating the size (in bytes) of the attachment.
	 */
	public int getSize() {
		
		return getIntProperty("Size");
	}
	
	/**
	 * Returns an AttachmentType constant indicating the type of the specified
	 * object. Read-only.
	 * 
	 * @return an AttachmentType constant indicating the type of the specified
	 *         object.
	 */
	public AttachmentType getType() {
		
		return AttachmentType.parse(getShortProperty("Type"));
	}

}
