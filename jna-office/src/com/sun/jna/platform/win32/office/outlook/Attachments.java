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
 * Contains a set of Attachment objects that represent the attachments in an
 * Outlook item (item: An item is the basic element that holds information in
 * Outlook (similar to a file in other programs). Items include e-mail messages,
 * appointments, contacts, tasks, journal entries, notes, posted items, and
 * documents.).
 * <p>
 * Use the Attachments property to return the Attachments collection for any
 * Outlook item (except notes).
 * </p>
 * <p>
 * Use the add method to add an attachment to an item.
 * </p>
 * <p>
 * To ensure consistent results, always save an item before adding or removing
 * objects in the Attachments collection of the item.
 * </p>
 * 
 * @author Ian Darby
 * 
 */
public class Attachments extends BaseOutlookObject {

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
	Attachments(IDispatch iDisp) {
		super(iDisp);
	}
	
	/**
	 * Creates a new attachment in the Attachments collection.
	 * <p>
	 * When an Attachment is added to the Attachments collection of an item, the
	 * Type property of the Attachment will always return olOLE (6) until the
	 * item is saved. To ensure consistent results, always save an item before
	 * adding or removing objects in the Attachments collection.
	 * </p>
	 * 
	 * @param src
	 *            The source of the attachment. This can be a file (represented
	 *            by the full file system path with a file name) or an Outlook
	 *            item that constitutes the attachment.
	 * 
	 * @return An Attachment object that represents the new attachment.
	 */
	public Attachment add(File src) {
		
		return new Attachment((IDispatch) invoke("Add", newVariant(src.getAbsolutePath())).getValue());
	}
	
	/**
	 * Creates a new attachment in the Attachments collection.
	 * <p>
	 * When an Attachment is added to the Attachments collection of an item, the
	 * Type property of the Attachment will always return olOLE (6) until the
	 * item is saved. To ensure consistent results, always save an item before
	 * adding or removing objects in the Attachments collection.
	 * </p>
	 * 
	 * @param src
	 *            The source of the attachment. This can be a file (represented
	 *            by the full file system path with a file name) or an Outlook
	 *            item that constitutes the attachment.
	 * 
	 * @return An Attachment object that represents the new attachment.
	 */
	public Attachment add(BaseItemLevel1 src) {
		
		return new Attachment((IDispatch) invoke("Add", newVariant(src.getIDispatch())).getValue());
	}
	
	/**
	 * Creates a new attachment in the Attachments collection.
	 * <p>
	 * When an Attachment is added to the Attachments collection of an item, the
	 * Type property of the Attachment will always return olOLE (6) until the
	 * item is saved. To ensure consistent results, always save an item before
	 * adding or removing objects in the Attachments collection.
	 * </p>
	 * 
	 * @param src
	 *            The source of the attachment. This can be a file (represented
	 *            by the full file system path with a file name) or an Outlook
	 *            item that constitutes the attachment.
	 * 
	 * @param typ
	 *            The type of the attachment. Can be one of the AttachmentType
	 *            constants.
	 * 
	 * @return An Attachment object that represents the new attachment.
	 */
	public Attachment add(File src, AttachmentType typ) {
		
		return new Attachment((IDispatch) invoke("Add", newVariant(src.getAbsolutePath()), newVariant(typ.value())).getValue());
	}
	
	/**
	 * Creates a new attachment in the Attachments collection.
	 * <p>
	 * When an Attachment is added to the Attachments collection of an item, the
	 * Type property of the Attachment will always return olOLE (6) until the
	 * item is saved. To ensure consistent results, always save an item before
	 * adding or removing objects in the Attachments collection.
	 * </p>
	 * 
	 * @param src
	 *            The source of the attachment. This can be a file (represented
	 *            by the full file system path with a file name) or an Outlook
	 *            item that constitutes the attachment.
	 * 
	 * @param typ
	 *            The type of the attachment. Can be one of the AttachmentType
	 *            constants.
	 * 
	 * @return An Attachment object that represents the new attachment.
	 */
	public Attachment add(BaseItemLevel1 src, AttachmentType typ) {
		
		return new Attachment((IDispatch) invoke("Add", newVariant(src.getIDispatch()), newVariant(typ.value())).getValue());
	}
	
	/**
	 * Creates a new attachment in the Attachments collection.
	 * <p>
	 * When an Attachment is added to the Attachments collection of an item, the
	 * Type property of the Attachment will always return olOLE (6) until the
	 * item is saved. To ensure consistent results, always save an item before
	 * adding or removing objects in the Attachments collection.
	 * </p>
	 * 
	 * @param src
	 *            The source of the attachment. This can be a file (represented
	 *            by the full file system path with a file name) or an Outlook
	 *            item that constitutes the attachment.
	 * 
	 * @param typ
	 *            The type of the attachment. Can be one of the AttachmentType
	 *            constants.
	 * 
	 * @param position
	 *            This parameter applies only to e-mail messages using the Rich
	 *            Text format: it is the position where the attachment should be
	 *            placed within the body text of the message. A value of 1 for
	 *            the Position parameter specifies that the attachment should be
	 *            positioned at the beginning of the message body. A value 'n'
	 *            greater than the number of characters in the body of the
	 *            e-mail item specifies that the attachment should be placed at
	 *            the end. A value of 0 makes the attachment hidden.
	 * 
	 * @return An Attachment object that represents the new attachment.
	 */
	public Attachment add(File src, AttachmentType typ, int position) {
		
		return new Attachment((IDispatch) invoke("Add", newVariant(src.getAbsolutePath()), newVariant(typ.value()), newVariant(position)).getValue());
	}
	
	/**
	 * Creates a new attachment in the Attachments collection.
	 * <p>
	 * When an Attachment is added to the Attachments collection of an item, the
	 * Type property of the Attachment will always return olOLE (6) until the
	 * item is saved. To ensure consistent results, always save an item before
	 * adding or removing objects in the Attachments collection.
	 * </p>
	 * 
	 * @param src
	 *            The source of the attachment. This can be a file (represented
	 *            by the full file system path with a file name) or an Outlook
	 *            item that constitutes the attachment.
	 * 
	 * @param typ
	 *            The type of the attachment. Can be one of the AttachmentType
	 *            constants.
	 * 
	 * @param position
	 *            This parameter applies only to e-mail messages using the Rich
	 *            Text format: it is the position where the attachment should be
	 *            placed within the body text of the message. A value of 1 for
	 *            the Position parameter specifies that the attachment should be
	 *            positioned at the beginning of the message body. A value 'n'
	 *            greater than the number of characters in the body of the
	 *            e-mail item specifies that the attachment should be placed at
	 *            the end. A value of 0 makes the attachment hidden.
	 * 
	 * @return An Attachment object that represents the new attachment.
	 */
	public Attachment add(BaseItemLevel1 src, AttachmentType typ, int position) {
		
		return new Attachment((IDispatch) invoke("Add", newVariant(src.getIDispatch()), newVariant(typ.value()), newVariant(position)).getValue());
	}
	
	/**
	 * Creates a new attachment in the Attachments collection.
	 * <p>
	 * When an Attachment is added to the Attachments collection of an item, the
	 * Type property of the Attachment will always return olOLE (6) until the
	 * item is saved. To ensure consistent results, always save an item before
	 * adding or removing objects in the Attachments collection.
	 * </p>
	 * 
	 * @param src
	 *            The source of the attachment. This can be a file (represented
	 *            by the full file system path with a file name) or an Outlook
	 *            item that constitutes the attachment.
	 * 
	 * @param typ
	 *            The type of the attachment. Can be one of the AttachmentType
	 *            constants.
	 * 
	 * @param position
	 *            This parameter applies only to e-mail messages using the Rich
	 *            Text format: it is the position where the attachment should be
	 *            placed within the body text of the message. A value of 1 for
	 *            the Position parameter specifies that the attachment should be
	 *            positioned at the beginning of the message body. A value 'n'
	 *            greater than the number of characters in the body of the
	 *            e-mail item specifies that the attachment should be placed at
	 *            the end. A value of 0 makes the attachment hidden.
	 * 
	 * @param displayName
	 *            This parameter applies only if the mail item is in Rich Text
	 *            format and Type is set to olByValue: the name is displayed in
	 *            an Inspector object for the attachment or when viewing the
	 *            properties of the attachment. If the mail item is in Plain
	 *            Text or HTML format, then the attachment is displayed using
	 *            the file name in the Source parameter.
	 * 
	 * @return An Attachment object that represents the new attachment.
	 */
	public Attachment add(File src, AttachmentType typ, int position, String displayName) {
		
		return new Attachment((IDispatch) invoke("Add", newVariant(src.getAbsolutePath()), newVariant(typ.value()), newVariant(position), newVariant(displayName)).getValue());
	}
	
	/**
	 * Creates a new attachment in the Attachments collection.
	 * <p>
	 * When an Attachment is added to the Attachments collection of an item, the
	 * Type property of the Attachment will always return olOLE (6) until the
	 * item is saved. To ensure consistent results, always save an item before
	 * adding or removing objects in the Attachments collection.
	 * </p>
	 * 
	 * @param src
	 *            The source of the attachment. This can be a file (represented
	 *            by the full file system path with a file name) or an Outlook
	 *            item that constitutes the attachment.
	 * 
	 * @param typ
	 *            The type of the attachment. Can be one of the AttachmentType
	 *            constants.
	 * 
	 * @param position
	 *            This parameter applies only to e-mail messages using the Rich
	 *            Text format: it is the position where the attachment should be
	 *            placed within the body text of the message. A value of 1 for
	 *            the Position parameter specifies that the attachment should be
	 *            positioned at the beginning of the message body. A value 'n'
	 *            greater than the number of characters in the body of the
	 *            e-mail item specifies that the attachment should be placed at
	 *            the end. A value of 0 makes the attachment hidden.
	 * 
	 * @param displayName
	 *            This parameter applies only if the mail item is in Rich Text
	 *            format and Type is set to olByValue: the name is displayed in
	 *            an Inspector object for the attachment or when viewing the
	 *            properties of the attachment. If the mail item is in Plain
	 *            Text or HTML format, then the attachment is displayed using
	 *            the file name in the Source parameter.
	 * 
	 * @return An Attachment object that represents the new attachment.
	 */
	public Attachment add(BaseItemLevel1 src, AttachmentType typ, int position, String displayName) {
		
		return new Attachment((IDispatch) invoke("Add", newVariant(src.getIDispatch()), newVariant(typ.value()), newVariant(position), newVariant(displayName)).getValue());
	}
	
	/**
	 * Returns an int indicating the count of objects in the specified
	 * collection. Read-only.
	 * 
	 * @return an int indicating the count of objects in the specified
	 *         collection.
	 */
	public int count() {
		
		return getIntProperty("Count");
	}
	
	/**
	 * Returns an Attachment object from the collection.
	 * 
	 * @param index
	 *            1 based index number of the attachment in the collection.
	 * 
	 * @return an Attachment object from the collection.
	 */
	public Attachment getItem(int index) {
		
		return new Attachment((IDispatch) invoke("Item", newVariant(index)).getValue());
	}
	
	/**
	 * Removes an object from the collection.
	 * 
	 * @param index
	 *            1 based index number of the attachment in the collection.
	 */
	public void remove(int index) {
		
		invokeNoReply("Remove", newVariant(index));
	}
	
}
