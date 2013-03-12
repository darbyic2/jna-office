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
 * Contains a set of Attachment objects that represent the selected attachments
 * in an Outlook item (item: An item is the basic element that holds information
 * in Outlook (similar to a file in other programs). Items include e-mail
 * messages, appointments, contacts, tasks, journal entries, notes, posted
 * items, and documents.).
 * <p>
 * The AttachmentSelection object contains a read-only collection of attachments
 * that are selected in an item that is in the active inspector or the active
 * explorer.
 * </p>
 * <p>
 * Added in Outlook 2007.
 * </p>
 * 
 * @author Ian Darby
 * 
 */
public class AttachmentSelection extends BaseOutlookObject {

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
	AttachmentSelection(IDispatch iDisp) {
		super(iDisp);
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
	 *            the index number of the object within the collection.
	 * 
	 * @return an Attachment object from the collection.
	 */
	public Attachment getItem(int index) {
		
		return new Attachment((IDispatch) invoke("Item", newVariant(index)).getValue());
	}
	
	/**
	 * Returns a SelectionLocation constant that specifies where the attachment
	 * selection is in the Microsoft Outlook user interface. Read-only.
	 * <p>
	 * This property always returns the constant olAttachmentWell.
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @return a SelectionLocation constant that specifies where the attachment
	 *         selection is in the Microsoft Outlook user interface.
	 */
	public SelectionLocation getLocation() {
		
		return SelectionLocation.parse(getShortProperty("Location"));
	}
}
