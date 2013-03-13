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
 * Represents an Outlook item that is in conflict with another Outlook item.
 * <p>
 * Each Outlook item has a Conflicts collection object associated with it that
 * represents all the items that are in conflict with that item.
 * </p>
 * <p>
 * Use the Item method to retrieve a particular Conflict object from the
 * Conflicts collection object, for example:
 * </p>
 * 
 * @author Ian Darby
 * 
 * @see BaseOutlookObject
 */
public class Conflict extends BaseOutlookObject {

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
	Conflict(IDispatch iDisp) {
		super(iDisp);
	}
	
	/**
	 * Returns an Object corresponding to the specified Outlook item. Read-only.
	 * <p>
	 * This implementation needs re-thinking. This return class provides very
	 * little in the way of useful methods and can not be cast to the correct
	 * type.
	 * </p>
	 * 
	 * @return an Object corresponding to the specified Outlook item.
	 */
	public BaseOutlookObject getItem() {
		
		return new BaseOutlookObject(getAutomationProperty("Item"));
	}
	
	/**
	 * Returns the display name for the object. Read-only.
	 * 
	 * @return the display name for the object.
	 */
	public String getName() {
		
		return getStringProperty("Name");
	}
	
	/**
	 * Returns a ClassEnum integer constant indicating the type of item
	 * represented by the Conflict object. Read-only.
	 * <p>
	 * The return value is identical to the coresponding value in the
	 * OlObjectClass enumeration in the Outlook model.
	 * </p>
	 * 
	 * @return a ClassEnum integer constant indicating the type of item
	 *         represented by the Conflict object.
	 */
	public int getType() {
		
		return getIntProperty("Type");
	}
}
