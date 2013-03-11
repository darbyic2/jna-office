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
 * Contains a collection of Action objects that represent all the specialised
 * actions that can be executed on an Outlook item (item: An item is the basic
 * element that holds information in Outlook (similar to a file in other
 * programs). Items include e-mail messages, appointments, contacts, tasks,
 * journal entries, notes, posted items, and documents.).
 * 
 * @author Ian Darby
 * 
 */
public class Actions extends BaseOutlookObject {

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
	Actions(IDispatch iDisp) {
		super(iDisp);
	}
	
	/**
	 * Creates a new action in the Actions collection.
	 * 
	 * @return a new action that has been added to the Actions collection.
	 */
	public Action add() {
		
		return new Action((IDispatch) invoke("Add").getValue());
	}
	
	/**
	 * @return an int indicating the count of objects in the specified
	 *         collection. Read-only.
	 */
	public int count() {
		
		return getIntProperty("Count");
	}
	
	/**
	 * Returns an Action object from the collection.
	 * 
	 * @param index
	 *            The 1-based index value of the object within the collection.
	 * 
	 * @return an Action object from the collection.
	 */
	public Action getItem(int index) {
		
		return new Action((IDispatch) invoke("Item", newVariant(index)).getValue());
	}
	
	/**
	 * Returns an Action object from the collection.
	 * 
	 * @param actionName
	 *            name of the action to be retrieved from the collection.
	 * 
	 * @return an Action object from the collection.
	 */
	public Action getItem(String actionName) {
		
		return new Action((IDispatch) invoke("Item", newVariant(actionName)).getValue());
	}
	
	/**
	 * Removes an object from the collection.
	 * 
	 * @param index
	 *            The 1-based index value of the object within the collection.
	 */
	public void remove(int index) {
		
		invokeNoReply("Display", newVariant(index));
	}
	
}
