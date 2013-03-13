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

import com.sun.jna.platform.win32.Variant;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant.VARIANT;

/**
 * Contains a collection of Conflict objects that represent all Microsoft
 * Outlook items that are in conflict with a particular Outlook item (item: An
 * item is the basic element that holds information in Outlook (similar to a
 * file in other programs). Items include e-mail messages, appointments,
 * contacts, tasks, journal entries, notes, posted items, and documents.).
 * <p>
 * Use the Conflicts property of any Outlook item, such as MailItem, to return
 * the Conflicts object.
 * </p>
 * <p>
 * Use the Count property of the Conflicts object to determine if the item is
 * invloved in a conflict. A non-zero value indicates conflict.
 * </p>
 * <p>
 * Use the Item method to retrieve a particular conflict item from the Conflicts
 * collection object.
 * </p>
 * <p>
 * Use the GetFirst, GetNext, GetPrevious, and GetLast methods to traverse the
 * Conflicts collection.
 * </p>
 * 
 * @author Ian Darby
 * 
 */
public class Conflicts extends BaseOutlookObject {

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
	Conflicts(IDispatch iDisp) {
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
	 * Underlying worker method to handle null and unexpected return values from
	 * the iteration methods.
	 * 
	 * @param methodName
	 *            name of iteration method to call. Legal values are:
	 *            &quot;GetFirst&quot;, &quot;GetNext&quot;,
	 *            &quot;GetLast&quot;, &quot;GetPrevious&quot;.
	 * 
	 * @return a new fully populated Conflict wrapper object.
	 */
	private Conflict getConflictFromMethod(String methodName) {

		VARIANT result = invoke(methodName);

		if (result == null || result.getValue() == null
				|| result.getVarType().intValue() != Variant.VT_DISPATCH) {

			return null;

		} else {
			return new Conflict((IDispatch) result.getValue());
		}
	}

	/**
	 * Returns the first object in the Conflicts collection.
	 * <p>
	 * Returns null if no first object exists, for example, if there are no
	 * objects in the collection. To ensure correct operation of the GetFirst,
	 * GetLast, GetNext, and GetPrevious methods in a large collection, call
	 * GetFirst before calling GetNext on that collection and call GetLast
	 * before calling GetPrevious. To ensure that you are always making the
	 * calls on the same collection, create an explicit variable that refers to
	 * that collection before entering the loop.
	 * </p>
	 * 
	 * @return the first object in the Conflicts collection or null if the
	 *         collection is empty.
	 */
	public Conflict getFirst() {
		
		return getConflictFromMethod("GetFirst");
	}

	/**
	 * Returns the last object in the Conflicts collection.
	 * <p>
	 * Returns null if no last object exists, for example, if there are no
	 * objects in the collection. To ensure correct operation of the GetFirst,
	 * GetLast, GetNext, and GetPrevious methods in a large collection, call
	 * GetFirst before calling GetNext on that collection and call GetLast
	 * before calling GetPrevious. To ensure that you are always making the
	 * calls on the same collection, create an explicit variable that refers to
	 * that collection before entering the loop.
	 * </p>
	 * 
	 * @return the last object in the Conflicts collection or null if the
	 *         collection is empty.
	 */
	public Conflict getLast() {
		
		return getConflictFromMethod("GetLast");
	}

	/**
	 * Returns the next object in the Conflicts collection.
	 * <p>
	 * Returns null if no next object exists, for example, if already positioned
	 * at the end of the collection. To ensure correct operation of the
	 * GetFirst, GetLast, GetNext, and GetPrevious methods in a large
	 * collection, call GetFirst before calling GetNext on that collection and
	 * call GetLast before calling GetPrevious. To ensure that you are always
	 * making the calls on the same collection, create an explicit variable that
	 * refers to that collection before entering the loop.
	 * </p>
	 * 
	 * @return the last object in the Conflicts collection or null if there is
	 *         no next conflict.
	 */
	public Conflict getNext() {
		
		return getConflictFromMethod("GetNext");
	}

	/**
	 * Returns the previous object in the Conflicts collection.
	 * <p>
	 * Returns null if no previous object exists, for example, if already
	 * positioned at the start of the collection. To ensure correct operation of
	 * the GetFirst, GetLast, GetNext, and GetPrevious methods in a large
	 * collection, call GetFirst before calling GetNext on that collection and
	 * call GetLast before calling GetPrevious. To ensure that you are always
	 * making the calls on the same collection, create an explicit variable that
	 * refers to that collection before entering the loop.
	 * </p>
	 * 
	 * @return the previous object in the Conflicts collection or null if there
	 *         is no previous conflict.
	 */
	public Conflict getPrevious() {
		
		return getConflictFromMethod("GetPrevious");
	}

	/**
	 * Returns a Conflict object from the collection.
	 * 
	 * @param index
	 *            the index number of the object within the collection.
	 * 
	 * @return the specified Conflict object from the collection.
	 */
	public Conflict getItem(int index) {
		
		return new Conflict((IDispatch) invoke("Item", newVariant(index)).getValue());
	}
	
	/**
	 * Returns a Conflict object from the collection.
	 * 
	 * @param conflictName
	 *            the display name of the required Conflict item.
	 * 
	 * @return the specified Conflict object from the collection.
	 */
	public Conflict getItem(String conflictName) {
		
		return new Conflict((IDispatch) invoke("Item", newVariant(conflictName)).getValue());
	}
	
}
