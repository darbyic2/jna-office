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
 * Contains a collection of addresses for an AddressList object.
 * 
 * @author Ian Darby
 *
 */
public class AddressEntries extends BaseOutlookObject {

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
	AddressEntries(IDispatch iDisp) {
		super(iDisp);
	}
	
	/**
	 * Adds a new entry to the AddressEntries collection.
	 * <p>
	 * New entries or changes to existing entries are not persisted in the
	 * collection until after calling the Update method.
	 * </p>
	 * 
	 * @param addressType
	 *            The type of the new entry.
	 * 
	 * @return An AddressEntry object that represents the new entry.
	 */
	public AddressEntry add(String addressType) {
		
		return new AddressEntry((IDispatch) invoke("Add", newVariant(addressType)).getValue());
	}
	
	/**
	 * Adds a new entry to the AddressEntries collection.
	 * <p>
	 * New entries or changes to existing entries are not persisted in the
	 * collection until after calling the Update method.
	 * </p>
	 * 
	 * @param addressType
	 *            The type of the new entry.
	 * 
	 * @param entryName
	 *            The name of the new entry.
	 * 
	 * @return An AddressEntry object that represents the new entry.
	 */
	public AddressEntry add(String addressType, String entryName) {
		
		return new AddressEntry((IDispatch) invoke("Add", newVariant(addressType), newVariant(entryName)).getValue());
	}
	
	/**
	 * Adds a new entry to the AddressEntries collection.
	 * <p>
	 * New entries or changes to existing entries are not persisted in the
	 * collection until after calling the Update method.
	 * </p>
	 * 
	 * @param addressType
	 *            The type of the new entry.
	 * 
	 * @param entryName
	 *            The name of the new entry.
	 * 
	 * @param address
	 *            The address.
	 * 
	 * @return An AddressEntry object that represents the new entry.
	 */
	public AddressEntry add(String addressType, String entryName, String address) {
		
		return new AddressEntry((IDispatch) invoke("Add", newVariant(addressType), newVariant(entryName), newVariant(address)).getValue());
	}
	
	/**
	 * Returns an int indicating the count of objects in the specified
	 * collection. Read-only.
	 * 
	 * @return an int indicating the count of objects in the specified
	 *         collection. Read-only.
	 */
	public int count() {

		return getIntProperty("Count");
	}

	/**
	 * Helper factory method used by getFirst/getLast/getNext/getPrevious.
	 * 
	 * @param methodName
	 *            name of method to obtain IDispatch object from.
	 * 
	 * @return a populated AddressEntry wrapper instance. Or returns null if no
	 *         such requested object exists.
	 */
	private AddressEntry getEntryFromMethod(String methodName) {

		VARIANT result = invoke(methodName);

		if (result == null
				|| result.getVarType().intValue() == Variant.VT_EMPTY
				|| result.getVarType().intValue() == Variant.VT_NULL
				|| result.getVarType().intValue() != Variant.VT_DISPATCH) {

			return null;

		} else {
			return new AddressEntry((IDispatch) result.getValue());
		}
	}

	/**
	 * Returns the first object in the AddressEntries collection.
	 * 
	 * @return the first object in the AddressEntries collection. Or returns
	 *         null if no such requested object exists.
	 */
	public AddressEntry getFirst() {
		
		return getEntryFromMethod("GetFirst");
	}

	/**
	 * Returns the last object in the AddressEntries collection.
	 * 
	 * @return the last object in the AddressEntries collection. Or returns null
	 *         if no such requested object exists.
	 */
	public AddressEntry getLast() {
		
		return getEntryFromMethod("GetLast");
	}

	/**
	 * Returns the next object in the AddressEntries collection. Prior to the
	 * first call to getNext(), getFirst() must have been called.
	 * 
	 * @return the next object in the AddressEntries collection. Or returns null
	 *         if no such requested object exists.
	 */
	public AddressEntry getNext() {
		
		return getEntryFromMethod("GetNext");
	}

	/**
	 * Returns the previous object in the AddressEntries collection. Prior to
	 * the first call to getPrevious(), getLast() must have been called.
	 * 
	 * @return the previous object in the AddressEntries collection. Or returns
	 *         null if no such requested object exists.
	 */
	public AddressEntry getPrevious() {
		
		return getEntryFromMethod("GetPrevious");
	}

	/**
	 * Returns an AddressEntry object from the collection.
	 * 
	 * @param index
	 *            1's based index number of the object in the collection.
	 * 
	 * @return an AddressEntry object from the collection.
	 */
	public AddressEntry getItem(int index) {
		
		return new AddressEntry((IDispatch) invoke("Item", newVariant(index)).getValue());
	}
	
	/**
	 * Returns an AddressEntry object from the collection.
	 * 
	 * @param entryName
	 *            name of the AddressEntry being sought.
	 * 
	 * @return an AddressEntry object from the collection.
	 */
	public AddressEntry getItem(String entryName) {
		
		return new AddressEntry((IDispatch) invoke("Item", newVariant(entryName)).getValue());
	}
	
	/**
	 * Sorts the collection of items by the specified property. The index for
	 * the collection is reset to 1 upon completion of this method.
	 * <p>
	 * Sort only affects the order of items in a collection. It does not affect
	 * the order of items in an explorer view.
	 * </p>
	 * 
	 * @param propertyNameToSortOn
	 *            The name of the property by which to sort, which may be
	 *            enclosed in brackets, for example, "[CompanyName]". May not be
	 *            a user-defined field, and may not be a multi-valued property,
	 *            such as a category.
	 * 
	 * @param order
	 *            The order for the specified address entries. Can be one of
	 *            these SortOrder constants: olAscending, olDescending, or
	 *            olSortNone.
	 */
	public void sort(String propertyNameToSortOn, SortOrder order) {
		
		String prop;
		
		if (propertyNameToSortOn.startsWith("[")) {
			prop = propertyNameToSortOn;
			
		} else {
			prop = "[" + propertyNameToSortOn + "]";
		}
		invokeNoReply("Sort", newVariant(prop), newVariant(order.value()));
	}
}
