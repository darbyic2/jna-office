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

/**
 * Identifies the type of Electronic Business Card (EBC) format associated with
 * a ContactItem object.
 * <p>
 * Added in Outlook 2007.
 * </p>
 * 
 * @author Ian Darby
 * 
 * @see AbstractEnum
 */
public class BusinessCardType extends AbstractEnum {
	
	/**
	 * Indicates that the ContactItem object uses the Microsoft Outlook format
	 * for the associated Electronic Business Card.
	 */
	public final static BusinessCardType olBusinessCardTypeOutlook = new BusinessCardType(0, "olBusinessCardTypeOutlook");
	
	/**
	 * Indicates that the ContactItem uses the Microsoft Office InterConnect
	 * format for the associated Electronic Business Card.
	 */
	public final static BusinessCardType olBusinessCardTypeInterConnect = new BusinessCardType(1, "olBusinessCardTypeInterConnect");

	/**
	 * One and only constructor. Scope is private to prevent the creation of
	 * anything other than the built-in constant instances.
	 * 
	 * @param typ
	 *            numeric value used to represent the enum in external storage.
	 * 
	 * @param name
	 *            constant name given to the enum.
	 */
	private BusinessCardType(int typ, String name) {
		super((short) typ, name);
	}

	/**
	 * Converts an external storage numeric representation in to one of the
	 * built-in constant objects. Unrecognised external values cause a
	 * {@link RuntimeException} to be thrown.
	 * 
	 * @param typ
	 *            external numeric representation.
	 * 
	 * @return one of the built-in constant objects that represents the enum in
	 *         a type-safe way.
	 */
	public static BusinessCardType parse(short typ) {
		
		switch(typ) {
		
		case 0:
			return olBusinessCardTypeOutlook;
			
		case 1:
			return olBusinessCardTypeInterConnect;
		
		default:
			throw new RuntimeException("BusinessCardType Enum: " + typ + " not recognised.");
		}
	}
}
