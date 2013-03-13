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
 * Specifies the telephone number type.
 * <p>
 * Added in Outlook 2007.
 * </p>
 * 
 * @author Ian Darby
 * 
 * @see AbstractEnum
 */
public class ContactPhoneNumber extends AbstractEnum {
	
	/**
	 * Telephone number of the person who is the assistant for the contact
	 */
	public final static ContactPhoneNumber	olContactPhoneAssistant		= new ContactPhoneNumber(0,  "olContactPhoneAssistant"); 
	
	/**
	 * Business telephone number
	 */
	public final static ContactPhoneNumber	olContactPhoneBusiness		= new ContactPhoneNumber(1,  "olContactPhoneBusiness"); 
	
	/**
	 * Second business telephone number
	 */
	public final static ContactPhoneNumber	olContactPhoneBusiness2		= new ContactPhoneNumber(2,  "olContactPhoneBusiness2"); 
	
	/**
	 * Business fax number
	 */
	public final static ContactPhoneNumber	olContactPhoneBusinessFax	= new ContactPhoneNumber(3,  "olContactPhoneBusinessFax"); 
	
	/**
	 * Callback telephone number
	 */
	public final static ContactPhoneNumber	olContactPhoneCallback		= new ContactPhoneNumber(4,  "olContactPhoneCallback"); 
	
	/**
	 * Car telephone number
	 */
	public final static ContactPhoneNumber	olContactPhoneCar			= new ContactPhoneNumber(5,  "olContactPhoneCar"); 
	
	/**
	 * Main company telephone number
	 */
	public final static ContactPhoneNumber	olContactPhoneCompany		= new ContactPhoneNumber(6,  "olContactPhoneCompany"); 
	
	/**
	 * Home telephone number
	 */
	public final static ContactPhoneNumber	olContactPhoneHome			= new ContactPhoneNumber(7,  "olContactPhoneHome"); 
	
	/**
	 * Second home telephone number
	 */
	public final static ContactPhoneNumber	olContactPhoneHome2			= new ContactPhoneNumber(8,  "olContactPhoneHome2"); 
	
	/**
	 * Home fax number
	 */
	public final static ContactPhoneNumber	olContactPhoneHomeFax		= new ContactPhoneNumber(9,  "olContactPhoneHomeFax"); 
	
	/**
	 * Integrated Services Digital Network (ISDN) phone number
	 */
	public final static ContactPhoneNumber	olContactPhoneISDN			= new ContactPhoneNumber(10, "olContactPhoneISDN"); 
	
	/**
	 * Mobile telephone number
	 */
	public final static ContactPhoneNumber	olContactPhoneMobile		= new ContactPhoneNumber(11, "olContactPhoneMobile"); 
	
	/**
	 * Other telephone number
	 */
	public final static ContactPhoneNumber	olContactPhoneOther			= new ContactPhoneNumber(12, "olContactPhoneOther"); 
	
	/**
	 * Other fax number
	 */
	public final static ContactPhoneNumber	olContactPhoneOtherFax		= new ContactPhoneNumber(13, "olContactPhoneOtherFax"); 
	
	/**
	 * Pager telephone number
	 */
	public final static ContactPhoneNumber	olContactPhonePager			= new ContactPhoneNumber(14, "olContactPhonePager"); 
	
	/**
	 * Primary telephone number
	 */
	public final static ContactPhoneNumber	olContactPhonePrimary		= new ContactPhoneNumber(15, "olContactPhonePrimary"); 
	
	/**
	 * Radio telephone number
	 */
	public final static ContactPhoneNumber	olContactPhoneRadio			= new ContactPhoneNumber(16, "olContactPhoneRadio"); 
	
	/**
	 * Telex telephone number
	 */
	public final static ContactPhoneNumber	olContactPhoneTelex			= new ContactPhoneNumber(17, "olContactPhoneTelex"); 
	
	/**
	 * TTD/TTY (Teletypewriting Device for the Deaf/Teletypewriter) telephone
	 * number
	 */
	public final static ContactPhoneNumber	olContactPhoneTTYTTD		= new ContactPhoneNumber(18, "olContactPhoneTTYTTD"); 
	
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
	private ContactPhoneNumber(int typ, String name) {
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
	public static ContactPhoneNumber parse(short typ) {
		
		switch(typ) {
		
		case 0:
			return olContactPhoneAssistant;
			
		case 1:
			return olContactPhoneBusiness;
			
		case 2:
			return olContactPhoneBusiness2;
			
		case 3:
			return olContactPhoneBusinessFax;
			
		case 4:
			return olContactPhoneCallback;
			
		case 5:
			return olContactPhoneCar;
			
		case 6:
			return olContactPhoneCompany;
			
		case 7:
			return olContactPhoneHome;
			
		case 8:
			return olContactPhoneHome2;
			
		case 9:
			return olContactPhoneHomeFax;
			
		case 10:
			return olContactPhoneISDN;
			
		case 11:
			return olContactPhoneMobile;
			
		case 12:
			return olContactPhoneOther;
			
		case 13:
			return olContactPhoneOtherFax;
			
		case 14:
			return olContactPhonePager;
			
		case 15:
			return olContactPhonePrimary;
			
		case 16:
			return olContactPhoneRadio;
			
		case 17:
			return olContactPhoneTelex;
			
		case 18:
			return olContactPhoneTTYTTD;
			
		default:
			throw new RuntimeException("ContactPhoneNumberType Enum: " + typ + " not recognised.");
		}
	}
}
