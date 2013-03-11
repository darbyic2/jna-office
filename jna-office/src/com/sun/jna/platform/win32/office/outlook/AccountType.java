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
 * Specifies the type of an Account. Type-safe representation of the Outlook
 * model olAccountType enum.
 * <p>
 * Added in Outlook 2007.
 * </p>
 * 
 * @author Ian Darby
 * 
 * @see {@link Account}
 * @see {@link AbstractEnum}
 */
public class AccountType extends AbstractEnum {
	
	/**
	 * An Exchange account.
	 */
	public final static AccountType	olExchange		= new AccountType(0, "olExchange");
	
	/**
	 * An IMAP account.
	 */
	public final static AccountType	olImap			= new AccountType(1, "olImap");
	
	/**
	 * A POP3 account.
	 */
	public final static AccountType	olPop3			= new AccountType(2, "olPop3");
	
	/**
	 * An HTTP account.
	 */
	public final static AccountType	olHttp			= new AccountType(3, "olHttp");
	
	/**
	 * Other or unknown account.
	 */
	public final static AccountType	olOtherAccount	= new AccountType(5, "olOtherAccount");

	/**
	 * One and only constructor. Scope is private to prevent the creation of
	 * anything other than the built-in constant instances.
	 * 
	 * @param typ
	 *            numeric value used to represent the enum in external storage.
	 * 
	 * @param name
	 *            constant name given to the enum.
	 * 
	 * @see {@link Account}
	 */
	private AccountType(int typ, String name) {
		super((short) typ, name);
	}

	/**
	 * Converts an external storage numeric representation in to one of the
	 * built-in constant objects. Unrecognised external values cause a type of
	 * olOtherAccount to be returned.
	 * 
	 * @param typ
	 *            external numeric representation.
	 * 
	 * @return one of the built-in constant objects that represents the enum in
	 *         a type-safe way.
	 */
	public static AccountType parse(short typ) {
		
		switch(typ) {
		
		case 0:
			return olExchange;
			
		case 1:
			return olImap;
			
		case 2:
			return olPop3;
			
		case 3:
			return olHttp;
			
		case 5:
		default:
			return olOtherAccount;
		}
	}
}
