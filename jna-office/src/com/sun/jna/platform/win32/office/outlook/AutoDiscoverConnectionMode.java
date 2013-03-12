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
 * Specifies the type of connection to the Exchange server for the
 * auto-discovery service.
 * <p>
 * Added in Outlook 2007.
 * </p>
 * 
 * @author Ian Darby
 * 
 * @see AbstractEnum
 */
public class AutoDiscoverConnectionMode extends AbstractEnum {
	
	/**
	 * Other or unknown connection, or no connection.
	 */
	public final static AutoDiscoverConnectionMode olAutoDiscoverConnectionUnknown		  = new AutoDiscoverConnectionMode(0, "olAutoDiscoverConnectionUnknown"); //	Other or unknown connection, or no connection.
	
	/**
	 * Connection is over the Internet.
	 */
	public final static AutoDiscoverConnectionMode olAutoDiscoverConnectionExternal		  = new AutoDiscoverConnectionMode(1, "olAutoDiscoverConnectionExternal"); //	Connection is over the Internet.
	
	/**
	 * Connection is over the Intranet.
	 */
	public final static AutoDiscoverConnectionMode olAutoDiscoverConnectionInternal		  = new AutoDiscoverConnectionMode(2, "olAutoDiscoverConnectionInternal"); //	Connection is over the Intranet.
	
	/**
	 * Connection is in the same domain over the Intranet.
	 */
	public final static AutoDiscoverConnectionMode olAutoDiscoverConnectionInternalDomain = new AutoDiscoverConnectionMode(3, "olAutoDiscoverConnectionInternalDomain"); //	Connection is in the same domain over the Intranet.

	/**
	 * One and only constructor. Scope is private to prevent the creation of
	 * anything other than the built-in constant instances.
	 * 
	 * @param mode
	 *            numeric value used to represent the enum in external storage.
	 * 
	 * @param name
	 *            constant name given to the enum.
	 */
	private AutoDiscoverConnectionMode(int mode, String name) {
		super((short) mode, name);
	}

	/**
	 * Converts an external storage numeric representation in to one of the
	 * built-in constant objects. Unrecognised external values cause a type of
	 * olAutoDiscoverConnectionUnknown to be returned.
	 * 
	 * @param mode
	 *            external numeric representation.
	 * 
	 * @return one of the built-in constant objects that represents the enum in
	 *         a type-safe way.
	 */
	public static AutoDiscoverConnectionMode parse(short mode) {
		
		switch(mode) {
		
		case 0:
			return olAutoDiscoverConnectionUnknown;
			
		case 1:
			return olAutoDiscoverConnectionExternal;
			
		case 2:
			return olAutoDiscoverConnectionInternal;
			
		case 3:
			return olAutoDiscoverConnectionInternalDomain;
			
		default:
			return olAutoDiscoverConnectionUnknown;
		}
	}
}
