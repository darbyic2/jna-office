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
 * Indicates a user's availability.
 * 
 * @author Ian Darby
 * 
 * @see AbstractEnum
 */
public class BusyStatus extends AbstractEnum {
	
	/**
	 * The user is available.
	 */
	public final static BusyStatus	olFree			= new BusyStatus(0, "olFree"); //The user is available.
	
	/**
	 * The user has a tentative appointment scheduled.
	 */
	public final static BusyStatus	olTentative		= new BusyStatus(1, "olTentative"); //The user has a tentative appointment scheduled.
	
	/**
	 * The user is busy.
	 */
	public final static BusyStatus	olBusy			= new BusyStatus(2, "olBusy"); //The user is busy.
	
	/**
	 * The user is out of office.
	 */
	public final static BusyStatus	olOutOfOffice	= new BusyStatus(3, "olOutOfOffice"); //The user is out of office.

	/**
	 * One and only constructor. Scope is private to prevent the creation of
	 * anything other than the built-in constant instances.
	 * 
	 * @param val
	 *            numeric value used to represent the enum in external storage.
	 * 
	 * @param name
	 *            constant name given to the enum.
	 */
	private BusyStatus(int val, String name) {
		super((short) val, name);
	}

	/**
	 * Converts an external storage numeric representation in to one of the
	 * built-in constant objects. Unrecognised external values cause a
	 * {@link RuntimeException} to be thrown.
	 * 
	 * @param val
	 *            external numeric representation.
	 * 
	 * @return one of the built-in constant objects that represents the enum in
	 *         a type-safe way.
	 */
	public static BusyStatus parse(short val) {
		
		switch(val) {
		
		case 0:
			return olFree;
			
		case 1:
			return olTentative;
			
		case 2:
			return olBusy;
			
		case 3:
			return olOutOfOffice;
			
		default:
			throw new RuntimeException("BusyStatus Enum: " + val + " not recognised.");
		}
	}
}
