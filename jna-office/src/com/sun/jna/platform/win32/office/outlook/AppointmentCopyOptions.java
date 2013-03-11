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
 * Specifies what actions to take when copying an AppointmentItem object to a
 * folder.
 * <p>
 * Added in Outlook 2010.
 * </p>
 * 
 * @author Ian Darby
 * 
 * @see {@link AppointmentItem}
 */
public class AppointmentCopyOptions extends AbstractEnum {
	
	/**
	 * Copies the appointment to the destination folder and prompts the user to
	 * accept the request before completing the copy operation.
	 */
	public final static AppointmentCopyOptions olPromptUser = new AppointmentCopyOptions(0, "olPromptUser");
	
	/**
	 * Creates an appointment in the destination folder without defaulting to a
	 * response or prompting for a response.
	 */
	public final static AppointmentCopyOptions olCreateAppointment = new AppointmentCopyOptions(1, "olCreateAppointment");
	
	/**
	 * Creates an appointment in the destination folder and accepts the meeting
	 * request automatically.
	 */
	public final static AppointmentCopyOptions olCopyAsAccept = new AppointmentCopyOptions(2, "olCopyAsAccept");

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
	private AppointmentCopyOptions(int val, String name) {
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
	public static AppointmentCopyOptions parse(short val) {
		
		switch(val) {
		
		case 0:
			return olPromptUser;
			
		case 1:
			return olCreateAppointment;
			
		case 2:
			return olCopyAsAccept;
			
		default:
			throw new RuntimeException("AppointmentCopyOption Enum: " + val + " not recognised.");
		}
	}
}
