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
 * Specifies the response style. Type-safe representation of the Outlook model
 * olActionReplyStyle enum.
 * 
 * @author Ian Darby
 * 
 * @see {@link Action}
 * @see {@link AbstractEnum}
 */
public class ActionResponseStyle extends AbstractEnum {
	
	/**
	 * Indicates that a form will be opened.
	 */
	public final static ActionResponseStyle	olOpen		= new ActionResponseStyle(0	, "olOpen");  
	
	/**
	 * Indicates that the form will be sent immediately.
	 */
	public final static ActionResponseStyle	olSend		= new ActionResponseStyle(1	, "olSend");  
	
	/**
	 * Indicates that the user will be prompted to open or send the form.
	 */
	public final static ActionResponseStyle	olPrompt	= new ActionResponseStyle(2	, "olPrompt");  

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
	 * @see {@link Action}
	 */
	public ActionResponseStyle(int typ, String name) {
		super((short) typ, name);
	}

	/**
	 * Converts an external storage numeric representation in to one of the
	 * built-in constant objects. Unrecognised external values return an
	 * olPrompt constant.
	 * 
	 * @param style
	 *            external numeric representation.
	 * 
	 * @return one of the built-in constant objects that represents the enum in
	 *         a type-safe way.
	 */
	public static ActionResponseStyle parse(short style) {
		switch(style) {
		
		case 0:
			return olOpen;
			
		case 1:
			return olSend;
		
		case 2:
			return olPrompt;
			
		default:
			return olPrompt;
		}
	}
}
