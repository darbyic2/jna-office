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
 * Type-safe mechanism for representing Enum values that is compatible with Java
 * versions earlier than Java 5. this is the base class from which all concrete
 * Enum classes should inherit.
 * 
 * @author Ian Darby
 * 
 */
public abstract class AbstractEnum {
	
	private short val;
	private String name;

	/**
	 * The one and only constructor. Sub-classes should over-ride this
	 * constructor and make it private.
	 * 
	 * @param val
	 *            short integer value to represent the enum.
	 * 
	 * @param name
	 *            name that has been given to the enum. This is the name that
	 *            will be shown by the {@link #toString()} method.
	 */
	protected AbstractEnum(short val, String name) {
		super();
		this.val = val;
		this.name = name;
	}
	
	/**
	 * @return the short integer value that represents the enum.
	 */
	public short value() {
		
		return val;
	}

	/**
	 * @return a string useful during debugging. Format of returned string is
	 *         &lt;enumName&gt;(&lt;enumValue&gt;).
	 */
	@Override
	public String toString() {
		
		return name + "(" + val + ")";
	}

}
