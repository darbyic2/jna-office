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
 * Specifies the attachment type.
 * 
 * @author Ian Darby
 *
 * @see AbstractEnum
 */
public class AttachmentType extends AbstractEnum {

	/**
	 * The attachment is a copy of the original file and can be accessed even if
	 * the original file is removed.
	 */
	public final static AttachmentType olByValue = new AttachmentType(1, "olByValue");
	
	/**
	 * The attachment is a shortcut to the location of the original file.
	 */
	public final static AttachmentType olByReference = new AttachmentType(4, "olByReference");
	
	/**
	 * The attachment is an Outlook message format file (.msg) and is a copy of
	 * the original message.
	 */
	public final static AttachmentType olEmbeddeditem = new AttachmentType(5, "olEmbeddeditem");
	
	/**
	 * The attachment is an OLE document.
	 */
	public final static AttachmentType olOLE = new AttachmentType(6, "olOLE");
	
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
	 * @see Attachment
	 */
	private AttachmentType(int typ, String name) {
		super((short) typ, name);
	}
	
	/**
	 * Converts an external storage numeric representation in to one of the
	 * built-in constant objects. Unrecognised external values cause a type of
	 * olOLE to be returned.
	 * 
	 * @param typeValue
	 *            external numeric representation.
	 * 
	 * @return one of the built-in constant objects that represents the enum in
	 *         a type-safe way.
	 */
	public static AttachmentType parse(short typeValue) {
		
		switch(typeValue) {
		case 1:
			return olByValue;
			
		case 4:
			return olByReference;
			
		case 5:
			return olEmbeddeditem;
			
		default:
			return olOLE;
		}
	}
}
