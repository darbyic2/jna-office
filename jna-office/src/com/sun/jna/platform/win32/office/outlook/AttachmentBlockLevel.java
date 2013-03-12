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
 * Specifies whether there is any restriction on the type of attachments for an
 * item in a type-safe fashion.
 * <p>
 * Attachments with the BlockLevel equal to olAttachmentBlockLevelOpen are on
 * the Level 2 list of attachments that administrators maintain for attachment
 * security. For more information on attachment security in Outlook, see the
 * Office Resource Kit Web site.
 * </p>
 * 
 * @author Ian Darby
 * 
 * @see AbstractEnum
 */
public class AttachmentBlockLevel extends AbstractEnum {

	/**
	 * There is no restriction on the type of the attachment based on its file
	 * extension.
	 */
	public final static AttachmentBlockLevel olAttachmentBlockLevelNone = new AttachmentBlockLevel(0, "olAttachmentBlockLevelNone");
	
	/**
	 * There is a restriction on the type of the attachment based on its file
	 * extension such that users must first save the attachment to disk before
	 * opening it.
	 */
	public final static AttachmentBlockLevel olAttachmentBlockLevelOpen = new AttachmentBlockLevel(1, "olAttachmentBlockLevelOpen");
	
	/**
	 * One and only constructor. Scope is private to prevent the creation of
	 * anything other than the built-in constant instances.
	 * 
	 * @param value
	 *            numeric value used to represent the enum in external storage.
	 * 
	 * @param name
	 *            constant name given to the enum.
	 * 
	 * @see Attachment
	 */
	private AttachmentBlockLevel(int value, String name) {
		super((short) value, name);
	}
	
	/**
	 * Converts an external storage numeric representation in to one of the
	 * built-in constant objects.
	 * 
	 * @param blockingLevel
	 *            external numeric representation.
	 * 
	 * @return one of the built-in constant objects that represents the enum in
	 *         a type-safe way.
	 */
	public static AttachmentBlockLevel parse(short blockingLevel) {
		
		if (blockingLevel == 0) {
			return olAttachmentBlockLevelNone;
			
		} else {
			return olAttachmentBlockLevelOpen;
		}
	}
	
}
