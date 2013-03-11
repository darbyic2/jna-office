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
 * Specifies how item properties will be copied. Type-safe representation of the
 * Outlook model olActionCopyLike enum.
 * 
 * @author Ian Darby
 * 
 * @see {@link Action}
 * @see {@link AbstractEnum}
 */
public class ActionCopyLike extends AbstractEnum {

	/**
	 * Properties of the new item will be set such that the new item is a reply
	 * to the original item. If creating a new MailItem, the value of the
	 * original To field will be copied to the SenderEmailAddress property of
	 * the new item, the CC field will be blank and the Subject field of the new
	 * item will be the original Subject with a prefix such as "RE:" (see Prefix
	 * property) added.
	 */
	public final static ActionCopyLike olReply = new ActionCopyLike(0, "olReply");
	
	/**
	 * Properties of the new item will be set such that the new item is a reply
	 * to all of the senders of the original item. If creating a new MailItem,
	 * the value of the SenderEmailAddress and CC properties will be copied to
	 * the To property of the new item and the Subject property of the new item
	 * will be the Subject of the original item with a prefix such as "RE:" (see
	 * Prefix property) added.
	 */
	public final static ActionCopyLike olReplyAll = new ActionCopyLike(1, "olReplyAll");
	
	/**
	 * Properties of the new item will be set such that the new item is a
	 * forward of the original item. If creating a new MailItem, the value of
	 * the To and CC properties in the new item will be empty and the Subject
	 * property of the new item will be the original Subject with a prefix such
	 * as "FW:" (see Prefix property) added. The attachments on the original
	 * item will be copied to the new item.
	 */
	public final static ActionCopyLike olForward = new ActionCopyLike(2, "olForward");
	
	/**
	 * If creating a new PostItem based on an old one, the Post To property of
	 * the new item will contain the active folder address, the Subject property
	 * of the original item will be copied to the ConversationTopic property of
	 * the new item, and the Subject property of the new item will be empty.
	 */
	public final static ActionCopyLike olReplyFolder = new ActionCopyLike(3, "olReplyFolder");
	
	/**
	 * Used exclusively for voting button actions.
	 */
	public final static ActionCopyLike olRespond = new ActionCopyLike(4, "olRespond");
	
	/**
	 * One and only constructor. Scope is private to prevent the creation of
	 * anything other than the built-in constant instances.
	 * 
	 * @param val
	 *            numeric value used to represent the enum in external storage.
	 * 
	 * @param name
	 *            constant name given to the enum.
	 * 
	 * @see {@link Action}
	 */
	private ActionCopyLike(int val, String name) {
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
	public static ActionCopyLike parse(short val) {
		
		switch(val) {
		
		case 0:
			return olReply;
			
		case 1:
			return olReplyAll;
			
		case 2:
			return olForward;
			
		case 3:
			return olReplyFolder;
			
		case 4:
			return olRespond;
			
		default:
			throw new RuntimeException("ActionCopyLike Enum: " + val + " not recognised.");
			
		}
	}
}
