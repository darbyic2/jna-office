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
 * Specifies constants that determine whether new items of the conversation are
 * always moved to the Deleted Items folder of the specified delivery store.
 * 
 * @author Ian Darby
 * 
 * @see {@link AbstractEnum}
 */
public class AlwaysDeleteConversation extends AbstractEnum {

	/**
	 * New items joining the conversation are not moved to the the Deleted Items
	 * folder on the specified delivery store, and existing conversation items
	 * in the Deleted Items folder are moved to the Inbox.
	 */
	public final static AlwaysDeleteConversation	olDoNotDelete	= new AlwaysDeleteConversation(0, "olDoNotDelete"); 
	
	/**
	 * New items of the conversation are always moved to the Deleted Items
	 * folder for the store that contains the items
	 */
	public final static AlwaysDeleteConversation	olAlwaysDelete	= new AlwaysDeleteConversation(1, "olAlwaysDelete"); 
	
	/**
	 * The specified store does not support the action of always moving items to
	 * the Deleted Items folder of that store.
	 */
	public final static AlwaysDeleteConversation	olAlwaysDeleteUnsupported	= new AlwaysDeleteConversation(2, "olAlwaysDeleteUnsupported"); 
	
	/**
	 * One and only constructor. Scope is private to prevent the creation of
	 * anything other than the built-in constant instances.
	 * 
	 * @param option
	 *            numeric value used to represent the enum in external storage.
	 * 
	 * @param name
	 *            constant name given to the enum.
	 */
	private AlwaysDeleteConversation(int option, String name) {
		super((short) option, name);
	}
	
	/**
	 * Converts an external storage numeric representation in to one of the
	 * built-in constant objects. Unrecognised external values return a
	 * olDoNotDelete constant.
	 * 
	 * @param option
	 *            external numeric representation.
	 * 
	 * @return one of the built-in constant objects that represents the enum in
	 *         a type-safe way.
	 */
	public static AlwaysDeleteConversation parse(short option) {
		
		switch(option) {
		
		case 0:
			return olDoNotDelete;
			
		case 1:
			return olAlwaysDelete;
			
		case 2:
			return olAlwaysDeleteUnsupported;
			
		default:
			return olDoNotDelete;
		}
	}
}
