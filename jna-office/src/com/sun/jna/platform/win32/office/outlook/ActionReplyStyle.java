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
 * Specifies the reply style. Type-safe representation of the
 * Outlook model olActionReplyStyle enum.
 * 
 * @author Ian Darby
 * 
 * @see {@link Action}
 * @see {@link AbstractEnum}
 */
public class ActionReplyStyle extends AbstractEnum {
	
	/**
	 * The reply will not include any references to the original item or its
	 * text.
	 */
	public final static ActionReplyStyle	olOmitOriginalText		= new ActionReplyStyle(	0	, "olOmitOriginalText");  
	
	/**
	 * The reply will include the original item embedded in it.
	 */
	public final static ActionReplyStyle	olEmbedOriginalItem		= new ActionReplyStyle(	1	, "olEmbedOriginalItem");  
	
	/**
	 * The reply will include the text of the original item.
	 */
	public final static ActionReplyStyle	olIncludeOriginalText	= new ActionReplyStyle(	2	, "olIncludeOriginalText");  
	
	/**
	 * The reply will include the indented text of the original item.
	 */
	public final static ActionReplyStyle	olIndentOriginalText	= new ActionReplyStyle(	3	, "olIndentOriginalText");  
	
	/**
	 * The reply will include a link to the original item.
	 */
	public final static ActionReplyStyle	olLinkOriginalItem		= new ActionReplyStyle(	4	, "olLinkOriginalItem");  
	
	/**
	 * The reply style will be set based on the user's preference.
	 */
	public final static ActionReplyStyle	olUserPreference		= new ActionReplyStyle(	5	, "olUserPreference");  
	
	/**
	 * The reply will include the original text with each line preceded by a
	 * symbol such as ">".
	 */
	public final static ActionReplyStyle	olReplyTickOriginalText	= new ActionReplyStyle(	1000, "olReplyTickOriginalText");  
	
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
	private ActionReplyStyle(int val, String name) {
		super((short) val, name);
	}

	/**
	 * Converts an external storage numeric representation in to one of the
	 * built-in constant objects. Unrecognised external values cause a
	 * {@link RuntimeException} to be thrown.
	 * 
	 * @param style
	 *            external numeric representation.
	 * 
	 * @return one of the built-in constant objects that represents the enum in
	 *         a type-safe way.
	 */
	public static ActionReplyStyle parse(short style) {
		switch(style) {
		
		case 0:
			return olOmitOriginalText;
			
		case 1:
			return olEmbedOriginalItem;
			
		case 2:
			return olIncludeOriginalText;
			
		case 3:
			return olIndentOriginalText;
			
		case 4:
			return olLinkOriginalItem;
			
		case 5:
			return olUserPreference;
			
		case 1000:
			return olReplyTickOriginalText;
			
		default:
			throw new RuntimeException("ActionReplyStyle Enum: " + style + " not recognised.");
		}
	}
}
