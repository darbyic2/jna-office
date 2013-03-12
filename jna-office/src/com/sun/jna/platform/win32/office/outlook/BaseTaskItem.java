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

import com.sun.jna.platform.win32.COM.IDispatch;

/**
 * Base class from which all variations of TaskItem extend.
 * 
 * @author Ian Darby
 * 
 * @see BaseItemLevel2
 * @see TaskItem
 * @see TaskRequestItem
 */
public class BaseTaskItem extends BaseItemLevel2 {

	/**
	 * Constructor scope is restricted to inheritance and package as it should
	 * not be used directly by user applications. It is only intended to be used
	 * from within factory methods and properties of the Outlook object model
	 * itself. It may also be called from unit tests which may supply a mock
	 * version of the IDispatch object.
	 * 
	 * @param iDisp
	 *            the IDispatch object which is the underlying Actions object
	 *            within the Outlook object model. All methods and properties of
	 *            this wrapper class ultimately delegate to IDispatch.
	 */
	protected BaseTaskItem(IDispatch iDisp) {
		super(iDisp);
	}
	
	/**
	 * Obtains a Conversation object that represents the conversation to which
	 * this item belongs.
	 * <p>
	 * GetConversation returns Null (Nothing in Visual Basic) if no conversation
	 * exists for the item. No conversation exists for an item in the following
	 * scenarios:
	 * <ul>
	 * <li>The item has not been saved. An item can be saved programmatically,
	 * by user action, or by auto-save.</li>
	 * 
	 * <li>For an item that can be sent (for example, a mail item, appointment
	 * item, or contact item), the item has not been sent.</li>
	 * 
	 * <li>Conversations have been disabled through the Windows registry.</li>
	 * 
	 * <li>The store does not support Conversation view (for example, Outlook is
	 * running in classic online mode against a version of Microsoft Exchange
	 * earlier than Microsoft Exchange Server 2010). Use the
	 * IsConversationEnabled property of the Store object to determine whether
	 * the store supports Conversation view.</li>
	 * </ul>
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @return a Conversation object that represents the conversation to which
	 *         this item belongs.
	 */
	public Conversation getConversation() {
		
		return new Conversation((IDispatch) invoke("GetConversation").getValue());
	}
	
}
