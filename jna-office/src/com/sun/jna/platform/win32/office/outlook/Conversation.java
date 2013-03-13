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

import com.sun.jna.platform.win32.Variant;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant.VARIANT;
import com.sun.jna.platform.win32.WinDef.LONG;

/**
 * Represents a conversation that includes one or more items stored in one or
 * more folders and stores.
 * <p>
 * The Conversation object is an abstract, aggregated object. Although a
 * conversation can include items (item: An item is the basic element that holds
 * information in Outlook (similar to a file in other programs). Items include
 * e-mail messages, appointments, contacts, tasks, journal entries, notes,
 * posted items, and documents.) of different types, the Conversation object
 * does not correspond to a particular underlying MAPI IMessage object.
 * </p>
 * <p>
 * A conversation represents one or more items in one or more folders and
 * stores. If you move an item in a conversation to the Deleted Items folder and
 * subsequently enumerate the conversation by using the GetChildren,
 * GetRootItems, or GetTable method, the item will not be included in the
 * returned object.
 * </p>
 * <p>
 * To obtain a Conversation object for an existing conversation, use the
 * GetConversation method of the item.
 * </p>
 * <p>
 * There are actions that you can apply to items in a conversation by calling
 * the SetAlwaysAssignCategories, SetAlwaysDelete, or SetAlwaysMoveToFolder
 * method. Each of these actions is applied to all items in the conversation
 * automatically when the method is called; the action is also applied to future
 * items in the conversation as long as the action is still applicable to the
 * conversation. There is no explicit save method on the Conversation object.
 * </p>
 * <p>
 * Also, when you apply an action to items in a conversation, the corresponding
 * event occurs. For example, the ItemChange event of the Items object occurs
 * when you call SetAlwaysAssignCategories, and the BeforeItemMove event of the
 * Folder object occurs when you call SetAlwaysMoveToFolder.
 * </p>
 * <p>
 * Added in Outlook 2010.
 * </p>
 * 
 * @author Ian Darby
 * 
 * @see BaseOutlookObject
 */
public class Conversation extends BaseOutlookObject {

	/**
	 * Constructor scope is restricted to package as it should not be used
	 * directly by user applications. It is only intended to be used from within
	 * factory methods and properties of the Outlook object model itself. It may
	 * also be called from unit tests which may supply a mock version of the
	 * IDispatch object.
	 * 
	 * @param iDisp
	 *            the IDispatch object which is the underlying Actions object
	 *            within the Outlook object model. All methods and properties of
	 *            this wrapper class ultimately delegate to IDispatch.
	 */
	Conversation(IDispatch iDisp) {
		super(iDisp);
	}
	
	/**
	 * Removes all categories from all items in the conversation and stops the
	 * action of always assigning categories to items in the conversation.
	 * <p>
	 * If the store specified by the Store parameter represents a non-delivery
	 * store such as an archive .pst store, the category removal action will
	 * apply to items of the conversation in the default delivery store.
	 * </p>
	 * <p>
	 * After you apply the ClearAlwaysAssignCategories method on a conversation,
	 * the GetAlwaysAssignCategories method will return Null (Nothing in Visual
	 * Basic) for that conversation. Categories on existing items are cleared,
	 * and no categories are assigned to new items in the conversation.
	 * </p>
	 * <p>
	 * If the SetAlwaysAssignCategories method has not been applied to a
	 * conversation, ClearAlwaysAssignCategories does not remove any categories.
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @param store
	 *            Specifies the store from which categories of items that belong
	 *            to the conversation should be removed.
	 */
	public void clearAlwaysAssignCategories(Store store) {
		
		invokeNoReply("ClearAlwaysAssignCategories", newVariant(store.getIDispatch()));
	}
	
	/**
	 * Returns a String that uniquely identifies a Conversation object.
	 * Read-only.
	 * <p>
	 * This property associates items with a conversation. These items and the
	 * conversation all have the same value in their ConversationID property.
	 * </p>
	 * <p>
	 * This property corresponds with the MAPI property PidTagConversationId.
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @return a String that uniquely identifies a Conversation object.
	 */
	public String getConversationID() {
		
		return getStringProperty("ConversationID");
	}
	
	/**
	 * Returns a String that indicates the category or categories that are
	 * assigned to all new items that arrive in the conversation.
	 * <p>
	 * Multiple categories are delimited by commas in the string of category
	 * names that this property returns. To convert the string of category names
	 * to an array of category names, use the Microsoft Visual Basic Split
	 * function.
	 * </p>
	 * <p>
	 * If the store specified by the Store parameter represents a non-delivery
	 * store such as an archive .pst store, the method returns a string of
	 * categories that are applied to conversation items in the default delivery
	 * store.
	 * </p>
	 * <p>
	 * If the SetAlwaysAssignCategories method has not been applied to a
	 * conversation, GetAlwaysAssignCategories returns Null (Nothing in Visual
	 * Basic).
	 * </p>
	 * <p>
	 * To stop the action of always assigning categories, use the
	 * ClearAlwaysAssignCategories method. After the ClearAlwaysAssignCategories
	 * method has been called, GetAlwaysAssignCategories returns an empty
	 * string.
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @param store
	 *            Specifies the store to which categories of items that belong
	 *            to the conversation should be returned.
	 * 
	 * @return a String that indicates the category or categories that are
	 *         assigned to all new items that arrive in the conversation.
	 */
	public String getAlwaysAssignCategories(Store store) {
		
		return invoke("GetAlwaysAssignCategories", newVariant(store.getIDispatch())).getValue().toString();
	}
	
	/**
	 * Returns a constant in the AlwaysDeleteConversation enumeration that
	 * indicates whether all new items that join the conversation are always
	 * moved to the Deleted Items folder in the specified delivery store.
	 * <p>
	 * If the Store parameter specifies a non-delivery store such as an archive
	 * .pst store, the GetAlwaysDelete method returns a constant from
	 * OlAlwaysDeleteConversation that applies to conversation items in the
	 * default delivery store. Items on a non-delivery store are not moved to
	 * the Deleted Items folder for the default delivery store.
	 * </p>
	 * <p>
	 * If GetAlwaysDelete returns olAlwaysDelete, items of the conversation are
	 * always moved to the Deleted Items folder for the store that contains the
	 * items. In a cross-store conversation, items are moved to the Deleted
	 * Items folder for the store that contains the items. When GetAlwaysDelete
	 * returns olAlwaysDelete, the GetAlwaysMoveToFolder method returns a folder
	 * object that represents the Deleted Items folder for the default store.
	 * </p>
	 * <p>
	 * If GetAlwaysDelete returns olAlwaysDeleteUnsupported, the specified store
	 * does not support the action of always moving items to the Deleted Items
	 * folder of that store.
	 * </p>
	 * <p>
	 * If GetAlwaysDelete returns olDoNotDelete, new items that arrive in the
	 * conversation are not moved to the the Deleted Items folder on the
	 * specified delivery store, and existing conversation items in the Deleted
	 * Items folder are moved to the Inbox.
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @param store
	 *            Specifies the store that holds the Deleted Items folder to
	 *            which items of the conversation are moved.
	 * 
	 * @return a constant in the AlwaysDeleteConversation enumeration that
	 *         indicates whether all new items that join the conversation are
	 *         always moved to the Deleted Items folder in the specified
	 *         delivery store.
	 */
	public AlwaysDeleteConversation getAlwaysDelete(Store store) {
		
		return AlwaysDeleteConversation.parse(((LONG) invoke("GetAlwaysDelete", newVariant(store.getIDispatch())).getValue()).shortValue());
	}
	
	/**
	 * Returns a Folder object that indicates the folder in the specified
	 * delivery store to which new items that arrive in the conversation are
	 * always moved.
	 * <p>
	 * If the Store parameter represents a non-delivery store such as an archive
	 * .pst store, the GetAlwaysMoveToFolder method returns a Folder object that
	 * applies to conversation items on the default delivery store.
	 * </p>
	 * <p>
	 * If no folder, other than the Deleted Items folder, has been specified to
	 * always move conversation items into, the GetAlwaysMoveToFolder method
	 * returns Null (Nothing in Visual Basic).
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @param store
	 *            The store where the folder to which conversation items are
	 *            moved resides.
	 * 
	 * @return a Folder object that indicates the folder in the specified
	 *         delivery store to which new items that arrive in the conversation
	 *         are always moved.
	 */
	public Folder getAlwaysMoveToFolder(Store store) {

		VARIANT result = invoke("GetAlwaysMoveToFolder", newVariant(store
				.getIDispatch()));

		if (result == null
				|| result.getVarType().intValue() == Variant.VT_EMPTY
				|| result.getVarType().intValue() == Variant.VT_NULL
				|| result.getVarType().intValue() != Variant.VT_DISPATCH) {

			return null;

		} else {
			return new Folder((IDispatch) result.getValue());
		}
	}
	
	/**
	 * Returns a SimpleItems collection that contains all items under the
	 * specified conversation node.
	 * <p>
	 * The returned SimpleItems collection contains immediate child items of the
	 * conversation node specified by the Item parameter. If the specified node
	 * does not exist in the conversation, the GetChildren method returns an
	 * error.
	 * </p>
	 * <p>
	 * If no child items exist under that node, the GetChildren method returns a
	 * SimpleItems collection with zero objects, in which case the Count
	 * property of the SimpleItems collection returns 0.
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @param item
	 *            A conversation node that is part of a conversation.
	 * 
	 * @return a SimpleItems collection that contains all items under the
	 *         specified conversation node.
	 */
	public SimpleItems getChildren(BaseItemLevel1 item) {
		
		VARIANT result = invoke("GetChildren", newVariant(item
				.getIDispatch()));

		if (result == null
				|| result.getVarType().intValue() == Variant.VT_EMPTY
				|| result.getVarType().intValue() == Variant.VT_NULL
				|| result.getVarType().intValue() != Variant.VT_DISPATCH) {

			return null;

		} else {
			return new SimpleItems((IDispatch) result.getValue());
		}
	}
	
	/**
	 * Returns a SimpleItems collection that contains all root items in the
	 * conversation.
	 * <p>
	 * A conversation can have one or more root items. For example, if the root
	 * item of the conversation has three child items and the root item is
	 * permanently deleted, all three child items become root items.
	 * </p>
	 * <p>
	 * If all items are deleted from the conversation after the Conversation
	 * object has been obtained, GetRootItems returns a SimpleItems collection
	 * with zero objects. In this case, the Count property of the SimpleItems
	 * collection returns 0.
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @return a SimpleItems collection that contains all root items in the
	 *         conversation.
	 */
	public SimpleItems getRootItems() {
		
		VARIANT result = invoke("GetChildren");

		if (result == null
				|| result.getVarType().intValue() == Variant.VT_EMPTY
				|| result.getVarType().intValue() == Variant.VT_NULL
				|| result.getVarType().intValue() != Variant.VT_DISPATCH) {

			return null;

		} else {
			return new SimpleItems((IDispatch) result.getValue());
		}
	}
	
	/**
	 * Returns a Table object that contains rows that represent all items in the
	 * conversation.
	 * <p>
	 * The GetTable method returns a Table that has all items of the
	 * conversation as the rows. The default set of columns is shown in the
	 * following table.
	 * <ul>
	 * <li>1 - EntryID</li>
	 * 
	 * <li>2 - Subject</li>
	 * 
	 * <li>3 - CreationTime</li>
	 * 
	 * <li>4 - LastModificationTime</li>
	 * 
	 * <li>5 - MessageClass</li>
	 * </ul>
	 * </p>
	 * <p>
	 * By default, the rows in the table are sorted by the ConversationIndex
	 * property of the items.
	 * </p>
	 * <p>
	 * To modify the default column set, use the Add, Remove, or RemoveAll
	 * methods of the Columns collection object.
	 * </p>
	 * <p>
	 * The Table object returned by this GetTable method does not include items
	 * in the conversation that have been moved to the Deleted Items folder.
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @return a Table object that contains rows that represent all items in the
	 *         conversation.
	 */
	public Table getTable() {
		
		return new Table((IDispatch) invoke("GetTable").getValue());
	}
	
	/**
	 * Marks all items in the conversation as read.
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 */
	public void markAsRead() {
		
		invokeNoReply("MarkAsRead");
	}
	
	/**
	 * Marks all items in the conversation as unread.
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 */
	public void markAsUnread() {
		
		invokeNoReply("MarkAsUnread");
	}
	
	/**
	 * Applies one or more categories to all existing items and future items of
	 * the conversation.
	 * <p>
	 * If the store specified by the Store parameter represents a non-delivery
	 * store such as an archive .pst store, the method returns a string of
	 * categories that are applied to conversation items in the default delivery
	 * store.
	 * </p>
	 * <p>
	 * The ItemChange event of the Items object occurs when you call the
	 * SetAlwaysAssignCategories method on a conversation.
	 * </p>
	 * <p>
	 * To determine existing master categories for the current user, examine the
	 * Categories property of the Store object that is specified by the Store
	 * parameter. If one or more categories specified by the Categories
	 * parameter do not exist in the master categories collection, the
	 * categories will be assigned to the conversation but will not be added to
	 * the master categories collection.
	 * </p>
	 * <p>
	 * To determine the existing categories that are always assigned to items of
	 * the conversation in the specified store, use the
	 * GetAlwaysAssignCategories method.
	 * </p>
	 * <p>
	 * If SetAlwaysAssignCategories is called more than once, the result is
	 * cumulative. For example, if you call SetAlwaysAssignCategories specifying
	 * the category “Important” and then call SetAlwaysAssignCategories again
	 * specifying the categories "Business" and "Social", the categories that
	 * are always assigned are "Important", "Business", and "Social".
	 * </p>
	 * <p>
	 * To stop the action of always assigning categories, use the
	 * ClearAlwaysAssignCategories method. After the ClearAlwaysAssignCategories
	 * method has been called, GetAlwaysAssignCategories returns an empty
	 * string.
	 * </p>
	 * <p>
	 * The SetAlwaysAssignToCategories method ignores any category names that
	 * are empty strings. For example, if the Categories parameter is set to the
	 * string "Work,,Play", "Work" and "Play" are assigned to the conversation
	 * and the empty string category is ignored.
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @param categories
	 *            A comma-delimited string of one or more category names that
	 *            are always assigned to all items in the conversation.
	 * 
	 * @param store
	 *            The store in which items of the conversation should always be
	 *            assigned the categories specified by the Categories parameter.
	 */
	public void setAlwaysAssignCategories(String categories, Store store) {
		
		invokeNoReply("SetAlwaysAssignCategories", newVariant(categories), newVariant(store.getIDispatch()));
	}
	
	/**
	 * Specifies a setting for the specified delivery store that indicates
	 * whether all existing items and all new items that arrive in the
	 * conversation are always moved to the Deleted Items folder in the
	 * specified delivery store.
	 * <p>
	 * The SetAlwaysDelete method operates on conversation items in the delivery
	 * store specified by the Store parameter. If the store specified by the
	 * Store parameter represents a non-delivery store such as an archive .pst
	 * store, the action is applied to conversation items in the default
	 * delivery store.
	 * </p>
	 * <p>
	 * If the AlwaysDelete parameter is olAlwaysDelete, conversation items are
	 * moved to the Deleted Items folder for the specfied store. In this case,
	 * the items are not permanently deleted, unless the user has specified a
	 * separate option to permanently delete items when Microsoft Outlook shuts
	 * down.
	 * </p>
	 * <p>
	 * If SetAlwaysDelete returns olDoNotDelete, existing conversation items and
	 * new items that arrive in the conversation are not moved to the the
	 * Deleted Items folder in the specified delivery store, and existing
	 * conversation items in the Deleted Items folder are moved to the Inbox.
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @param option
	 *            A constant that indicates whether all existing and new items
	 *            that arrive in the conversation are always moved to the
	 *            Deleted Folder of the store specified by the Store parameter.
	 * 
	 * @param store
	 *            Specifies the store that contains the Deleted Items folder to
	 *            which existing and new items of the conversation are to be
	 *            moved.
	 */
	public void setAlwaysDelete(AlwaysDeleteConversation option, Store store) {
		
		invokeNoReply("SetAlwaysDelete", newVariant(option.value()), newVariant(store.getIDispatch()));
	}
	
	/**
	 * Sets a Folder object that indicates the folder to which all existing
	 * conversation items and new items that arrive in the conversation are
	 * always moved.
	 * <p>
	 * The SetAlwaysMoveToFolder method operates on conversation items in the
	 * delivery store specified by the Store parameter. If the Store parameter
	 * represents a non-delivery store such as an archive .pst store, the move
	 * action will apply to conversation items in the default delivery store.
	 * </p>
	 * <p>
	 * If the MoveToFolder parameter specifies an invalid folder that does not
	 * exist, has been moved, or is read-only, Outlook will raise an error.
	 * </p>
	 * <p>
	 * To stop the always-move-to-folder action for conversations items in a
	 * store, call the StopAlwaysMoveToFolder method.
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @param destFolder
	 *            Specifies the folder to which all existing items and new items
	 *            that arrive in the conversation are always moved.
	 * 
	 * @param store
	 *            Specifies the store that contains the folder to which items of
	 *            the conversation are moved.
	 */
	public void setAlwaysMoveToFolder(Folder destFolder, Store store) {
		
		invokeNoReply("SetAlwaysMoveToFolder", newVariant(destFolder.getIDispatch()), newVariant(store.getIDispatch()));
	}
	
	/**
	 * Stops the action of always moving conversation items in the specified
	 * store to the Deleted Items folder in that store.
	 * <p>
	 * If the always-delete action has not been turned on, StopAlwaysDelete does
	 * not carry out any action.
	 * </p>
	 * <p>
	 * If the always-delete action has been turned on, StopAlwaysDelete moves
	 * existing conversation items in the Deleted Items folder to the Inbox.
	 * </p>
	 * <p>
	 * After calling the StopAlwaysDelete method for a conversation in a store,
	 * calling the GetAlwaysDelete method on that conversation and store returns
	 * the constant olDoNotDelete.
	 * </p>
	 * <p>
	 * If the store specified by the Store parameter represents a non-delivery
	 * store such as an archive .pst store, the stop-always-delete action is
	 * applied to conversation items in the default delivery store.
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @param store
	 *            Specifies the store to which the stop-always-delete action
	 *            applies.
	 */
	public void stopAlwaysDelete(Store store) {
		
		invokeNoReply("StopAlwaysDelete", newVariant(store.getIDispatch()));
	}
	
	/**
	 * Stops the action of always moving conversation items in the specified
	 * store to a specific folder.
	 * <p>
	 * If the always-move action has not been turned on, StopAlwaysMoveToFolder
	 * does not carry out any action.
	 * </p>
	 * <p>
	 * If the Store parameter represents a non-delivery store such as an archive
	 * .pst store, the stop-always-move action will apply to conversation items
	 * in the default delivery store.
	 * </p>
	 * <p>
	 * After you call the StopAlwaysMoveToFolder method, calling the
	 * GetAlwaysMoveToFolder method returns Null (Nothing in Visual Basic).
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @param store
	 *            The store where the conversation items to be cleaned up
	 *            reside.
	 */
	public void stopAlwaysMoveToFolder(Store store) {
		
		invokeNoReply("StopAlwaysMoveToFolder", newVariant(store.getIDispatch()));
	}
	
}
