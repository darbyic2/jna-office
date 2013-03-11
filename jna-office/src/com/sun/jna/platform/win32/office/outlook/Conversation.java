package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.Variant;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant.VARIANT;
import com.sun.jna.platform.win32.WinDef.LONG;

public class Conversation extends BaseOutlookObject {

	Conversation(IDispatch iDisp) {
		super(iDisp);
	}
	
	public void clearAlwaysAssignCategories(Store store) {
		
		invokeNoReply("ClearAlwaysAssignCategories", newVariant(store.getIDispatch()));
	}
	
	public String getConversationID() {
		
		return getStringProperty("ConversationID");
	}
	
	public String getAlwaysAssignCategories(Store store) {
		
		return invoke("GetAlwaysAssignCategories", newVariant(store.getIDispatch())).getValue().toString();
	}
	
	public AlwaysDeleteConversation getAlwaysDelete(Store store) {
		
		return AlwaysDeleteConversation.parse(((LONG) invoke("GetAlwaysDelete", newVariant(store.getIDispatch())).getValue()).shortValue());
	}
	
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
	
	public Table getTable() {
		
		return new Table((IDispatch) invoke("GetTable").getValue());
	}
	
	public void markAsRead() {
		
		invokeNoReply("MarkAsRead");
	}
	
	public void markAsUnread() {
		
		invokeNoReply("MarkAsUnread");
	}
	
	public void setAlwaysAssignCategories(String categories, Store store) {
		
		invokeNoReply("SetAlwaysAssignCategories", newVariant(categories), newVariant(store.getIDispatch()));
	}
	
	public void setAlwaysDelete(AlwaysDeleteConversation option, Store store) {
		
		invokeNoReply("SetAlwaysDelete", newVariant(option.value()), newVariant(store.getIDispatch()));
	}
	
	public void setAlwaysMoveToFolder(Folder destFolder, Store store) {
		
		invokeNoReply("SetAlwaysMoveToFolder", newVariant(destFolder.getIDispatch()), newVariant(store.getIDispatch()));
	}
	
	public void stopAlwaysDelete(Store store) {
		
		invokeNoReply("StopAlwaysDelete", newVariant(store.getIDispatch()));
	}
	
	public void stopAlwaysMoveToFolder(Store store) {
		
		invokeNoReply("StopAlwaysMoveToFolder", newVariant(store.getIDispatch()));
	}
	
}
