package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.Variant;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant.VARIANT;

public class Folders extends BaseOutlookObject {

	public Folders(IDispatch iDispatch) {
		super(iDispatch);
	}
	
	public Folder add(String name) {
		
		return new Folder((IDispatch) invoke("Add", newVariant(name)).getValue());
	}
	
	public Folder add(String name, FolderType typ) {
		
		return new Folder((IDispatch) invoke("Add", newVariant(name), newVariant(typ.value())).getValue());
	}
	
	public int count() {
		
		return getIntProperty("Count");
	}
	
	private Folder getHelper(String command) {
		
		VARIANT var = invoke(command);
		
		if (var == null || var.getVarType().intValue() == Variant.VT_EMPTY || var.getVarType().intValue() == Variant.VT_NULL) {
			return null;
			
		} else {
			return new Folder((IDispatch) var.getValue());
		}
	}
	
	public Folder getFirst() {
		
		return getHelper("GetFirst");
	}
	
	public Folder getLast() {
		
		return getHelper("GetLast");
	}
	
	public Folder getNext() {
		
		return getHelper("GetNext");
	}
	
	public Folder getPrevious() {
		
		return getHelper("GetPrevious");
	}
	
	public Folder get(int index) {
		
		return new Folder((IDispatch) invoke("Item", newVariant(index)).getValue());
	}
	
	public Folder get(String folderName) {
		
		return new Folder((IDispatch) invoke("Item", newVariant(folderName)).getValue());
	}
	
	public void remove(int index) {
		
		invokeNoReply("Display", newVariant(index));
	}
	
}
