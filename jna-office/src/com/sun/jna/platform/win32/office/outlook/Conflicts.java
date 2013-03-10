package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.Variant;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant.VARIANT;

public class Conflicts extends BaseOutlookObject {

	Conflicts(IDispatch iDisp) {
		super(iDisp);
	}

	public int count() {

		return getIntProperty("Count");
	}
	
	private Conflict getConflictFromMethod(String methodName) {

		VARIANT result = invoke(methodName);

		if (result == null
				|| result.getVarType().intValue() == Variant.VT_EMPTY
				|| result.getVarType().intValue() == Variant.VT_NULL
				|| result.getVarType().intValue() != Variant.VT_DISPATCH) {

			return null;

		} else {
			return new Conflict((IDispatch) result.getValue());
		}
	}

	public Conflict getFirst() {
		
		return getConflictFromMethod("GetFirst");
	}

	public Conflict getLast() {
		
		return getConflictFromMethod("GetLast");
	}

	public Conflict getNext() {
		
		return getConflictFromMethod("GetNext");
	}

	public Conflict getPrevious() {
		
		return getConflictFromMethod("GetPrevious");
	}

	public Conflict getItem(int index) {
		
		return new Conflict((IDispatch) invoke("Item", newVariant(index)).getValue());
	}
	
	public Conflict getItem(String conflictName) {
		
		return new Conflict((IDispatch) invoke("Item", newVariant(conflictName)).getValue());
	}
	
}
