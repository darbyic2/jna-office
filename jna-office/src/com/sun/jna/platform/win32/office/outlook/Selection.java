package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class Selection extends BaseOutlookObject {

	Selection(IDispatch iDisp) {
		super(iDisp);
	}
	
	public int count() {

		return getIntProperty("Count");
	}
	
	public BaseItemLevel1 getItem(int index) {
		
		return new BaseItemLevel1((IDispatch) invoke("Item", newVariant(index)).getValue());
	}
	
	public BaseItemLevel1 getItem(String name) {
		
		return new BaseItemLevel1((IDispatch) invoke("Item", newVariant(name)).getValue());
	}
	
	public SelectionLocation getLocation() {
		
		return SelectionLocation.parse(getShortProperty("Location"));
	}
	
	public Selection getSelection(SelectionContents contents) {
		
		return new Selection((IDispatch) invoke("GetSelection", newVariant(contents.value())).getValue());
	}
}
