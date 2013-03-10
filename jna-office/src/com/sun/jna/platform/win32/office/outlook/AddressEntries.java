package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.Variant;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant.VARIANT;

public class AddressEntries extends BaseOutlookObject {

	AddressEntries(IDispatch iDisp) {
		super(iDisp);
	}
	
	public AddressEntry add(String addressType) {
		
		return new AddressEntry((IDispatch) invoke("Add", newVariant(addressType)).getValue());
	}
	
	public AddressEntry add(String addressType, String entryName) {
		
		return new AddressEntry((IDispatch) invoke("Add", newVariant(addressType), newVariant(entryName)).getValue());
	}
	
	public AddressEntry add(String addressType, String entryName, String address) {
		
		return new AddressEntry((IDispatch) invoke("Add", newVariant(addressType), newVariant(entryName), newVariant(address)).getValue());
	}
	
	public int count() {

		return getIntProperty("Count");
	}
	
	private AddressEntry getEntryFromMethod(String methodName) {

		VARIANT result = invoke(methodName);

		if (result == null
				|| result.getVarType().intValue() == Variant.VT_EMPTY
				|| result.getVarType().intValue() == Variant.VT_NULL
				|| result.getVarType().intValue() != Variant.VT_DISPATCH) {

			return null;

		} else {
			return new AddressEntry((IDispatch) result.getValue());
		}
	}

	public AddressEntry getFirst() {
		
		return getEntryFromMethod("GetFirst");
	}

	public AddressEntry getLast() {
		
		return getEntryFromMethod("GetLast");
	}

	public AddressEntry getNext() {
		
		return getEntryFromMethod("GetNext");
	}

	public AddressEntry getPrevious() {
		
		return getEntryFromMethod("GetPrevious");
	}

	public AddressEntry getItem(int index) {
		
		return new AddressEntry((IDispatch) invoke("Item", newVariant(index)).getValue());
	}
	
	public AddressEntry getItem(String entryName) {
		
		return new AddressEntry((IDispatch) invoke("Item", newVariant(entryName)).getValue());
	}
	
	public void sort(String propertyNameToSortOn, SortOrder order) {
		
		String prop;
		
		if (propertyNameToSortOn.startsWith("[")) {
			prop = propertyNameToSortOn;
			
		} else {
			prop = "[" + propertyNameToSortOn + "]";
		}
		invokeNoReply("Sort", newVariant(prop), newVariant(order.value()));
	}
}
