package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.OleAuto;
import com.sun.jna.platform.win32.COM.COMException;
import com.sun.jna.platform.win32.COM.COMUtils;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant.VARIANT;
import com.sun.jna.platform.win32.WinNT.HRESULT;
import com.sun.jna.platform.win32.office.COMObjectHelper;

public class BaseOutlookObject extends COMObjectHelper {
	
	protected BaseOutlookObject(IDispatch iDispatch) {
		super(iDispatch);
	}
	
	BaseOutlookObject(String progId, boolean useActiveInstance)
			throws COMException {
		super(progId, useActiveInstance);
	}

	
	public Outlook getApplication() {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		HRESULT hr = this.oleMethod(OleAuto.DISPATCH_PROPERTYGET, result, this.iDispatch,
				"Application");

		if (COMUtils.SUCCEEDED(hr))
			return new Outlook((IDispatch) result.getValue());
		else
			return null;
	}
	
	public int getClassEnumValue() {
		
		return getIntProperty("Class");
	}
	
	public Namespace getSession() {
		
		return new Namespace(getAutomationProperty("Session"));
	}
	
}
