package com.sun.jna.platform.win32.atmextra;

import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.office.COMObjectHelper;

public class ExtraSession extends COMObjectHelper {

	ExtraSession(IDispatch iDisp) {
		super(iDisp);
	}

	public String getName() {
		
		return getStringProperty("Name");
	}
}
