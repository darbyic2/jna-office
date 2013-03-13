package com.sun.jna.platform.win32.atmextra;

import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.office.COMObjectHelper;

public class ExtraSessions extends COMObjectHelper {

	ExtraSessions(IDispatch iDisp) {
		super(iDisp);
	}
	
	public int count() {
		
		return getIntProperty("Count");
	}
}
