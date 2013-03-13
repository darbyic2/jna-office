package com.sun.jna.platform.win32.atmextra;

import com.sun.jna.platform.win32.OleAuto;
import com.sun.jna.platform.win32.Variant;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant.VARIANT;
import com.sun.jna.platform.win32.office.COMObjectHelper;

public class ExtraSystem extends COMObjectHelper {

	public ExtraSystem() {
		super("EXTRA.System", true);
	}
	
	public ExtraSession getActiveSession() {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_PROPERTYGET, result, this.iDispatch,
				"ActiveSession");

		if (result.getValue() == null || result.getVarType().intValue() != Variant.VT_DISPATCH) {
			return null;
			
		} else {
			return new ExtraSession((IDispatch) result.getValue());
		}
	}
	
	public void quit() {
		
		invokeNoReply("Quit");
	}
	
	public ExtraSessions getSessions() {
		
		return new ExtraSessions(getAutomationProperty("Sessions"));
	}
	
	public static void main(String[] args) {
		
		ExtraSystem atmSys = new ExtraSystem();
		ExtraSessions atmSessions = atmSys.getSessions();
		System.out.println("Session count: " + atmSessions.count());
		
		ExtraSession atmSession = atmSys.getActiveSession();
		if (atmSession == null)
			System.out.println("No active sessions");
		else
			System.out.println("Active Session Name: " + atmSession.getName());
		
		atmSessions = null;
		atmSys = null;
	}
}
