package com.sun.jna.platform.win32.office.outlook;

public class PermissionService extends AbstractEnum {
	
	public final static PermissionService olUnknown = new PermissionService(0, "olUnknown");
	public final static PermissionService olWindows = new PermissionService(1, "olWindows");
	public final static PermissionService olPassport = new PermissionService(2, "olPassport");

	private PermissionService(int svc, String name) {
		super((short) svc, name);
	}

	public static PermissionService parse(short svc) {
		
		switch(svc) {
		
		case 1:
			return olWindows;
			
		case 2:
			return olPassport;
			
		case 0:
		default:
			return olUnknown;
		}
	}
}
