package com.sun.jna.platform.win32.office.outlook;

public class Permission extends AbstractEnum {

	public final static Permission olUnrestricted = new Permission(0, "olUnrestricted");
	public final static Permission olDoNotForward = new Permission(1, "olDoNotForward");
	public final static Permission olPermissionTemplate = new Permission(2, "olPermissionTemplate");
	
	private Permission(int val, String name) {
		super((short) val, name);
	}

	public static Permission parse(short permission) {
		
		switch(permission) {
		
		case 1:
			return olDoNotForward;
			
		case 2:
			return olPermissionTemplate;
		
		case 0:
		default:
			return olUnrestricted;
		}
	}
}
