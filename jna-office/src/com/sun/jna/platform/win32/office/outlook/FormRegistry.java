package com.sun.jna.platform.win32.office.outlook;

public class FormRegistry extends AbstractEnum {

	public final static FormRegistry	olDefaultRegistry	= new FormRegistry(0, "olDefaultRegistry"); //The Form is registered in the user's default form registry.
	public final static FormRegistry	olPersonalRegistry	= new FormRegistry(2, "olPersonalRegistry"); //The Form is registered in the user's personal registry and is only accessible to that user.
	public final static FormRegistry	olFolderRegistry	= new FormRegistry(3, "olFolderRegistry"); //The Form is registered in a form registry specific to a particular folder, and can only be accessed from that folder.
	public final static FormRegistry	olOrganizationRegistry	= new FormRegistry(4, "olOrganizationRegistry"); //The Form is registered in the organizational form registry. The form is available to all users.

	private FormRegistry(int val, String name) {
		super((short) val, name);
	}
	
	public static FormRegistry parse(short typ) {
		
		switch(typ) {
		
		case 0:
			return olDefaultRegistry;
			
		case 2:
			return olPersonalRegistry;
			
		case 3:
			return olFolderRegistry;
			
		case 4:
			return olOrganizationRegistry;
			
		default:
			throw new RuntimeException("FormRegistry Enum: " + typ + " not recognised.");
		}
	}
}
