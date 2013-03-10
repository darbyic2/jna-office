package com.sun.jna.platform.win32.office.outlook;

public class UnknownRecipientType extends RecipientType {

	public final static UnknownRecipientType Unknown = new UnknownRecipientType(0, "Unknown");
	
	private UnknownRecipientType(int typ, String name) {
		super(typ, name);
	}
}
