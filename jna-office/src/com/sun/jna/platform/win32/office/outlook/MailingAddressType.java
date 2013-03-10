package com.sun.jna.platform.win32.office.outlook;

public class MailingAddressType extends AbstractEnum {
	
	public final static MailingAddressType olNone = new MailingAddressType(0, "olNone");
	public final static MailingAddressType olHome = new MailingAddressType(1, "olHome");
	public final static MailingAddressType olBusiness = new MailingAddressType(2, "olBusiness");
	public final static MailingAddressType olOther = new MailingAddressType(3, "olOther");
	
	private MailingAddressType(int typ, String name) {
		super((short) typ, name);
	}

	public static MailingAddressType parse(short typ) {
		
		switch(typ) {
		
		case 0:
			return olNone;
			
		case 1:
			return olHome;
			
		case 2:
			return olBusiness;
			
		case 3:
			return olOther;
		
		default:
			return olOther;
		}
	}

}
