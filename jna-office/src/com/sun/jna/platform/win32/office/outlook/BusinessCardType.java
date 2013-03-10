package com.sun.jna.platform.win32.office.outlook;

public class BusinessCardType extends AbstractEnum {
	
	public final static BusinessCardType olBusinessCardTypeOutlook = new BusinessCardType(0, "olBusinessCardTypeOutlook");
	public final static BusinessCardType olBusinessCardTypeInterConnect = new BusinessCardType(1, "olBusinessCardTypeInterConnect");

	private BusinessCardType(int typ, String name) {
		super((short) typ, name);
	}

	public static BusinessCardType parse(short typ) {
		
		switch(typ) {
		
		case 0:
			return olBusinessCardTypeOutlook;
			
		case 1:
			return olBusinessCardTypeInterConnect;
		
		default:
			throw new RuntimeException("BusinessCardType Enum: " + typ + " not recognised.");
		}
	}
}
