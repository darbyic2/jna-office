package com.sun.jna.platform.win32.office.outlook;

public class MobileFormat extends AbstractEnum {
	
	public final static MobileFormat olSMS = new MobileFormat(0, "olSMS");
	public final static MobileFormat olMMS = new MobileFormat(1, "olMMS");
	
	private MobileFormat(int format, String name) {
		super((short) format, name);
	}

	public static MobileFormat parse(short format) {
		
		switch(format) {
		
		case 0:
			return olSMS;
			
		case 1:
			return olMMS;
			
		default:
			throw new RuntimeException("MobileFormat Enum: " + format + " not recognised.");
		}
	}
}
