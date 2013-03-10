package com.sun.jna.platform.win32.office.outlook;

public class Gender extends AbstractEnum {
	
	public final static Gender olUnspecified = new Gender(0, "olUnspecified");
	public final static Gender olFemale = new Gender(1, "olFemale");
	public final static Gender olMale = new Gender(2, "olMale");
	
	private Gender(int sex, String name) {
		super((short) sex, name);
	}

	public static Gender parse(short sex) {
		
		switch(sex) {
		
		case 0:
			return olUnspecified;
			
		case 1:
			return olFemale;
			
		case 2:
			return olMale;
			
		default:
			return olUnspecified;
		}
	}

}
