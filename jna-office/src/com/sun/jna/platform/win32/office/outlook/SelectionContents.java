package com.sun.jna.platform.win32.office.outlook;

public class SelectionContents extends AbstractEnum {

	public final static SelectionContents olConversionHeaders = new SelectionContents(1, "olConversionHeaders");
	
	private SelectionContents(int val, String name) {
		super((short) val, name);
	}
	
	public static SelectionContents parse(short typ) {
		
		switch(typ) {
		
		case 1:
			return olConversionHeaders;
			
		default:
			throw new RuntimeException("SelectionContents Enum: " + typ + " not recognised.");
		}
	}
}
