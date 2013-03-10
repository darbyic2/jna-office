package com.sun.jna.platform.win32.office.outlook;

public class Importance extends AbstractEnum {
	
	public final static Importance olImportanceLow = new Importance(0, "olImportanceLow");
	public final static Importance olImportanceNormal = new Importance(1, "olImportanceNormal");
	public final static Importance olImportanceHigh = new Importance(2, "olImportanceHigh");

	private Importance(int val, String name) {
		super((short) val, name);
	}

	public static Importance parse(short importanceLevel) {
		
		switch(importanceLevel) {
		
		case 0:
			return olImportanceLow;
			
		case 2:
			return olImportanceHigh;
		
		case 1:
		default:
			return olImportanceNormal;
		}
	}
}
