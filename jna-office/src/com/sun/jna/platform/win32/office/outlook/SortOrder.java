package com.sun.jna.platform.win32.office.outlook;

public class SortOrder extends AbstractEnum {

	public final static SortOrder olSortNone = new SortOrder(0, "olSortNone");
	public final static SortOrder olAscending = new SortOrder(1, "olAscending");
	public final static SortOrder olDescending = new SortOrder(2, "olDescending");
	
	private SortOrder(int val, String name) {
		super((short) val, name);
	}
	
	public static SortOrder parse(short val) {
		
		switch(val) {
		
		case 0:
			return olSortNone;
			
		case 1:
			return olAscending;
			
		case 2:
			return olDescending;
			
		default:
			return olSortNone;
		}
	}
}
