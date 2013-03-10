package com.sun.jna.platform.win32.office.outlook;

public abstract class AbstractEnum {
	
	private short val;
	private String name;

	protected AbstractEnum(short val, String name) {
		super();
		this.val = val;
		this.name = name;
	}
	
	public short value() {
		
		return val;
	}

	@Override
	public String toString() {
		
		return name + "(" + val + ")";
	}

}
