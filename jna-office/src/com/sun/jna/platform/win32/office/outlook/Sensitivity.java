package com.sun.jna.platform.win32.office.outlook;

public class Sensitivity extends AbstractEnum {
	
	public final static Sensitivity	olNormal		= new Sensitivity(0, "olNormal");		//Normal sensitivity
	public final static Sensitivity	olPersonal		= new Sensitivity(1, "olPersonal");		//Personal
	public final static Sensitivity	olPrivate		= new Sensitivity(2, "olPrivate");		//Private
	public final static Sensitivity	olConfidential	= new Sensitivity(3, "olConfidential");	//Confidential

	private Sensitivity(int val, String name) {
		super((short) val, name);
	}

	public static Sensitivity parse(short val) {
		
		switch(val) {
		
		case 1:
			return olPersonal;
			
		case 2:
			return olPrivate;
			
		case 3:
			return olConfidential;
			
		case 0:
		default:
			return olNormal;
		}
	}
}
