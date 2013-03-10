package com.sun.jna.platform.win32.office.outlook;

public class BusyStatus extends AbstractEnum {
	
	public final static BusyStatus	olFree			= new BusyStatus(0, "olFree"); //The user is available.
	public final static BusyStatus	olTentative		= new BusyStatus(1, "olTentative"); //The user has a tentative appointment scheduled.
	public final static BusyStatus	olBusy			= new BusyStatus(2, "olBusy"); //The user is busy.
	public final static BusyStatus	olOutOfOffice	= new BusyStatus(3, "olOutOfOffice"); //The user is out of office.

	private BusyStatus(int val, String name) {
		super((short) val, name);
	}

	public static BusyStatus parse(short val) {
		
		switch(val) {
		
		case 0:
			return olFree;
			
		case 1:
			return olTentative;
			
		case 2:
			return olBusy;
			
		case 3:
			return olOutOfOffice;
			
		default:
			throw new RuntimeException("BusyStatus Enum: " + val + " not recognised.");
		}
	}
}
