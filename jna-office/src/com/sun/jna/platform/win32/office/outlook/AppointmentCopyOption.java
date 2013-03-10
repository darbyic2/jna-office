package com.sun.jna.platform.win32.office.outlook;

public class AppointmentCopyOption extends AbstractEnum {
	
	public final static AppointmentCopyOption olPromptUser = new AppointmentCopyOption(0, "olPromptUser");
	public final static AppointmentCopyOption olCreateAppointment = new AppointmentCopyOption(1, "olCreateAppointment");
	public final static AppointmentCopyOption olCopyAsAccept = new AppointmentCopyOption(2, "olCopyAsAccept");

	private AppointmentCopyOption(int val, String name) {
		super((short) val, name);
	}
	
	public static AppointmentCopyOption parse(short val) {
		
		switch(val) {
		
		case 0:
			return olPromptUser;
			
		case 1:
			return olCreateAppointment;
			
		case 2:
			return olCopyAsAccept;
			
		default:
			throw new RuntimeException("AppointmentCopyOption Enum: " + val + " not recognised.");
		}
	}
}
