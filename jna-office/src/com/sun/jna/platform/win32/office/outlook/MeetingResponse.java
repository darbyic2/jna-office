package com.sun.jna.platform.win32.office.outlook;

public class MeetingResponse extends AbstractEnum {
	
	public final static MeetingResponse olMeetingTentative = new MeetingResponse(2, "olMeetingTentative");
	public final static MeetingResponse olMeetingAccepted = new MeetingResponse(3, "olMeetingAccepted");
	public final static MeetingResponse olMeetingDeclined = new MeetingResponse(4, "olMeetingDeclined");
	
	private MeetingResponse(int val, String name) {
		super((short) val, name);
	}
	
	public static MeetingResponse parse(short val) {
		
		switch(val) {
		
		case 2:
			return olMeetingTentative;
			
		case 3:
			return olMeetingAccepted;
			
		case 4:
			return olMeetingDeclined;
			
		default:
			return olMeetingTentative;
		}
	}
}
