package com.sun.jna.platform.win32.office.outlook;

public class MeetingStatus extends AbstractEnum {
	
	public final static MeetingStatus	olNonMeeting					= new MeetingStatus(0, "olNonMeeting"); //An Appointment item without attendees has been scheduled. This status can be used to set up holidays on a calendar.
	public final static MeetingStatus	olMeeting						= new MeetingStatus(1, "olMeeting"); //The meeting has been scheduled.
	public final static MeetingStatus	olMeetingReceived				= new MeetingStatus(3, "olMeetingReceived"); //The meeting request has been received.
	public final static MeetingStatus	olMeetingCanceled				= new MeetingStatus(5, "olMeetingCanceled"); //The scheduled meeting has been cancelled.
	public final static MeetingStatus	olMeetingReceivedAndCanceled	= new MeetingStatus(7, "olMeetingReceivedAndCanceled"); //The scheduled meeting has been cancelled but still appears on the user's calendar.
	
	private MeetingStatus(int status, String name) {
		super((short) status, name);
	}

	public static MeetingStatus parse(short status) {
		
		switch(status) {
		
		case 0:
			return olNonMeeting;
			
		case 1:
			return olMeeting;
			
		case 3:
			return olMeetingReceived;
			
		case 5:
			return olMeetingCanceled;
			
		case 7:
			return olMeetingReceivedAndCanceled;
			
		default:
			throw new RuntimeException("MeetingStatus Enum: " + status + " not recognised.");
		}
	}

}
