package com.sun.jna.platform.win32.office.outlook;

public class ResponseStatus extends AbstractEnum {

	public final static ResponseStatus	olResponseNone			= new ResponseStatus(0	, "olResponseNone");   	     //	The appointment is a simple appointment and does not require a response.
	public final static ResponseStatus	olResponseOrganized		= new ResponseStatus(1	, "olResponseOrganized");    //	The AppointmentItem is on the Organizer's calendar or the recipient is the Organizer of the meeting.
	public final static ResponseStatus	olResponseTentative		= new ResponseStatus(2	, "olResponseTentative");    //	Meeting tentatively accepted.
	public final static ResponseStatus	olResponseAccepted		= new ResponseStatus(3	, "olResponseAccepted");     //	Meeting accepted.
	public final static ResponseStatus	olResponseDeclined		= new ResponseStatus(4	, "olResponseDeclined");     //	Meeting declined.
	public final static ResponseStatus	olResponseNotResponded	= new ResponseStatus(5	, "olResponseNotResponded"); //	Recipient has not responded.
	
	private ResponseStatus(int status, String name) {
		super((short) status, name);
	}
	
	public static ResponseStatus parse(short status) {
		
		switch(status) {
			case 1:
				return olResponseOrganized;
				
			case 2:
				return olResponseTentative;
				
			case 3:
				return olResponseAccepted;
				
			case 4:
				return olResponseDeclined;
				
			case 5:
				return olResponseNotResponded;
				
			default:
				return olResponseNone;
		}
	}
}
