package com.sun.jna.platform.win32.office.outlook;

public class MeetingRecipientType extends RecipientType {
	
	public final static MeetingRecipientType	olOrganizer	= new MeetingRecipientType(	0, "olOrganizer");   //	Meeting organizer
	public final static MeetingRecipientType	olRequired	= new MeetingRecipientType(	1, "olRequired");   //	Required attendee
	public final static MeetingRecipientType	olOptional	= new MeetingRecipientType(	2, "olOptional");   //	Optional attendee
	public final static MeetingRecipientType	olResource	= new MeetingRecipientType(	3, "olResource");   //	A resource such as a conference room

	private MeetingRecipientType(int typ, String name) {
		super(typ,name);
	}
	
	public static RecipientType parse(int typ) {
		
		switch(typ) {
			case 0:
				return olOrganizer;
				
			case 1:
				return olRequired;
				
			case 2:
				return olOptional;
				
			case 3:
				return olResource;
				
			default:
				return UnknownRecipientType.Unknown;
		}
	}
}
