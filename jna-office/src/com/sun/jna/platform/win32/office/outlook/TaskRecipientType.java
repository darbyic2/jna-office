package com.sun.jna.platform.win32.office.outlook;

public class TaskRecipientType extends RecipientType {

	public final static TaskRecipientType	olUpdate	= new TaskRecipientType(2, "olUpdate");   //	Meeting organizer
	public final static TaskRecipientType	olFinalStatus	= new TaskRecipientType(3, "olFinalStatus");   //	Required attendee
	
	private TaskRecipientType(int typ, String name) {
		super(typ,name);
	}
	
	public static RecipientType parse(int typ) {
		
		switch(typ) {
		
			case 2:
				return olUpdate;
				
			case 3:
				return olFinalStatus;
				
			default:
				return UnknownRecipientType.Unknown;
		}
	}
}
