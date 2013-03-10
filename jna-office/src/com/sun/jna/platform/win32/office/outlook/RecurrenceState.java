package com.sun.jna.platform.win32.office.outlook;

public class RecurrenceState extends AbstractEnum {
	
	public final static RecurrenceState	olApptNotRecurring	= new RecurrenceState(0, "olApptNotRecurring"); //The appointment is not a recurring appointment.
	public final static RecurrenceState	olApptMaster		= new RecurrenceState(1, "olApptMaster"); //The appointment is a master appointment.
	public final static RecurrenceState	olApptOccurrence	= new RecurrenceState(2, "olApptOccurrence"); //The appointment is an occurrence of a recurring appointment defined by a master appointment.
	public final static RecurrenceState	olApptException		= new RecurrenceState(3, "olApptException"); //The appointment is an exception to a recurrence pattern defined by a master appointment.

	private RecurrenceState(int state, String name) {
		super((short) state, name);
	}

	public static RecurrenceState parse(short state) {
		
		switch(state) {
		
		case 0:
			return olApptNotRecurring;
			
		case 1:
			return olApptMaster;
			
		case 2:
			return olApptOccurrence;
			
		case 3:
			return olApptException;
			
		default:
			throw new RuntimeException("RecurrenceState Enum: " + state + " not recognised.");
		}
	}
}
