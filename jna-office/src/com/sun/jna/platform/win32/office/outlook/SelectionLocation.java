package com.sun.jna.platform.win32.office.outlook;

public class SelectionLocation extends AbstractEnum {

	public final static SelectionLocation	olViewList	= new SelectionLocation(0, "olViewList"); //The selection is in a list of items in an explorer.
	public final static SelectionLocation	olToDoBarTaskList	= new SelectionLocation(1, "olToDoBarTaskList"); //The selection is in the list of tasks in the To-Do Bar.
	public final static SelectionLocation	olToDoBarAppointmentList	= new SelectionLocation(2, "olToDoBarAppointmentList"); //The selection is in the list of appointments in the To-Do Bar.
	public final static SelectionLocation	olDailyTaskList	= new SelectionLocation(3, "olDailyTaskList"); //The selection is in the daily Tasks list in the calendar view.
	public final static SelectionLocation	olAttachmentWell	= new SelectionLocation(4, "olAttachmentWell"); //The selection is an attachment of an item in the Reading Pane or inspector.

	private SelectionLocation(int val, String name) {
		super((short) val, name);
	}
	
	public static SelectionLocation parse(short loc) {
		
		switch(loc) {
		
		case 0:
			return olViewList;
			
		case 1:
			return olToDoBarTaskList;
			
		case 2:
			return olToDoBarAppointmentList;
			
		case 3:
			return olDailyTaskList;
			
		case 4:
			return olAttachmentWell;
			
		default:
			throw new RuntimeException("SelectionLocation Enum: " + loc + " not recognised.");
		}
	}
}
