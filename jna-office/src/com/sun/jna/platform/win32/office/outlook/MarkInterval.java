package com.sun.jna.platform.win32.office.outlook;

public class MarkInterval extends AbstractEnum {
	
	public final static MarkInterval	olMarkToday		= new MarkInterval(0, "olMarkToday");		//Mark the task due today.
	public final static MarkInterval	olMarkTomorrow	= new MarkInterval(1, "olMarkTomorrow");	//Mark the task due tomorrow.
	public final static MarkInterval	olMarkThisWeek	= new MarkInterval(2, "olMarkThisWeek");	//Mark the task due this week.
	public final static MarkInterval	olMarkNextWeek	= new MarkInterval(3, "olMarkNextWeek");	//Mark the task due next week.
	public final static MarkInterval	olMarkNoDate	= new MarkInterval(4, "olMarkNoDate");		//Mark the task due with no date.
	public final static MarkInterval	olMarkComplete	= new MarkInterval(5, "olMarkComplete");	//Mark the task as complete.
	
	private MarkInterval(int interval, String name) {
		super((short) interval, name);
	}
	
	public static MarkInterval parse(short interval) {
		
		switch(interval) {
		
		case 0:
			return olMarkToday;
			
		case 1:
			return olMarkTomorrow;
			
		case 2:
			return olMarkThisWeek;
			
		case 3:
			return olMarkNextWeek;
			
		case 5:
			return olMarkComplete;
		
		case 4:
		default:
			return olMarkNoDate;
		}
	}

}
