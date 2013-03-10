package com.sun.jna.platform.win32.office.outlook;

public class TaskResponse extends AbstractEnum {
	
	public final static TaskResponse olTaskSimple = new TaskResponse(0, "olTaskSimple");
	public final static TaskResponse olTaskAssign = new TaskResponse(1, "olTaskAssign");
	public final static TaskResponse olTaskAccept = new TaskResponse(2, "olTaskAccept");
	public final static TaskResponse olTaskDecline = new TaskResponse(3, "olTaskDecline");
	
	private TaskResponse(int response, String name) {
		super((short) response, name);
	}

	public static TaskResponse parse(short response) {
		
		switch(response) {
		
		case 0:
			return olTaskSimple;
			
		case 1:
			return olTaskAssign;
			
		case 2:
			return olTaskAccept;
			
		case 3:
			return olTaskDecline;
		
		default:
			throw new RuntimeException("TaskResponse Enum: " + response + " not recognised.");
		}
	}
}
