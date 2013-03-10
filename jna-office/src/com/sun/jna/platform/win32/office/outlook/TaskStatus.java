package com.sun.jna.platform.win32.office.outlook;

public class TaskStatus extends AbstractEnum {
	
	public final static TaskStatus	olTaskNotStarted	= new TaskStatus(0, "olTaskNotStarted"); //The task has not yet started.
	public final static TaskStatus	olTaskInProgress	= new TaskStatus(1, "olTaskInProgress"); //The task is in progress.
	public final static TaskStatus	olTaskComplete		= new TaskStatus(2, "olTaskComplete"); //The task is complete.
	public final static TaskStatus	olTaskWaiting		= new TaskStatus(3, "olTaskWaiting"); //The task is waiting on someone else.
	public final static TaskStatus	olTaskDeferred		= new TaskStatus(4, "olTaskDeferred"); //The task is deferred.
	
	private TaskStatus(int status, String name) {
		super((short) status, name);
	}

	public static TaskStatus parse(short status) {
		
		switch(status) {
		
		case 0:
			return olTaskNotStarted;
			
		case 1:
			return olTaskInProgress;
			
		case 2:
			return olTaskComplete;
			
		case 3:
			return olTaskWaiting;
			
		case 4:
			return olTaskDeferred;
		
		default:
			throw new RuntimeException("TaskStatus Enum: " + status + " not recognised.");
		}
	}

}
