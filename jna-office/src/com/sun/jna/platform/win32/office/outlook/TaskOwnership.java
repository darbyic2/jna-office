package com.sun.jna.platform.win32.office.outlook;

public class TaskOwnership extends AbstractEnum {
	
	public final static TaskOwnership olNewTask = new TaskOwnership(0,"olNewTask");
	public final static TaskOwnership olDelegatedTask = new TaskOwnership(1,"olDelegatedTask");
	public final static TaskOwnership olOwnTask = new TaskOwnership(2,"olOwnTask");
	
	private TaskOwnership(int val, String name) {
		super((short) val, name);
	}

	public static TaskOwnership parse(short val) {
		
		switch(val) {
		
		case 0:
			return olNewTask;
			
		case 1:
			return olDelegatedTask;
			
		case 2:
			return olOwnTask;
			
		default:
			throw new RuntimeException("TaskOwnership Enum: " + val + " not recognised.");
		}
	}

}
