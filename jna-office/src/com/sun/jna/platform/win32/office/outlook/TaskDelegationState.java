package com.sun.jna.platform.win32.office.outlook;

public class TaskDelegationState extends AbstractEnum {
	
	public final static TaskDelegationState	olTaskNotDelegated			= new TaskDelegationState(0, "olTaskNotDelegated"); //The task has not been delegated.
	public final static TaskDelegationState	olTaskDelegationUnknown		= new TaskDelegationState(1, "olTaskDelegationUnknown"); //The delegate response to the task is unknown.
	public final static TaskDelegationState	olTaskDelegationAccepted	= new TaskDelegationState(2, "olTaskDelegationAccepted"); //The delegate accepted the task.
	public final static TaskDelegationState	olTaskDelegationDeclined	= new TaskDelegationState(3, "olTaskDelegationDeclined"); //The delegate declined the task.
	
	private TaskDelegationState(int state, String name) {
		super((short) state, name);
	}

	public static TaskDelegationState parse(short state) {
		
		switch(state) {
		
		case 0:
			return olTaskNotDelegated;
		
		case 1:
			return olTaskDelegationUnknown;
		
		case 2:
			return olTaskDelegationAccepted;
		
		case 3:
			return olTaskDelegationDeclined;
		
		default:
			return olTaskDelegationUnknown;
		}
	}

}
