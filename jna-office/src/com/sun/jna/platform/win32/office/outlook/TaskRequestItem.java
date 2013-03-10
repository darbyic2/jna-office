package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class TaskRequestItem extends BaseTaskItem {
	
	public TaskRequestItem(IDispatch iDisp) {
		super(iDisp);
	}

	public TaskItem getAssociatedTaskItem(boolean addTasktoList) {
		
		return new TaskItem((IDispatch) invoke("", newVariant(addTasktoList)).getValue());
	}
	
}
