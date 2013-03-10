package com.sun.jna.platform.win32.office.outlook;

public class ActionCopyLike extends AbstractEnum {

	public final static ActionCopyLike olReply = new ActionCopyLike(0, "olReply");				//Subject will be prefixed with "RE:" and To = Original From
	public final static ActionCopyLike olReplyAll = new ActionCopyLike(1, "olReplyAll");		//Subject will be prefixed with "RE:" and (To = Original From/To) + Cc = Cc
	public final static ActionCopyLike olForward = new ActionCopyLike(2, "olForward");			//Subject will be prefixed with "FW:" and addresses blank
	public final static ActionCopyLike olReplyFolder = new ActionCopyLike(3, "olReplyFolder");
	public final static ActionCopyLike olRespond = new ActionCopyLike(4, "olRespond");
	
	private ActionCopyLike(int val, String name) {
		super((short) val, name);
	}
	
	public static ActionCopyLike parse(short val) {
		
		switch(val) {
		
		case 0:
			return olReply;
			
		case 1:
			return olReplyAll;
			
		case 2:
			return olForward;
			
		case 3:
			return olReplyFolder;
			
		case 4:
			return olRespond;
			
		default:
			throw new RuntimeException("ActionCopyLike Enum: " + val + " not recognised.");
			
		}
	}
}
