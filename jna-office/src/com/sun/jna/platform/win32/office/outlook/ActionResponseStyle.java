package com.sun.jna.platform.win32.office.outlook;

public class ActionResponseStyle extends AbstractEnum {
	
	public final static ActionResponseStyle	olOpen		= new ActionResponseStyle(0	, "olOpen");  //	Indicates that a form will be opened.
	public final static ActionResponseStyle	olSend		= new ActionResponseStyle(1	, "olSend");  //	Indicates that the form will be sent immediately.
	public final static ActionResponseStyle	olPrompt	= new ActionResponseStyle(2	, "olPrompt");  //	Indicates that the user will be prompted to open or send the form.

	public ActionResponseStyle(int typ, String name) {
		super((short) typ, name);
	}

	public static ActionResponseStyle parse(short style) {
		switch(style) {
		
		case 0:
			return olOpen;
			
		case 1:
			return olSend;
		
		case 2:
			return olPrompt;
			
		default:
			return olPrompt;
		}
	}
}
