package com.sun.jna.platform.win32.office.outlook;

public class ActionShowOn extends AbstractEnum {
	
	public final static ActionShowOn	olDontShow		= new ActionShowOn(0, "olDontShow");  //	Indicates that the action will not be displayed on the menu or toolbar.
	public final static ActionShowOn	olMenu			= new ActionShowOn(1, "olMenu");  //	Indicates that the action will be displayed as an available action on the menu.
	public final static ActionShowOn	olMenuAndToolbar = new ActionShowOn(2, "olMenuAndToolbar");  //	Indicates that the action will be displayed as an available action on the menu and the toolbar.

	private ActionShowOn(int typ, String name) {
		super((short) typ, name);
	}

	public static ActionShowOn parse(short style) {
		
		switch(style) {
		
		case 0:
			return olDontShow;
			
		case 1:
			return olMenu;
		
		case 2:
			return olMenuAndToolbar;
			
		default:
			throw new RuntimeException("ActionShowOn Enum: " + style + " not recognised.");
		}
	}
}
