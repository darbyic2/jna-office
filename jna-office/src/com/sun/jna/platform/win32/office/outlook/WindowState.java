package com.sun.jna.platform.win32.office.outlook;

public class WindowState extends AbstractEnum {

	public final static WindowState	olMaximized	= new WindowState(0, "olMaximized"); //The window is maximized.
	public final static WindowState	olMinimized	= new WindowState(1, "olMinimized"); //The window is minimized.
	public final static WindowState	olNormalWindow	= new WindowState(2, "olNormalWindow"); //The window is in the normal state (not minimized or maximized).

	private WindowState(int state, String name) {
		super((short) state, name);
	}
	
	public static WindowState parse(short state) {
		
		switch(state) {
		
		case 0:
			return olMaximized;
			
		case 1:
			return olMinimized;
			
		case 2:
			return olNormalWindow;
			
		default:
			throw new RuntimeException("WindowState Enum: " + state + " not recognised.");
		}
	}
}
