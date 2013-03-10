package com.sun.jna.platform.win32.office.outlook;

public class InspectorCloseOption extends AbstractEnum {

	public final static InspectorCloseOption SAVE = new InspectorCloseOption(0, "SAVE");
	public final static InspectorCloseOption DISCARD = new InspectorCloseOption(1, "DISCARD");
	public final static InspectorCloseOption PROMPT_FOR_SAVE = new InspectorCloseOption(2, "PROMPT_FOR_SAVE");
	
	private InspectorCloseOption(int value, String name) {
		super((short) value, name);
	}
	
	public static InspectorCloseOption parse(short value) {
		
		switch(value) {
		case 0:
			return SAVE;
			
		case 1:
			return DISCARD;
			
		default:
			return PROMPT_FOR_SAVE;
		}
	}
}
