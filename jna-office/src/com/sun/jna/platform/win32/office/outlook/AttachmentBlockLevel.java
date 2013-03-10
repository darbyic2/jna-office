package com.sun.jna.platform.win32.office.outlook;

public class AttachmentBlockLevel extends AbstractEnum {

	public final static AttachmentBlockLevel NONE = new AttachmentBlockLevel(0, "NONE");
	public final static AttachmentBlockLevel OPEN = new AttachmentBlockLevel(1, "OPEN");
	
	public static AttachmentBlockLevel parse(short blockingLevel) {
		
		if (blockingLevel > 0) {
			return OPEN;
			
		} else {
			return NONE;
		}
	}
	
	private AttachmentBlockLevel(int value, String name) {
		super((short) value, name);
	}
	
}
