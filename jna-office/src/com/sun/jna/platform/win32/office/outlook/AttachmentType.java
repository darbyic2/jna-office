package com.sun.jna.platform.win32.office.outlook;

public class AttachmentType extends AbstractEnum {

	public final static AttachmentType BY_VALUE = new AttachmentType(1, "BY_VALUE");
	public final static AttachmentType BY_REFERENCE = new AttachmentType(4, "BY_REFERENCE");
	public final static AttachmentType EMBEDDED_ITEM = new AttachmentType(5, "EMBEDDED_ITEM");
	public final static AttachmentType OLE = new AttachmentType(6, "OLE");
	
	private AttachmentType(int typ, String name) {
		super((short) typ, name);
	}
	
	public static AttachmentType parse(short typeValue) {
		
		switch(typeValue) {
		case 1:
			return BY_VALUE;
			
		case 4:
			return BY_REFERENCE;
			
		case 5:
			return EMBEDDED_ITEM;
			
		default:
			return OLE;
		}
	}
}
