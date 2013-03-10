package com.sun.jna.platform.win32.office.outlook;

public class ItemType extends AbstractEnum {

	public final static ItemType MAIL_ITEM = new ItemType(0, "MAIL_ITEM");
	public final static ItemType APPOINTMENT_ITEM = new ItemType(1, "APPOINTMENT_ITEM");
	public final static ItemType CONTACT_ITEM = new ItemType(2, "CONTACT_ITEM");
	public final static ItemType TASK_ITEM = new ItemType(3, "TASK_ITEM");
	public final static ItemType JOURNAL_ITEM = new ItemType(4, "JOURNAL_ITEM");
	public final static ItemType NOTE_ITEM = new ItemType(5, "NOTE_ITEM");
	public final static ItemType POST_ITEM = new ItemType(6, "POST_ITEM");
	public final static ItemType DISTRIBUTION_LIST_ITEM = new ItemType(7, "DISTRIBUTION_LIST_ITEM");
	public final static ItemType MOBILE_ITEM_SMS = new ItemType(11, "MOBILE_ITEM_SMS");
	public final static ItemType MOBILE_ITEM_MMS = new ItemType(12, "MOBILE_ITEM_MMS");
	
	private ItemType(int olType, String name) {
		super((short) olType, name);
	}
	
	public static ItemType parse(short value) {
		switch(value) {
		case 1:
			return APPOINTMENT_ITEM;
			
		case 2:
			return CONTACT_ITEM;
			
		case 3:
			return TASK_ITEM;
			
		case 4:
			return JOURNAL_ITEM;
			
		case 5:
			return NOTE_ITEM;
			
		case 6:
			return POST_ITEM;
			
		case 7:
			return DISTRIBUTION_LIST_ITEM;
			
		case 11:
			return MOBILE_ITEM_SMS;
			
		case 12:
			return MOBILE_ITEM_MMS;
			
		default:
			return MAIL_ITEM;
		}
	}
}
