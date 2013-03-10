package com.sun.jna.platform.win32.office.outlook;

public class FolderType extends AbstractEnum {
	
	public final static FolderType DELETED_ITEMS = new FolderType(3, "DELETED_ITEMS");
	public final static FolderType OUTBOX = new FolderType(4, "OUTBOX");
	public final static FolderType SENTMAIL = new FolderType(5, "SENTMAIL");
	public final static FolderType INBOX = new FolderType(6, "INBOX");
	public final static FolderType CALENDAR = new FolderType(9, "CALENDAR");
	public final static FolderType CONTACTS = new FolderType(10, "CONTACTS");
	public final static FolderType JOURNAL = new FolderType(11, "JOURNAL");
	public final static FolderType NOTES = new FolderType(12, "NOTES");
	public final static FolderType TASKS = new FolderType(13, "TASKS");
	public final static FolderType DRAFTS = new FolderType(16, "DRAFTS");
	public final static FolderType ALL_PUBLIC_FOLDERS = new FolderType(18, "ALL_PUBLIC_FOLDERS");
	public final static FolderType CONFLICTS = new FolderType(19, "CONFLICTS");
	public final static FolderType SYNC_ISSUES = new FolderType(20, "SYNC_ISSUES");
	public final static FolderType LOCAL_FAILURES = new FolderType(21, "LOCAL_FAILURES");
	public final static FolderType SERVER_FAILURES = new FolderType(22, "SERVER_FAILURES");
	public final static FolderType JUNK = new FolderType(23, "JUNK");
	public final static FolderType RSS_FEEDS = new FolderType(25, "RSS_FEEDS");
	public final static FolderType TODO = new FolderType(28, "TODO");
	public final static FolderType MANAGED_EMAIL = new FolderType(29, "MANAGED_EMAIL");
	public final static FolderType SUGGESTED_CONTACTS = new FolderType(30, "SUGGESTED_CONTACTS");

	private FolderType(int typ, String name) {
		super((short) typ, name);
	}
	
	public static FolderType parse(short typValue) {
		
		switch(typValue) {
		case 3:
			return DELETED_ITEMS;
			
		case 4:
			return OUTBOX;
			
		case 5:
			return SENTMAIL;
			
		case 6:
			return INBOX;
			
		case 9:
			return CALENDAR;
			
		case 10:
			return CONTACTS;
			
		case 11:
			return JOURNAL;
			
		case 12:
			return NOTES;
			
		case 13:
			return TASKS;
			
		case 16:
			return DRAFTS;
			
		case 18:
			return ALL_PUBLIC_FOLDERS;
			
		case 19:
			return CONFLICTS;
			
		case 20:
			return SYNC_ISSUES;
			
		case 21:
			return LOCAL_FAILURES;
			
		case 22:
			return SERVER_FAILURES;
			
		case 23:
			return JUNK;
			
		case 25:
			return RSS_FEEDS;
			
		case 28:
			return TODO;
			
		case 29:
			return MANAGED_EMAIL;
			
		case 30:
			return SUGGESTED_CONTACTS;
			
		default:
			return INBOX;
		}
	}
}
