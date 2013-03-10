package com.sun.jna.platform.win32.office.outlook;

public class DefaultFolder extends AbstractEnum {
	
	public final static DefaultFolder	olFolderDeletedItems			= new DefaultFolder(3, "olFolderDeletedItems"); //The Deleted Items folder.
	public final static DefaultFolder	olFolderOutbox					= new DefaultFolder(4, "olFolderOutbox"); //The Outbox folder.
	public final static DefaultFolder	olFolderSentMail				= new DefaultFolder(5, "olFolderSentMail"); //The Sent Mail folder.
	public final static DefaultFolder	olFolderInbox					= new DefaultFolder(6, "olFolderInbox"); //The Inbox folder.
	public final static DefaultFolder	olFolderCalendar				= new DefaultFolder(9, "olFolderCalendar"); //The Calendar folder.
	public final static DefaultFolder	olFolderContacts				= new DefaultFolder(10, "olFolderContacts"); //The Contacts folder.
	public final static DefaultFolder	olFolderJournal					= new DefaultFolder(11, "olFolderJournal"); //The Journal folder.
	public final static DefaultFolder	olFolderNotes					= new DefaultFolder(12, "olFolderNotes"); //The Notes folder.
	public final static DefaultFolder	olFolderTasks					= new DefaultFolder(13, "olFolderTasks"); //The Tasks folder.
	public final static DefaultFolder	olFolderDrafts					= new DefaultFolder(16, "olFolderDrafts"); //The Drafts folder.
	public final static DefaultFolder	olPublicFoldersAllPublicFolders	= new DefaultFolder(18, "olPublicFoldersAllPublicFolders"); //The All Public Folders folder in the Exchange Public Folders store. Only available for an Exchange account.
	public final static DefaultFolder	olFolderConflicts				= new DefaultFolder(19, "olFolderConflicts"); //The Conflicts folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
	public final static DefaultFolder	olFolderSyncIssues				= new DefaultFolder(20, "olFolderSyncIssues"); //The Sync Issues folder. Only available for an Exchange account.
	public final static DefaultFolder	olFolderLocalFailures			= new DefaultFolder(21, "olFolderLocalFailures"); //The Local Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
	public final static DefaultFolder	olFolderServerFailures			= new DefaultFolder(22, "olFolderServerFailures"); //The Server Failures folder (subfolder of the Sync Issues folder). Only available for an Exchange account.
	public final static DefaultFolder	olFolderJunk					= new DefaultFolder(23, "olFolderJunk"); //The Junk E-Mail folder.
	public final static DefaultFolder	olFolderRssFeeds				= new DefaultFolder(25, "olFolderRssFeeds"); //The RSS Feeds folder.
	public final static DefaultFolder	olFolderToDo					= new DefaultFolder(28, "olFolderToDo"); //The To Do folder.
	public final static DefaultFolder	olFolderManagedEmail			= new DefaultFolder(29, "olFolderManagedEmail"); //The top-level folder in the Managed Folders group. For more information on Managed Folders, see the Help in Microsoft Outlook. Only available for an Exchange account.
	public final static DefaultFolder	olFolderSuggestedContacts		= new DefaultFolder(30, "olFolderSuggestedContacts"); //The Suggested Contacts folder.

	private DefaultFolder(int val, String name) {
		super((short) val, name);
	}

	public static DefaultFolder parse(short val) {
		
		switch(val) {
		
		case 3:
			return olFolderDeletedItems;
			
		case 4:
			return olFolderOutbox;
			
		case 5:
			return olFolderSentMail;
			
		case 6:
			return olFolderInbox;
			
		case 9:
			return olFolderCalendar;
			
		case 10:
			return olFolderContacts;
			
		case 11:
			return olFolderJournal;
			
		case 12:
			return olFolderNotes;
			
		case 13:
			return olFolderTasks;
			
		case 16:
			return olFolderDrafts;
			
		case 18:
			return olPublicFoldersAllPublicFolders;
			
		case 19:
			return olFolderConflicts;
			
		case 20:
			return olFolderSyncIssues;
			
		case 21:
			return olFolderLocalFailures;
			
		case 22:
			return olFolderServerFailures;
			
		case 23:
			return olFolderJunk;
			
		case 25:
			return olFolderRssFeeds;
			
		case 28:
			return olFolderToDo;
			
		case 29:
			return olFolderManagedEmail;
			
		case 30:
			return olFolderSuggestedContacts;
			
		default:
			throw new RuntimeException("DefaultFolder Enum: " + val + " not recognised.");
		}
	}
	
}
