/* Copyright (c) 2013 Ian Darby, All Rights Reserved
 * 
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 * 
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.  
 */

package com.sun.jna.platform.win32.office.outlook;
/*
 */

/**
 * Would really like this to have been a type-safe enum sort of class. However
 * it has to be just a set of public int constants so that it can be used in
 * case statements.
 * <p>
 * All objects within the Outlook object model support a property
 * &quot;Class&quot; which within this set of wrappers has been renamed
 * {@link BaseOutlookObject#getClassEnumValue()}. Both of these return an
 * integer constant value defined within this class.
 * </p>
 * 
 * @author Ian Darby
 * 
 */
public final class ClassEnum {

	/**
	 * An {@link Account} object.
	 */
	public final static int	olAccount	=	105	;  
	
	/**
	 * An {@link AccountRuleCondition} object.
	 */
	public final static int	olAccountRuleCondition	=	135	; 
	
	/**
	 * An {@link Accounts} object.
	 */
	public final static int	olAccounts	=	106	;  
	
	/**
	 * An {@link Action} object.
	 */
	public final static int	olAction	=	32	;  
	
	/**
	 * An {@link Actions} object.
	 */
	public final static int	olActions	=	33	;  
	
	/**
	 * An {@link AddressEntries} object.
	 */
	public final static int	olAddressEntries	=	21	;  
	
	/**
	 * An {@link AddressEntry} object.
	 */
	public final static int	olAddressEntry	=	8	;  
	
	/**
	 * An {@link AddressList} object.
	 */
	public final static int	olAddressList	=	7	;  
	
	/**
	 * An {@link AddressLists} object.
	 */
	public final static int	olAddressLists	=	20	;  
	
	/**
	 * An {@link AddressRuleCondition} object.
	 */
	public final static int	olAddressRuleCondition	=	170	;  
	
	/**
	 * An {@link Outlook} application object.
	 */
	public final static int	olApplication	=	0	;  
	
	/**
	 * An {@link AppointmentItem} object.
	 */
	public final static int	olAppointment	=	26	;  
	
	/**
	 * An {@link AssignToCategoryRuleAction} object.
	 */
	public final static int	olAssignToCategoryRuleAction	=	122	;  
	
	/**
	 * An {@link Attachment} object.
	 */
	public final static int	olAttachment	=	5	;  
	
	/**
	 * An {@link Attachments} object.
	 */
	public final static int	olAttachments	=	18	;  
	
	/**
	 * An {@link AttachmentSelection} object.
	 */
	public final static int	olAttachmentSelection	=	169	;  
	
	/**
	 * An {@link AutoFormatRule} object.
	 */
	public final static int	olAutoFormatRule	=	147	;  
	
	/**
	 * An {@link AutoFormatRules} object.
	 */
	public final static int	olAutoFormatRules	=	148	;  
	
	/**
	 * A {@link CalendarModule} object.
	 */
	public final static int	olCalendarModule	=	159	;  
	
	/**
	 * A {@link CalendarSharing} object.
	 */
	public final static int	olCalendarSharing	=	151	;  
	
	/**
	 * A {@link Categories} object.
	 */
	public final static int	olCategories	=	153	;  
	
	/**
	 * A {@link Category} object.
	 */
	public final static int	olCategory	=	152	;  
	
	/**
	 * A {@link CategoryRuleCondition} object.
	 */
	public final static int	olCategoryRuleCondition	=	130	;  
	
	/**
	 * A {@link BusinessCardView} object.
	 */
	public final static int	olClassBusinessCardView	=	168	;  
	
	/**
	 * A {@link CalendarView} object.
	 */
	public final static int	olClassCalendarView	=	139	;  
	
	/**
	 * A {@link CardView} object.
	 */
	public final static int	olClassCardView	=	138	;  
	
	/**
	 * An {@link IconView} object.
	 */
	public final static int	olClassIconView	=	137	;  
	
	/**
	 * A {@link NavigationPane} object.
	 */
	public final static int	olClassNavigationPane	=	155	;  
	
	/**
	 * A {@link TableView} object.
	 */
	public final static int	olClassTableView	=	136	;  
	
	/**
	 * A {@link TimelineView} object.
	 */
	public final static int	olClassTimeLineView	=	140	;  
	
	/**
	 * A {@link TimeZone} object.
	 */
	public final static int	olClassTimeZone	=	174	;  
	
	/**
	 * A {@link TimeZones} object.
	 */
	public final static int	olClassTimeZones	=	175	;  
	
	/**
	 * A {@link Column} object.
	 */
	public final static int	olColumn	=	154	;  
	
	/**
	 * A {@link ColumnFormat} object.
	 */
	public final static int	olColumnFormat	=	149	;  
	
	/**
	 * A {@link Columns} object.
	 */
	public final static int	olColumns	=	150	;  
	
	/**
	 * A {@link Conflict} object.
	 */
	public final static int	olConflict	=	102	;  
	
	/**
	 * A {@link Conflicts} object.
	 */
	public final static int	olConflicts	=	103	;  
	
	/**
	 * A {@link ContactItem} object.
	 */
	public final static int	olContact	=	40	;  
	
	/**
	 * A {@link ContactsModule} object.
	 */
	public final static int	olContactsModule	=	160	;  
	
	/**
	 * A {@link Conversation} object.
	 */
	public final static int	olConversation	=	178	;  
	
	/**
	 * A {@link ConversationHeader} object.
	 */
	public final static int	olConversationHeader	=	182	;  
	
	/**
	 * An {@link ExchangeDistributionList} object.
	 */
	public final static int	olDistributionList	=	69	;  
	
	/**
	 * A {@link DocumentItem} object.
	 */
	public final static int	olDocument	=	41	;  
	
	/**
	 * An {@link Exception} object.
	 */
	public final static int	olException	=	30	;  
	
	/**
	 * An {@link Exceptions} object.
	 */
	public final static int	olExceptions	=	29	;  
	
	/**
	 * An {@link ExchangeDistributionList} object.
	 */
	public final static int	olExchangeDistributionList	=	111	;  
	
	/**
	 * An {@link ExchangeUser} object.
	 */
	public final static int	olExchangeUser	=	110	;  
	
	/**
	 * An {@link Explorer} object.
	 */
	public final static int	olExplorer	=	34	;  
	
	/**
	 * An {@link Explorers} object.
	 */
	public final static int	olExplorers	=	60	;  
	
	/**
	 * A {@link Folder} object.
	 */
	public final static int	olFolder	=	2	;  
	
	/**
	 * A {@link Folders} object.
	 */
	public final static int	olFolders	=	15	;  
	
	/**
	 * A {@link UserDefinedProperties} object.
	 */
	public final static int	olFolderUserProperties	=	172	;  
	
	/**
	 * A {@link UserDefinedProperty} object.
	 */
	public final static int	olFolderUserProperty	=	171	;  
	
	/**
	 * A {@link FormDescription} object.
	 */
	public final static int	olFormDescription	=	37	;  
	
	/**
	 * A {@link FormNameRuleCondition} object.
	 */
	public final static int	olFormNameRuleCondition	=	131	;  
	
	/**
	 * A {@link FormRegion} object.
	 */
	public final static int	olFormRegion	=	129	;  
	
	/**
	 * A {@link FromRssFeedRuleCondition} object.
	 */
	public final static int	olFromRssFeedRuleCondition	=	173	;  
	
	/**
	 * A {@link ToOrFromRuleCondition} object.
	 */
	public final static int	olFromRuleCondition	=	132	;  
	
	/**
	 * An {@link ImportanceRuleCondition} object.
	 */
	public final static int	olImportanceRuleCondition	=	128	;  
	
	/**
	 * An {@link Inspector} object.
	 */
	public final static int	olInspector	=	35	;  
	
	/**
	 * An {@link Inspectors} object.
	 */
	public final static int	olInspectors	=	61	;  
	
	/**
	 * An {@link ItemProperties} object.
	 */
	public final static int	olItemProperties	=	98	;  
	
	/**
	 * An {@link ItemProperty} object.
	 */
	public final static int	olItemProperty	=	99	;  
	
	/**
	 * An {@link Items} object.
	 */
	public final static int	olItems	=	16	;  
	
	/**
	 * A {@link JournalItem} object.
	 */
	public final static int	olJournal	=	42	;  
	
	/**
	 * A {@link JournalModule} object.
	 */
	public final static int	olJournalModule	=	162	;  
	
	/**
	 * A {@link Link} object.
	 */
	public final static int	olLink	=	75	;  
	
	/**
	 * A {@link Links} object.
	 */
	public final static int	olLinks	=	76	;  
	
	/**
	 * A {@link MailItem} object.
	 */
	public final static int	olMail	=	43	;  
	
	/**
	 * A {@link MailModule} object.
	 */
	public final static int	olMailModule	=	158	;  
	
	/**
	 * A {@link MarkAsTaskRuleAction} object.
	 */
	public final static int	olMarkAsTaskRuleAction	=	124	;  
	
	/**
	 * A {@link MeetingItem} object that is a meeting cancellation notice.
	 */
	public final static int	olMeetingCancellation	=	54	;  
	
	/**
	 * A {@link MeetingItem} object that is a notice about forwarding the meeting request.
	 */
	public final static int	olMeetingForwardNotification	=	181	;  
	
	/**
	 * A {@link MeetingItem} object that is a meeting request.
	 */
	public final static int	olMeetingRequest	=	53	;  
	
	/**
	 * A {@link MeetingItem} object that is a refusal of a meeting request.
	 */
	public final static int	olMeetingResponseNegative	=	55	;  
	
	/**
	 * A {@link MeetingItem} object that is an acceptance of a meeting request.
	 */
	public final static int	olMeetingResponsePositive	=	56	;  
	
	/**
	 * A {@link MeetingItem} object that is a tentative acceptance of a meeting request.
	 */
	public final static int	olMeetingResponseTentative	=	57	;  
	
	/**
	 * A {@link MobileItem} object that is a text message item or a multimedia message item.
	 */
	public final static int	olMobile	=	176	;  
	
	/**
	 * A {@link MoveOrCopyRuleAction} object.
	 */
	public final static int	olMoveOrCopyRuleAction	=	118	;  
	
	/**
	 * A {@link Namespace} object.
	 */
	public final static int	olNamespace	=	1	;  
	
	/**
	 * A {@link NavigationFolder} object.
	 */
	public final static int	olNavigationFolder	=	167	;  
	
	/**
	 * A {@link NavigationFolders} object.
	 */
	public final static int	olNavigationFolders	=	166	;  
	
	/**
	 * A {@link NavigationGroup} object.
	 */
	public final static int	olNavigationGroup	=	165	;  
	
	/**
	 * A {@link NavigationGroups} object.
	 */
	public final static int	olNavigationGroups	=	164	;  
	
	/**
	 * A {@link NavigationModule} object.
	 */
	public final static int	olNavigationModule	=	157	;  
	
	/**
	 * A {@link NavigationModules} object.
	 */
	public final static int	olNavigationModules	=	156	;  
	
	/**
	 * A {@link NewItemAlertRuleAction} object.
	 */
	public final static int	olNewItemAlertRuleAction	=	125	;  
	
	/**
	 * A {@link NoteItem} object.
	 */
	public final static int	olNote	=	44	;  
	
	/**
	 * A {@link NotesModule} object.
	 */
	public final static int	olNotesModule	=	163	;  
	
	/**
	 * An {@link OrderField} object.
	 */
	public final static int	olOrderField	=	144	;  
	
	/**
	 * An {@link OrderFields} object.
	 */
	public final static int	olOrderFields	=	145	;  
	
	/**
	 * An {@link OutlookBarGroup} object.
	 */
	public final static int	olOutlookBarGroup	=	66	;  
	
	/**
	 * An {@link OutlookBarGroups} object.
	 */
	public final static int	olOutlookBarGroups	=	65	;  
	
	/**
	 * An {@link OutlookBarPane} object.
	 */
	public final static int	olOutlookBarPane	=	63	;  
	
	/**
	 * An {@link OutlookBarShortcut} object.
	 */
	public final static int	olOutlookBarShortcut	=	68	;  
	
	/**
	 * An {@link OutlookBarShortcuts} object.
	 */
	public final static int	olOutlookBarShortcuts	=	67	;  
	
	/**
	 * An {@link OutlookBarStorage} object.
	 */
	public final static int	olOutlookBarStorage	=	64	;  
	
	/**
	 * An {@link AccountSelector} object.
	 */
	public final static int	olOutspace	=	180	;  
	
	/**
	 * A {@link Pages} object.
	 */
	public final static int	olPages	=	36	;  
	
	/**
	 * A {@link Panes} object.
	 */
	public final static int	olPanes	=	62	;  
	
	/**
	 * A {@link PlaySoundRuleAction} object.
	 */
	public final static int	olPlaySoundRuleAction	=	123	;  
	
	/**
	 * A {@link PostItem} object.
	 */
	public final static int	olPost	=	45	;  
	
	/**
	 * A {@link PropertyAccessor} object.
	 */
	public final static int	olPropertyAccessor	=	112	;  
	
	/**
	 * A {@link PropertyPages} object.
	 */
	public final static int	olPropertyPages	=	71	;  
	
	/**
	 * A {@link PropertyPageSite} object.
	 */
	public final static int	olPropertyPageSite	=	70	;  
	
	/**
	 * A {@link Recipient} object.
	 */
	public final static int	olRecipient	=	4	;  
	
	/**
	 * A {@link Recipients} object.
	 */
	public final static int	olRecipients	=	17	;  
	
	/**
	 * A {@link RecurrencePattern} object.
	 */
	public final static int	olRecurrencePattern	=	28	;  
	
	/**
	 * A {@link Reminder} object.
	 */
	public final static int	olReminder	=	101	;  
	
	/**
	 * A {@link Reminders} object.
	 */
	public final static int	olReminders	=	100	;  
	
	/**
	 * A {@link RemoteItem} object.
	 */
	public final static int	olRemote	=	47	;  
	
	/**
	 * A {@link ReportItem} object.
	 */
	public final static int	olReport	=	46	;  
	
	/**
	 * A {@link Results} object.
	 */
	public final static int	olResults	=	78	;  
	
	/**
	 * A {@link Row} object.
	 */
	public final static int	olRow	=	121	;  
	
	/**
	 * A {@link Rule} object.
	 */
	public final static int	olRule	=	115	;  
	
	/**
	 * A {@link RuleAction} object.
	 */
	public final static int	olRuleAction	=	117	;  
	
	/**
	 * A {@link RuleActions} object.
	 */
	public final static int	olRuleActions	=	116	;  
	
	/**
	 * A {@link RuleCondition} object.
	 */
	public final static int	olRuleCondition	=	127	;  
	
	/**
	 * A {@link RuleConditions} object.
	 */
	public final static int	olRuleConditions	=	126	;  
	
	/**
	 * A {@link Rules} object.
	 */
	public final static int	olRules	=	114	;  
	
	/**
	 * A {@link Search} object.
	 */
	public final static int	olSearch	=	77	;  
	
	/**
	 * A {@link Selection} object.
	 */
	public final static int	olSelection	=	74	;  
	
	/**
	 * A {@link SelectNamesDialog} object.
	 */
	public final static int	olSelectNamesDialog	=	109	;  
	
	/**
	 * A {@link SenderInAddressListRuleCondition} object.
	 */
	public final static int	olSenderInAddressListRuleCondition	=	133	;  
	
	/**
	 * A {@link SendRuleAction} object.
	 */
	public final static int	olSendRuleAction	=	119	;  
	
	/**
	 * A {@link SharingItem} object.
	 */
	public final static int	olSharing	=	104	;  
	
	/**
	 * A {@link SimpleItems} object.
	 */
	public final static int	olSimpleItems	=	179	;  
	
	/**
	 * A {@link SolutionsModule} object.
	 */
	public final static int	olSolutionsModule	=	177	;  
	
	/**
	 * A {@link StorageItem} object.
	 */
	public final static int	olStorageItem	=	113	;  
	
	/**
	 * A {@link Store} object.
	 */
	public final static int	olStore	=	107	;  
	
	/**
	 * A {@link Stores} object.
	 */
	public final static int	olStores	=	108	;  
	
	/**
	 * A {@link SyncObject} object.
	 */
	public final static int	olSyncObject	=	72	;  
	
	/**
	 * A {@link SyncObjects} object.
	 */
	public final static int	olSyncObjects	=	73	;  
	
	/**
	 * A {@link Table} object.
	 */
	public final static int	olTable	=	120	;  
	
	/**
	 * A {@link TaskItem} object.
	 */
	public final static int	olTask	=	48	;  
	
	/**
	 * A {@link TaskRequestItem} object.
	 */
	public final static int	olTaskRequest	=	49	;  
	
	/**
	 * A {@link TaskRequestAcceptItem} object.
	 */
	public final static int	olTaskRequestAccept	=	51	;  
	
	/**
	 * A {@link TaskRequestDeclineItem} object.
	 */
	public final static int	olTaskRequestDecline	=	52	;  
	
	/**
	 * A {@link TaskRequestUpdateItem} object.
	 */
	public final static int	olTaskRequestUpdate	=	50	;  
	
	/**
	 * A {@link TasksModule} object.
	 */
	public final static int	olTasksModule	=	161	;  
	
	/**
	 * A {@link TextRuleCondition} object.
	 */
	public final static int	olTextRuleCondition	=	134	;  
	
	/**
	 * A {@link UserDefinedProperties} object.
	 */
	public final static int	olUserDefinedProperties	=	172	;  
	
	/**
	 * A {@link UserDefinedProperty} object.
	 */
	public final static int	olUserDefinedProperty	=	171	;  
	
	/**
	 * A {@link UserProperties} object.
	 */
	public final static int	olUserProperties	=	38	;  
	
	/**
	 * A {@link UserProperty} object.
	 */
	public final static int	olUserProperty	=	39	;  
	
	/**
	 * A {@link View} object.
	 */
	public final static int	olView	=	80	;  
	
	/**
	 * A {@link ViewField} object.
	 */
	public final static int	olViewField	=	142	;  
	
	/**
	 * A {@link ViewFields} object.
	 */
	public final static int	olViewFields	=	141	;  
	
	/**
	 * A {@link ViewFont} object.
	 */
	public final static int	olViewFont	=	146	;  
	
	/**
	 * A {@link Views} object.
	 */
	public final static int	olViews	=	79	;  
	
}
