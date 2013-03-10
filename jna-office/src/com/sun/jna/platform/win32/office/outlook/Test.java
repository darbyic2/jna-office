package com.sun.jna.platform.win32.office.outlook;

import java.text.DateFormat;

public class Test {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		Outlook ol = new Outlook();
		Namespace ns = ol.getSession();
		Folder inbox = ns.getDefaultFolder(FolderType.INBOX);
		Items items = inbox.getItems();
		
		MailItem firstItem = (MailItem) items.getFirst();
		System.out.println("First item subject: " + firstItem.getSubject());
		System.out.println("First item creationTime: " + DateFormat.getDateTimeInstance(DateFormat.SHORT, DateFormat.MEDIUM).format(firstItem.getCreationTime()));
		
		MailItem lastItem = (MailItem) items.getLast();
		System.out.println("Last item subject: " + lastItem.getSubject());
		System.out.println("Last item creationTime: " + DateFormat.getDateTimeInstance(DateFormat.SHORT, DateFormat.MEDIUM).format(lastItem.getCreationTime()));

	}

}
