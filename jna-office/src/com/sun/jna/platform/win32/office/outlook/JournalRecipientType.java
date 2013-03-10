package com.sun.jna.platform.win32.office.outlook;

public class JournalRecipientType extends RecipientType {
	
	public final static JournalRecipientType	olAssociatedContact	= new JournalRecipientType(1, "olAssociatedContact");   //	The Contact associated with the Journal item.

	private JournalRecipientType(int typ, String name) {
		super(typ, name);
	}
	
	public static RecipientType parse(int typ) {
		
		switch(typ) {
			case 1:
				return olAssociatedContact;
				
			default:
				return UnknownRecipientType.Unknown;
		}
	}
	
}
