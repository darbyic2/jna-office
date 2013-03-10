package com.sun.jna.platform.win32.office.outlook;

public class MailRecipientType extends RecipientType {
	
	public final static MailRecipientType	olOriginator	= new MailRecipientType(0, "olOriginator");   //	Originator (sender) of the Item.
	public final static MailRecipientType	olTo			= new MailRecipientType(1, "olTo");   //	The recipient is specified in the To property of the Item.
	public final static MailRecipientType	olCC			= new MailRecipientType(2, "olCC");   //	The recipient is specified in the CC property of the Item.
	public final static MailRecipientType	olBCC			= new MailRecipientType(3, "olBCC");   //	The recipient is specified in the BCC property of the Item.

	private MailRecipientType(int typ, String name) {
		super(typ, name);
	}
	
	public static RecipientType parse(int typ) {
		switch(typ) {
			case 0:
				return olOriginator;
				
			case 1:
				return olTo;
				
			case 2:
				return olCC;
				
			case 3:
				return olBCC;
				
			default:
				return UnknownRecipientType.Unknown;
		}
	}
}
