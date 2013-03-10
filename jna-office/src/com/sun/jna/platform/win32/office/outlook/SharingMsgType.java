package com.sun.jna.platform.win32.office.outlook;

public class SharingMsgType extends AbstractEnum {
	
	public final static SharingMsgType	olSharingMsgTypeUnknown				= new SharingMsgType(0, "olSharingMsgTypeUnknown"); //Represents an unknown type of sharing message.
	public final static SharingMsgType	olSharingMsgTypeRequest				= new SharingMsgType(1, "olSharingMsgTypeRequest"); //Represents a sharing request.
	public final static SharingMsgType	olSharingMsgTypeInvite				= new SharingMsgType(2, "olSharingMsgTypeInvite"); //Represents a sharing invitation.
	public final static SharingMsgType	olSharingMsgTypeInviteAndRequest	= new SharingMsgType(3, "olSharingMsgTypeInviteAndRequest"); //Represents both a sharing invitation and a sharing request.
	public final static SharingMsgType	olSharingMsgTypeResponseAllow		= new SharingMsgType(4, "olSharingMsgTypeResponseAllow"); //Represents a sharing response, which indicates that a sharing request or sharing invitation has been allowed.
	public final static SharingMsgType	olSharingMsgTypeResponseDeny		= new SharingMsgType(5, "olSharingMsgTypeResponseDeny"); //Represents a sharing response, which indicates that a sharing request or sharing invitation has been denied.

	private SharingMsgType(int typ, String name) {
		super((short) typ, name);
	}

	public static SharingMsgType parse(short typ) {
		
		switch(typ) {
		
		case 0:
			return olSharingMsgTypeUnknown;
			
		case 1:
			return olSharingMsgTypeRequest;
			
		case 2:
			return olSharingMsgTypeInvite;
			
		case 3:
			return olSharingMsgTypeInviteAndRequest;
			
		case 4:
			return olSharingMsgTypeResponseAllow;
			
		case 5:
			return olSharingMsgTypeResponseDeny;
			
		default:
			return olSharingMsgTypeUnknown;
		}
	}
}
