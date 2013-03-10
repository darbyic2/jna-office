package com.sun.jna.platform.win32.office.outlook;

public class DisplayType extends AbstractEnum {

	public final static DisplayType	olUser	= new DisplayType(0, "olUser");   //	User address
	public final static DisplayType	olDistList	= new DisplayType(1, "olDistList");   //	Exchange distribution list
	public final static DisplayType	olForum	= new DisplayType(2, "olForum");   //	Forum address
	public final static DisplayType	olAgent	= new DisplayType(3, "olAgent");   //	Agent address
	public final static DisplayType	olOrganization	= new DisplayType(4, "olOrganization");   //	Organization address
	public final static DisplayType	olPrivateDistList	= new DisplayType(5, "olPrivateDistList");   //	Outlook private distribution list
	public final static DisplayType	olRemoteUser	= new DisplayType(6, "olRemoteUser");   //	Remote user address
	
	private DisplayType(int typ, String name) {
		super((short) typ, name);
	}
	
	public static DisplayType parse(short typ) {
		
		switch(typ) {
		
			case 0:
				return olUser;
				
			case 1:
				return olDistList;
				
			case 2:
				return olForum;
				
			case 3:
				return olAgent;
				
			case 4:
				return olOrganization;
				
			case 5:
				return olPrivateDistList;
				
			case 6:
				return olRemoteUser;
				
			default:
				return null;
		}
	}
}
