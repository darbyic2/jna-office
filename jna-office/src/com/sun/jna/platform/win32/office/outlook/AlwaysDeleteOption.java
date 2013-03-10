package com.sun.jna.platform.win32.office.outlook;

public class AlwaysDeleteOption extends AbstractEnum {

	public final static AlwaysDeleteOption	olDoNotDelete	= new AlwaysDeleteOption(0, "olDoNotDelete"); //New items joining the conversation are not moved to the the Deleted Items folder on the specified delivery store, and existing conversation items in the Deleted Items folder are moved to the Inbox.
	public final static AlwaysDeleteOption	olAlwaysDelete	= new AlwaysDeleteOption(1, "olAlwaysDelete"); //New items of the conversation are always moved to the Deleted Items folder for the store that contains the items
	public final static AlwaysDeleteOption	olAlwaysDeleteUnsupported	= new AlwaysDeleteOption(2, "olAlwaysDeleteUnsupported"); //The specified store does not support the action of always moving items to the Deleted Items folder of that store.
	
	private AlwaysDeleteOption(int option, String name) {
		super((short) option, name);
	}
	
	public static AlwaysDeleteOption parse(short option) {
		
		switch(option) {
		
		case 0:
			return olDoNotDelete;
			
		case 1:
			return olAlwaysDelete;
			
		case 2:
			return olAlwaysDeleteUnsupported;
			
		default:
			return olDoNotDelete;
		}
	}
}
