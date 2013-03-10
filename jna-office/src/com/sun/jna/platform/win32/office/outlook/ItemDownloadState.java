package com.sun.jna.platform.win32.office.outlook;

public class ItemDownloadState extends AbstractEnum {
	
	public final static ItemDownloadState olHeaderOnly = new ItemDownloadState(0, "olHeaderOnly");	//Only the header has been downloaded.
	public final static ItemDownloadState olFullItem = new ItemDownloadState(1, "olFullItem");		//Full item has been downloaded.

	private ItemDownloadState(int type, String name) {
		super((short) type, name);
	}

	public static ItemDownloadState parse(short state) {
		
		switch(state) {
		
		case 1:
			return olFullItem;
			
		case 0:
		default:
			return olHeaderOnly;
		}
	}
}
