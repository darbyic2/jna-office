package com.sun.jna.platform.win32.office.outlook;

public class SharingProvider extends AbstractEnum {
	
	public final static SharingProvider	olProviderUnknown		= new SharingProvider(0, "olProviderUnknown"); //Represents an unknown sharing provider. This value is used if the sharing provider GUID in the sharing message does not match the GUID of any of the sharing providers represented in this enumeration.
	public final static SharingProvider	olProviderExchange		= new SharingProvider(1, "olProviderExchange"); //Represents the Exchange sharing provider.
	public final static SharingProvider	olProviderWebCal		= new SharingProvider(2, "olProviderWebCal"); //Represents the WebCal sharing provider.
	public final static SharingProvider	olProviderPubCal		= new SharingProvider(3, "olProviderPubCal"); //Represents the PubCal sharing provider.
	public final static SharingProvider	olProviderICal			= new SharingProvider(4, "olProviderICal"); //Represents the iCalendar sharing provider.
	public final static SharingProvider	olProviderSharePoint	= new SharingProvider(5, "olProviderSharePoint"); //Represents the SharePoint sharing provider.
	public final static SharingProvider	olProviderRSS			= new SharingProvider(6, "olProviderRSS"); //Represents the Really Simple Syndication (RSS) sharing provider.
	public final static SharingProvider	olProviderFederate		= new SharingProvider(7, "olProviderFederate"); //Represents a federated sharing provider. A SharingItem object with this type of provider is used for sharing relationships across organizational boundares (for example, between two organizations using Microsoft Exchange Server 2010).
	
	private SharingProvider(int val, String name) {
		super((short) val, name);
	}

	public static SharingProvider parse(short val) {
		
		switch(val) {
		
		case 0:
			return olProviderUnknown;
		
		case 1:
			return olProviderExchange;
		
		case 2:
			return olProviderWebCal;
		
		case 3:
			return olProviderPubCal;
		
		case 4:
			return olProviderICal;
		
		case 5:
			return olProviderSharePoint;
		
		case 6:
			return olProviderRSS;
		
		case 7:
			return olProviderFederate;
		
		default:
			return olProviderUnknown;
		}
	}
}
