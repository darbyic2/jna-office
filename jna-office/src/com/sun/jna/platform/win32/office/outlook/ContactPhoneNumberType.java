package com.sun.jna.platform.win32.office.outlook;

public class ContactPhoneNumberType extends AbstractEnum {
	
	public final static ContactPhoneNumberType	olContactPhoneAssistant		= new ContactPhoneNumberType(0,  "olContactPhoneAssistant"); //Telephone number of the person who is the assistant for the contact
	public final static ContactPhoneNumberType	olContactPhoneBusiness		= new ContactPhoneNumberType(1,  "olContactPhoneBusiness"); //Business telephone number
	public final static ContactPhoneNumberType	olContactPhoneBusiness2		= new ContactPhoneNumberType(2,  "olContactPhoneBusiness2"); //Second business telephone number
	public final static ContactPhoneNumberType	olContactPhoneBusinessFax	= new ContactPhoneNumberType(3,  "olContactPhoneBusinessFax"); //Business fax number
	public final static ContactPhoneNumberType	olContactPhoneCallback		= new ContactPhoneNumberType(4,  "olContactPhoneCallback"); //Callback telephone number
	public final static ContactPhoneNumberType	olContactPhoneCar			= new ContactPhoneNumberType(5,  "olContactPhoneCar"); //Car telephone number
	public final static ContactPhoneNumberType	olContactPhoneCompany		= new ContactPhoneNumberType(6,  "olContactPhoneCompany"); //Main company telephone number
	public final static ContactPhoneNumberType	olContactPhoneHome			= new ContactPhoneNumberType(7,  "olContactPhoneHome"); //Home telephone number
	public final static ContactPhoneNumberType	olContactPhoneHome2			= new ContactPhoneNumberType(8,  "olContactPhoneHome2"); //Second home telephone number
	public final static ContactPhoneNumberType	olContactPhoneHomeFax		= new ContactPhoneNumberType(9,  "olContactPhoneHomeFax"); //Home fax number
	public final static ContactPhoneNumberType	olContactPhoneISDN			= new ContactPhoneNumberType(10, "olContactPhoneISDN"); //Integrated Services Digital Network (ISDN) phone number
	public final static ContactPhoneNumberType	olContactPhoneMobile		= new ContactPhoneNumberType(11, "olContactPhoneMobile"); //Mobile telephone number
	public final static ContactPhoneNumberType	olContactPhoneOther			= new ContactPhoneNumberType(12, "olContactPhoneOther"); //Other telephone number
	public final static ContactPhoneNumberType	olContactPhoneOtherFax		= new ContactPhoneNumberType(13, "olContactPhoneOtherFax"); //Other fax number
	public final static ContactPhoneNumberType	olContactPhonePager			= new ContactPhoneNumberType(14, "olContactPhonePager"); //Pager telephone number
	public final static ContactPhoneNumberType	olContactPhonePrimary		= new ContactPhoneNumberType(15, "olContactPhonePrimary"); //Primary telephone number
	public final static ContactPhoneNumberType	olContactPhoneRadio			= new ContactPhoneNumberType(16, "olContactPhoneRadio"); //Radio telephone number
	public final static ContactPhoneNumberType	olContactPhoneTelex			= new ContactPhoneNumberType(17, "olContactPhoneTelex"); //Telex telephone number
	public final static ContactPhoneNumberType	olContactPhoneTTYTTD		= new ContactPhoneNumberType(18, "olContactPhoneTTYTTD"); //TTD/TTY (Teletypewriting Device for the Deaf/Teletypewriter) telephone number
	
	private ContactPhoneNumberType(int typ, String name) {
		super((short) typ, name);
	}

	public static ContactPhoneNumberType parse(short typ) {
		
		switch(typ) {
		
		case 0:
			return olContactPhoneAssistant;
			
		case 1:
			return olContactPhoneBusiness;
			
		case 2:
			return olContactPhoneBusiness2;
			
		case 3:
			return olContactPhoneBusinessFax;
			
		case 4:
			return olContactPhoneCallback;
			
		case 5:
			return olContactPhoneCar;
			
		case 6:
			return olContactPhoneCompany;
			
		case 7:
			return olContactPhoneHome;
			
		case 8:
			return olContactPhoneHome2;
			
		case 9:
			return olContactPhoneHomeFax;
			
		case 10:
			return olContactPhoneISDN;
			
		case 11:
			return olContactPhoneMobile;
			
		case 12:
			return olContactPhoneOther;
			
		case 13:
			return olContactPhoneOtherFax;
			
		case 14:
			return olContactPhonePager;
			
		case 15:
			return olContactPhonePrimary;
			
		case 16:
			return olContactPhoneRadio;
			
		case 17:
			return olContactPhoneTelex;
			
		case 18:
			return olContactPhoneTTYTTD;
			
		default:
			throw new RuntimeException("ContactPhoneNumberType Enum: " + typ + " not recognised.");
		}
	}
}
