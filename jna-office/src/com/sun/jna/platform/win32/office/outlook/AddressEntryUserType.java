package com.sun.jna.platform.win32.office.outlook;

public class AddressEntryUserType extends AbstractEnum {
	
	public final static AddressEntryUserType	olExchangeUserAddressEntry	= new AddressEntryUserType(0, "olExchangeUserAddressEntry"); //An Exchange user that belongs to the same Exchange forest.
	public final static AddressEntryUserType	olExchangeDistributionListAddressEntry	= new AddressEntryUserType(1, "olExchangeDistributionListAddressEntry"); //An address entry that is an Exchange distribution list.
	public final static AddressEntryUserType	olExchangePublicFolderAddressEntry	= new AddressEntryUserType(2, "olExchangePublicFolderAddressEntry"); //An address entry that is an Exchange public folder.
	public final static AddressEntryUserType	olExchangeAgentAddressEntry	= new AddressEntryUserType(3, "olExchangeAgentAddressEntry"); //An address entry that is an Exchange agent.
	public final static AddressEntryUserType	olExchangeOrganizationAddressEntry	= new AddressEntryUserType(4, "olExchangeOrganizationAddressEntry"); //An address entry that is an Exchange organization.
	public final static AddressEntryUserType	olExchangeRemoteUserAddressEntry	= new AddressEntryUserType(5, "olExchangeRemoteUserAddressEntry"); //An Exchange user that belongs to a different Exchange forest.
	public final static AddressEntryUserType	olOutlookContactAddressEntry	= new AddressEntryUserType(10, "olOutlookContactAddressEntry"); //An address entry in an Outlook Contacts folder.
	public final static AddressEntryUserType	olOutlookDistributionListAddressEntry	= new AddressEntryUserType(11, "olOutlookDistributionListAddressEntry"); //An address entry that is an Outlook distribution list.
	public final static AddressEntryUserType	olLdapAddressEntry	= new AddressEntryUserType(20, "olLdapAddressEntry"); //An address entry that uses the Lightweight Directory Access Protocol (LDAP).
	public final static AddressEntryUserType	olSmtpAddressEntry	= new AddressEntryUserType(30, "olSmtpAddressEntry"); //An address entry that uses the Simple Mail Transfer Protocol (SMTP).
	public final static AddressEntryUserType	olOtherAddressEntry	= new AddressEntryUserType(40, "olOtherAddressEntry"); //A custom or some other type of address entry such as FAX.

	private AddressEntryUserType(int type, String name) {
		super((short) type, name);
	}

	public static AddressEntryUserType parse(short type) {
		
		switch(type) {
		
		case 0:
			return olExchangeUserAddressEntry;
			
		case 1:
			return olExchangeDistributionListAddressEntry;
			
		case 2:
			return olExchangePublicFolderAddressEntry;
			
		case 3:
			return olExchangeAgentAddressEntry;
			
		case 4:
			return olExchangeOrganizationAddressEntry;
			
		case 5:
			return olExchangeRemoteUserAddressEntry;
			
		case 10:
			return olOutlookContactAddressEntry;
			
		case 11:
			return olOutlookDistributionListAddressEntry;
			
		case 20:
			return olLdapAddressEntry;
			
		case 30:
			return olSmtpAddressEntry;
			
		case 40:
			return olOtherAddressEntry;
		
		default:
			throw new RuntimeException("AddressEntryUserType Enum: " + type + " not recognised.");
		}
	}
}
