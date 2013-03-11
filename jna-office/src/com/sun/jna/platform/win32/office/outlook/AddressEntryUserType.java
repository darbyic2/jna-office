/* Copyright (c) 2013 Ian Darby, All Rights Reserved
 * 
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 * 
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.  
 */

package com.sun.jna.platform.win32.office.outlook;

/**
 * Represents the type of user for the AddressEntry or object derived from
 * AddressEntry.
 * <p>
 * Added in Outlook 2007.
 * </p>
 * 
 * @author Ian Darby
 * 
 * @see {@link AddressEntry}
 * @see {@link AbstractEnum}
 */
public class AddressEntryUserType extends AbstractEnum {
	
	/**
	 * An Exchange user that belongs to the same Exchange forest.
	 */
	public final static AddressEntryUserType	olExchangeUserAddressEntry	= new AddressEntryUserType(0, "olExchangeUserAddressEntry"); 
	
	/**
	 * An address entry that is an Exchange distribution list.
	 */
	public final static AddressEntryUserType	olExchangeDistributionListAddressEntry	= new AddressEntryUserType(1, "olExchangeDistributionListAddressEntry"); 
	
	/**
	 * An address entry that is an Exchange public folder.
	 */
	public final static AddressEntryUserType	olExchangePublicFolderAddressEntry	= new AddressEntryUserType(2, "olExchangePublicFolderAddressEntry"); 
	
	/**
	 * An address entry that is an Exchange agent.
	 */
	public final static AddressEntryUserType	olExchangeAgentAddressEntry	= new AddressEntryUserType(3, "olExchangeAgentAddressEntry"); 
	
	/**
	 * An address entry that is an Exchange organization.
	 */
	public final static AddressEntryUserType	olExchangeOrganizationAddressEntry	= new AddressEntryUserType(4, "olExchangeOrganizationAddressEntry"); 
	
	/**
	 * An Exchange user that belongs to a different Exchange forest.
	 */
	public final static AddressEntryUserType	olExchangeRemoteUserAddressEntry	= new AddressEntryUserType(5, "olExchangeRemoteUserAddressEntry"); 
	
	/**
	 * An address entry in an Outlook Contacts folder.
	 */
	public final static AddressEntryUserType	olOutlookContactAddressEntry	= new AddressEntryUserType(10, "olOutlookContactAddressEntry"); 
	
	/**
	 * An address entry that is an Outlook distribution list.
	 */
	public final static AddressEntryUserType	olOutlookDistributionListAddressEntry	= new AddressEntryUserType(11, "olOutlookDistributionListAddressEntry"); 
	
	/**
	 * An address entry that uses the Lightweight Directory Access Protocol
	 * (LDAP).
	 */
	public final static AddressEntryUserType	olLdapAddressEntry	= new AddressEntryUserType(20, "olLdapAddressEntry"); 
	
	/**
	 * An address entry that uses the Simple Mail Transfer Protocol (SMTP).
	 */
	public final static AddressEntryUserType	olSmtpAddressEntry	= new AddressEntryUserType(30, "olSmtpAddressEntry"); 
	
	/**
	 * A custom or some other type of address entry such as FAX.
	 */
	public final static AddressEntryUserType	olOtherAddressEntry	= new AddressEntryUserType(40, "olOtherAddressEntry"); 

	/**
	 * One and only constructor. Scope is private to prevent the creation of
	 * anything other than the built-in constant instances.
	 * 
	 * @param type
	 *            numeric value used to represent the enum in external storage.
	 * 
	 * @param name
	 *            constant name given to the enum.
	 * 
	 * @see {@link AddressEntry}
	 */
	private AddressEntryUserType(int type, String name) {
		super((short) type, name);
	}

	/**
	 * Converts an external storage numeric representation in to one of the
	 * built-in constant objects. Unrecognised external values cause a
	 * {@link RuntimeException} to be thrown.
	 * 
	 * @param type
	 *            external numeric representation.
	 * 
	 * @return one of the built-in constant objects that represents the enum in
	 *         a type-safe way.
	 */
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
