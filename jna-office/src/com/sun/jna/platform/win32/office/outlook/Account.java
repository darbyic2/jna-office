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

import com.sun.jna.platform.win32.COM.IDispatch;

/**
 * Wrapper for the Account object which represents an account that is defined
 * for the current profile.
 * <p>
 * Added in Outlook 2007.
 * </p>
 * 
 * @author Ian Darby
 * 
 * @see {@link BaseOutlookObject}
 */
public class Account extends BaseOutlookObject {

	/**
	 * Constructor scope is restricted to package as it should not be used
	 * directly by user applications. It is only intended to be used from within
	 * factory methods and properties of the Outlook object model itself. It may
	 * also be called from unit tests which may supply a mock version of the
	 * IDispatch object.
	 * 
	 * @param iDisp
	 *            the IDispatch object which is the underlying Account object
	 *            within the Outlook object model. All methods and properties of
	 *            this wrapper class ultimately delegate to IDispatch.
	 */
	Account(IDispatch iDisp) {
		super(iDisp);
	}

	/**
	 * @return a constant of the AccountType enumeration that indicates the type
	 *         of the Account. Read-only. The AccountType object is a type-safe
	 *         implementation of the OlAccountType enum value returned by the
	 *         Outlook model.
	 *         <p>
	 *         Added in Outlook 2007.
	 *         </p>
	 * 
	 * @see {@link AccountType}
	 */
	public AccountType getAccountType() {

		return AccountType.parse(getShortProperty("AccountType"));
	}

	/**
	 * @return an AutoDiscoverConnectionMode constant that specifies the type of
	 *         connection to use for the auto-discovery service of the Microsoft
	 *         Exchange server that hosts the account mailbox. Read-only. The
	 *         AutoDiscoverConnectionMode object is a type-safe implementation
	 *         of the OlAutoDiscoverConnectionMode enum value returned by the
	 *         Outlook model.
	 *         <p>
	 *         Added in Outlook 2010.
	 *         </p>
	 * 
	 * @see {@link AutoDiscoverConnectionMode}
	 */
	public AutoDiscoverConnectionMode getAutoDiscoverConnectionMode() {

		return AutoDiscoverConnectionMode
				.parse(getShortProperty("AutoDiscoverConnectionMode"));
	}

	/**
	 * @return a String that represents information in XML retrieved from the
	 *         auto-discovery service of the Microsoft Exchange Server that is
	 *         associated with the account. Read-only.
	 *         <p>
	 *         Added in Outlook 2010.
	 *         </p>
	 */
	public String getAutoDiscoverXML() {

		return getStringProperty("AutoDiscoverXML");
	}

	/**
	 * @return a Recipient object that represents the current user identity for
	 *         the account. Read-only.
	 *         <p>
	 *         Added in Outlook 2010.
	 *         </p>
	 * 
	 * @see {@link Recipient}
	 */
	public Recipient getCurrentUser() {

		return new Recipient(getAutomationProperty("CurrentUser"));
	}

	/**
	 * @return a Store object that represents the default delivery store for the
	 *         account. Read-only.
	 *         <p>
	 *         Added in Outlook 2010.
	 *         </p>
	 * 
	 * @see {@link Store}
	 */
	public Store getDeliveryStore() {

		return new Store(getAutomationProperty("DeliveryStore"));
	}

	/**
	 * @return a String representing the display name of the e-mail Account.
	 *         Read-only.
	 *         <p>
	 *         Added in Outlook 2007.
	 *         </p>
	 */
	public String getDisplayName() {

		return getStringProperty("DisplayName");
	}

	/**
	 * @return an ExchangeConnectionMode constant that indicates the current
	 *         connection mode for the Microsoft Exchange Server that hosts the
	 *         account mailbox. Read-only. The ExchangeConnectionMode object is
	 *         a type-safe implementation of the olExchangeConnectionMode enum
	 *         value returned by the Outlook model.
	 *         <p>
	 *         Added in Outlook 2010.
	 *         </p>
	 * 
	 * @see {@link ExchangeConnectionMode}
	 */
	public ExchangeConnectionMode getExchangeConnectionMode() {

		return ExchangeConnectionMode
				.parse(getShortProperty("ExchangeConnectionMode"));
	}

	/**
	 * @return a String value that represents the name of the Microsoft Exchange
	 *         Server that hosts the account mailbox. Read-only.
	 *         <p>
	 *         Added in Outlook 2010.
	 *         </p>
	 */
	public String getExchangeMailboxServerName() {

		return getStringProperty("ExchangeMailboxServerName");
	}

	/**
	 * @return the full version number of the Microsoft Exchange Server that
	 *         hosts the account mailbox. Read-only.
	 *         <p>
	 *         Added in Outlook 2010.
	 *         </p>
	 */
	public String getExchangeMailboxServerVersion() {

		return getStringProperty("ExchangeMailboxServerVersion");
	}

	/**
	 * Returns an AddressEntry object that represents the address entry
	 * specified by the given entry ID.
	 * 
	 * @param id
	 *            Used to identify an address entry that is maintained for the
	 *            session.
	 * 
	 * @return an AddressEntry object that represents the address entry
	 *         specified by the given entry ID.
	 *         <p>
	 *         Added in Outlook 2010.
	 *         </p>
	 * 
	 * @see {@link AddressEntry}
	 */
	public AddressEntry getAddressEntryFromID(String id) {

		return new AddressEntry((IDispatch) invoke("GetAddressEntryFromID",
				newVariant(id)).getValue());
	}

	/**
	 * Returns the Recipient object that is identified by the given entry ID.
	 * 
	 * @param id
	 *            The EntryID of the recipient.
	 * 
	 * @return the Recipient object that is identified by the given entry ID.
	 *         <p>
	 *         Added in Outlook 2010.
	 *         </p>
	 * 
	 * @see {@link Recipient}
	 */
	public Recipient GetRecipientFromID(String id) {

		return new Recipient((IDispatch) invoke("GetRecipientFromID",
				newVariant(id)).getValue());
	}

	/**
	 * @return a String representing the Simple Mail Transfer Protocol (SMTP)
	 *         address for the Account. Read-only.
	 *         <p>
	 *         Added in Outlook 2007.
	 *         </p>
	 */
	public String getSmtpAddress() {

		return getStringProperty("SmtpAddress");
	}

	/**
	 * @return a String representing the user name for the Account. Read-only.
	 *         <p>
	 *         Added in Outlook 2007.
	 *         </p>
	 */
	public String getUserName() {

		return getStringProperty("UserName");
	}

}
