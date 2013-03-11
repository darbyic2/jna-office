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

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;

/**
 * Represents a person, group, or public folder to which the messaging system
 * can deliver messages.
 * 
 * @author Ian Darby
 * 
 * @see {@link AddressEntries}
 */
public class AddressEntry extends BaseOutlookObject {

	/**
	 * Constructor scope is restricted to package as it should not be used
	 * directly by user applications. It is only intended to be used from within
	 * factory methods and properties of the Outlook object model itself. It may
	 * also be called from unit tests which may supply a mock version of the
	 * IDispatch object.
	 * 
	 * @param iDisp
	 *            the IDispatch object which is the underlying Actions object
	 *            within the Outlook object model. All methods and properties of
	 *            this wrapper class ultimately delegate to IDispatch.
	 */
	AddressEntry(IDispatch iDisp) {
		super(iDisp);
	}
	
	/**
	 * Returns a String representing the e-mail address of the AddressEntry.
	 * Read/write.
	 * 
	 * @return a String representing the e-mail address of the AddressEntry.
	 *         Read/write.
	 */
	public String getAddress() {
		
		return getStringProperty("Address");
	}
	
	/**
	 * Sets a String representing the e-mail address of the AddressEntry.
	 * Read/write.
	 * 
	 * @param address
	 *            email address to use for this AddressEntry.
	 */
	public void setAddress(String address) {
		
		setProperty("Address", address);
	}
	
	/**
	 * Returns a constant from the AddressEntryUserType enumeration representing
	 * the user type of the AddressEntry. Read-only.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a constant from the AddressEntryUserType enumeration representing
	 *         the user type of the AddressEntry.
	 */
	public AddressEntryUserType getAddressEntryUserType() {
		
		return AddressEntryUserType.parse(getShortProperty("AddressEntryUserType"));
	}
	
	/**
	 * Deletes itself from the collection.
	 */
	public void delete() {
		
		invokeNoReply("Delete");
	}
	
	/**
	 * Displays a modeless dialog box that provides detailed information about
	 * an AddressEntry object.
	 * <p>
	 * You must use error handling to handle run-time errors when the user
	 * clicks Cancel in the dialog box. The Details method actually stops the
	 * code from running while the dialog box is displayed.
	 * </p>
	 */
	public void showDetails() {
		
		invokeNoReply("Details");
	}
	
	/**
	 * Displays a modeless dialog box that provides detailed information about
	 * an AddressEntry object.
	 * <p>
	 * You must use error handling to handle run-time errors when the user
	 * clicks Cancel in the dialog box. The Details method actually stops the
	 * code from running while the dialog box is displayed.
	 * </p>
	 * 
	 * @param hWnd
	 *            The parent window handle for the Details dialog box. A zero
	 *            value (the default) specifies that the dialog is parented to
	 *            Outlook.
	 */
	public void showDetails(int hWnd) {
		
		invokeNoReply("Details", newVariant(hWnd));
	}
	
	/**
	 * Returns a constant belonging to the DisplayType enumeration that
	 * describes the nature of the AddressEntry. Read-only.
	 * 
	 * @return a constant belonging to the DisplayType enumeration that
	 *         describes the nature of the AddressEntry.
	 */
	public DisplayType getDisplayType() {
		
		return DisplayType.parse(getShortProperty("DisplayType"));
	}
	
	/**
	 * Returns a ContactItem object that represents the AddressEntry, if the
	 * AddressEntry corresponds to a contact in an Outlook Contacts Address Book
	 * (CAB).
	 * 
	 * @return a ContactItem object that represents the AddressEntry, if the
	 *         AddressEntry corresponds to a contact in an Outlook Contacts
	 *         Address Book (CAB).
	 */
	public ContactItem getContact() {
		
		return new ContactItem((IDispatch) invoke("GetContact").getValue());
	}
	
	/**
	 * Returns an ExchangeDistributionList object that represents the
	 * AddressEntry if the AddressEntry belongs to an Exchange AddressList
	 * object such as the Global Address List (GAL) and corresponds to an
	 * Exchange distribution list.
	 * 
	 * @return an ExchangeDistributionList object that represents the
	 *         AddressEntry if the AddressEntry belongs to an Exchange
	 *         AddressList object such as the Global Address List (GAL) and
	 *         corresponds to an Exchange distribution list.
	 */
	public ExchangeDistributionList getExchangeDistributionList() {
		
		return new ExchangeDistributionList((IDispatch) invoke("GetExchangeDistributionList").getValue());
	}
	
	/**
	 * Returns an ExchangeUser object that represents the AddressEntry if the
	 * AddressEntry belongs to an Exchange AddressList object such as the Global
	 * Address List (GAL) and corresponds to an Exchange user.
	 * <p>
	 * You have to be connected to the Exchange server to use this method.
	 * </p>
	 * 
	 * @return an ExchangeUser object that represents the AddressEntry if the
	 *         AddressEntry belongs to an Exchange AddressList object such as
	 *         the Global Address List (GAL) and corresponds to an Exchange
	 *         user.
	 */
	public ExchangeUser getExchangeUser() {
		
		return new ExchangeUser((IDispatch) invoke("GetExchangeUser").getValue());
	}
	
	/**
	 * Returns a String value that represents the availability of the individual
	 * user for a period of 30 days from the start date, beginning at midnight
	 * of the date specified.
	 * 
	 * @param start
	 *            Specifies the start date.
	 * 
	 * @param minsPerChar
	 *            Specifies the length of each time slot in minutes. The default
	 *            value is 30.
	 * 
	 * @return a String value that represents the availability of the individual
	 *         user for a period of 30 days from the start date, beginning at
	 *         midnight of the date specified.
	 */
	public String getFreeBusy(Date start, int minsPerChar) {
		
		return getFreeBusy(start, minsPerChar, false);
	}
	
	/**
	 * Returns a String value that represents the availability of the individual
	 * user for a period of 30 days from the start date, beginning at midnight
	 * of the date specified.
	 * 
	 * @param start
	 *            Specifies the start date.
	 * 
	 * @param minsPerChar
	 *            Specifies the length of each time slot in minutes. The default
	 *            value is 30.
	 * 
	 * @param useCompleteFormat
	 *            Specifies a Boolean value that represents the level of
	 *            information returned for each time slot. The default value is
	 *            False.
	 * 
	 * @return a String value that represents the availability of the individual
	 *         user for a period of 30 days from the start date, beginning at
	 *         midnight of the date specified.
	 */
	public String getFreeBusy(Date start, int minsPerChar, boolean useCompleteFormat) {
		
		return invoke("GetFreeBusy", newVariant(start), newVariant(minsPerChar), newVariant(useCompleteFormat)).getValue().toString();
	}
	
	/**
	 * Returns a String representing the unique identifier for the object.
	 * Read-only.
	 * 
	 * @return a String representing the unique identifier for the object.
	 */
	public String getID() {
		
		return getStringProperty("ID");
	}
	
	/**
	 * Returns the display name for the object. Read/write.
	 * 
	 * @return the display name for the object.
	 */
	public String getName() {
		
		return getStringProperty("Name");
	}
	
	/**
	 * Sets the display name for the object. Read/write.
	 * 
	 * @param name
	 *            the display name to be used.
	 */
	public void setName(String name) {
		
		setProperty("Name", name);
	}
	
	/**
	 * Returns a PropertyAccessor object that supports creating, getting,
	 * setting, and deleting properties of the parent AddressEntry object.
	 * Read-only.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a PropertyAccessor object that supports creating, getting,
	 *         setting, and deleting properties of the parent AddressEntry
	 *         object.
	 */
	public PropertyAccessor getPropertyAccessor() {
		
		return new PropertyAccessor(getAutomationProperty("PropertyAccessor"));
	}
	
	/**
	 * Returns a String representing the type of entry for this address such as
	 * an Internet Address, MacMail Address, or Microsoft Mail Address.
	 * Read/write.
	 * 
	 * @return Returns a String representing the type of entry for this address
	 *         such as an Internet Address, MacMail Address, or Microsoft Mail
	 *         Address.
	 */
	public String getType() {
		
		return getStringProperty("Type");
	}
	
	/**
	 * Sets a String representing the type of entry for this address such as an
	 * Internet Address, MacMail Address, or Microsoft Mail Address. Read/write.
	 * 
	 * @param typ
	 *            entry type.
	 */
	public void setType(String typ) {
		
		setProperty("Type", typ);
	}
	
	/**
	 * Posts a change to the AddressEntry object in the messaging system.
	 * <p>
	 * Equivalent to update(true, false);
	 * </p>
	 */
	public void update() {
		
		update(true, false);
	}
	
	/**
	 * Posts a change to the AddressEntry object in the messaging system.
	 * <p>
	 * Equivalent to update(makePermanent, false);
	 * </p>
	 * 
	 * @param makePermanent
	 *            A value of True indicates that the property cache is flushed
	 *            and all changes are committed in the underlying address book.
	 *            A value of False indicates that the property cache is flushed
	 *            but not committed to persistent storage. The default value is
	 *            True.
	 */
	public void update(boolean makePermanent) {
		
		update(makePermanent, false);
	}
	
	/**
	 * Posts a change to the AddressEntry object in the messaging system.
	 * 
	 * @param makePermanent
	 *            A value of True indicates that the property cache is flushed
	 *            and all changes are committed in the underlying address book.
	 *            A value of False indicates that the property cache is flushed
	 *            but not committed to persistent storage. The default value is
	 *            True.
	 * 
	 * @param refresh
	 *            A value of True indicates that the property cache is reloaded
	 *            from the values in the underlying address book. A value of
	 *            False indicates that the property cache is not reloaded. The
	 *            default value is False.
	 */
	public void update(boolean makePermanent, boolean refresh) {
		
		invokeNoReply("Update", newVariant(makePermanent), newVariant(refresh));
	}
	
}
