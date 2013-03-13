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
 * Represents a contact in a Contacts folder.
 * <p>
 * A contact can represent any person with whom you have any personal or
 * professional contact.
 * </p>
 * <p>
 * Use the {@link Outlook#createContactItem()} method to create a ContactItem
 * object that represents a new contact.
 * </p>
 * <p>
 * Use Items (index), where index is the index number of a contact or a value
 * used to match the default property of a contact, to return a single
 * ContactItem object from a Contacts folder.
 * </p>
 * 
 * @author Ian Darby
 * 
 * @see BaseOutlookObject
 * @see BaseItemLevel1
 * @see BaseItemLevel2
 * @see BaseItemLevel3
 */
public class ContactItem extends BaseItemLevel3 {

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
	ContactItem(IDispatch iDisp) {
		super(iDisp);
	}
	
	/**
	 * Returns or sets a String representing the account for the contact.
	 * Read/write.
	 * 
	 * @return a String representing the account for the contact.
	 */
	public String getAccount() {
		
		return getStringProperty("Account");
	}
	
	/**
	 * Returns or sets a String representing the account for the contact.
	 * Read/write.
	 * 
	 * @param ac
	 *            a String representing the account for the contact.
	 */
	public void setAccount(String ac) {
		
		setProperty("Account", ac);
	}
	
	/**
	 * Adds a logo picture to the current Electronic Business Card of the
	 * contact item.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param path
	 *            The full path name that specifies the picture file to load.
	 */
	public void addBusinessCardLogoPicture(String path) {
		
		invokeNoReply("AddBusinessCardLogoPicture", newVariant(path));
	}
	
	/**
	 * Adds a picture to a contact item.
	 * <p>
	 * If the contact item already has a picture attached to it, this method
	 * will overwrite the existing picture.
	 * </p>
	 * <p>
	 * The picture can be an icon, GIF, JPEG, BMP, TIFF, WMF, EMF, or PNG file.
	 * Microsoft Outlook will automatically perform the necessary resizing of
	 * the picture.
	 * </p>
	 * 
	 * @param path
	 *            A string containing the complete path and filename of the
	 *            picture to be added to the contact item.
	 */
	public void addPicture(String path) {
		
		invokeNoReply("AddPicture", newVariant(path));
	}
	
	/**
	 * Returns or sets a Date indicating the anniversary date for the contact.
	 * Read/write.
	 * 
	 * @return a Date indicating the anniversary date for the contact.
	 */
	public Date getAnniversary() {
		
		return getDateProperty("Anniversary");
	}
	
	/**
	 * Returns or sets a Date indicating the anniversary date for the contact.
	 * Read/write.
	 * 
	 * @param dat
	 *            a Date indicating the anniversary date for the contact.
	 */
	public void setAnniversary(Date dat) {
		
		setProperty("Anniversary", dat);
	}
	
	/**
	 * Returns or sets a String representing the name of the person who is the
	 * assistant for the contact. Read/write.
	 * 
	 * @return a String representing the name of the person who is the assistant
	 *         for the contact.
	 */
	public String getAssistantName() {
		
		return getStringProperty("AssistantName");
	}
	
	/**
	 * Returns or sets a String representing the name of the person who is the
	 * assistant for the contact. Read/write.
	 * 
	 * @param name
	 *            a String representing the name of the person who is the
	 *            assistant for the contact.
	 */
	public void setAssistantName(String name) {
		
		setProperty("AssistantName", name);
	}
	
	/**
	 * Returns or sets a String representing the telephone number of the person
	 * who is the assistant for the contact. Read/write.
	 * 
	 * @return a String representing the telephone number of the person who is
	 *         the assistant for the contact.
	 */
	public String getAssistantTelephoneNumber() {
		
		return getStringProperty("AssistantTelephoneNumber");
	}
	
	/**
	 * Returns or sets a String representing the telephone number of the person
	 * who is the assistant for the contact. Read/write.
	 * 
	 * @param tel
	 *            a String representing the telephone number of the person who
	 *            is the assistant for the contact.
	 */
	public void setAssistantTelephoneNumber(String tel) {
		
		setProperty("AssistantTelephoneNumber", tel);
	}
	
	/**
	 * Returns or sets a Date indicating the birthday for the contact.
	 * Read/write.
	 * 
	 * @return a Date indicating the birthday for the contact.
	 */
	public Date getBirthday() {
		
		return getDateProperty("Birthday");
	}
	
	/**
	 * Returns or sets a Date indicating the birthday for the contact.
	 * Read/write.
	 * 
	 * @param dat
	 *            a Date indicating the birthday for the contact.
	 */
	public void setBirthday(Date dat) {
		
		setProperty("Birthday", dat);
	}
	
	/**
	 * Returns or sets a String representing the second business telephone
	 * number for the contact. Read/write.
	 * 
	 * @return a String representing the second business telephone number for
	 *         the contact.
	 */
	public String getBusiness2TelephoneNumber() {
		
		return getStringProperty("Business2TelephoneNumber");
	}
	
	/**
	 * Returns or sets a String representing the second business telephone
	 * number for the contact. Read/write.
	 * 
	 * @param tel
	 *            a String representing the second business telephone number for
	 *            the contact.
	 */
	public void setBusiness2TelephoneNumber(String tel) {
		
		setProperty("Business2TelephoneNumber", tel);
	}
	
	/**
	 * Returns or sets a String representing the whole, unparsed business
	 * address for the contact. Read/write.
	 * 
	 * @return a String representing the whole, unparsed business address for
	 *         the contact.
	 */
	public String getBusinessAddress() {
		
		return getStringProperty("BusinessAddress");
	}
	
	/**
	 * Returns or sets a String representing the whole, unparsed business
	 * address for the contact. Read/write.
	 * 
	 * @param addr
	 *            a String representing the whole, unparsed business address for
	 *            the contact.
	 */
	public void setBusinessAddress(String addr) {
		
		setProperty("BusinessAddress", addr);
	}
	
	/**
	 * Returns or sets a String representing the city name portion of the
	 * business address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the BusinessAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to the BusinessAddress property.
	 * </p>
	 * 
	 * @return a String representing the city name portion of the business
	 *         address for the contact.
	 */
	public String getBusinessAddressCity() {
		
		return getStringProperty("BusinessAddressCity");
	}
	
	/**
	 * Returns or sets a String representing the city name portion of the
	 * business address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the BusinessAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to the BusinessAddress property.
	 * </p>
	 * 
	 * @param addr
	 *            a String representing the city name portion of the business
	 *            address for the contact.
	 */
	public void setBusinessAddressCity(String addr) {
		
		setProperty("BusinessAddressCity", addr);
	}
	
	/**
	 * Returns or sets a String representing the country/region code portion of
	 * the business address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the BusinessAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to the BusinessAddress property.
	 * </p>
	 * 
	 * @return
	 */
	public String getBusinessAddressCountry() {
		
		return getStringProperty("BusinessAddressCountry");
	}
	
	/**
	 * Returns or sets a String representing the country/region code portion of
	 * the business address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the BusinessAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to the BusinessAddress property.
	 * </p>
	 * 
	 * @param addr
	 *            a String representing the country/region code portion of the
	 *            business address for the contact.
	 */
	public void setBusinessAddressCountry(String addr) {
		
		setProperty("BusinessAddressCountry", addr);
	}
	
	/**
	 * Returns or sets a String representing the postal code (zip code) portion
	 * of the business address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the BusinessAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to the BusinessAddress property.
	 * </p>
	 * 
	 * @return a String representing the postal code (zip code) portion of the
	 *         business address for the contact.
	 */
	public String getBusinessAddressPostalCode() {
		
		return getStringProperty("BusinessAddressPostalCode");
	}
	
	/**
	 * Returns or sets a String representing the postal code (zip code) portion
	 * of the business address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the BusinessAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to the BusinessAddress property.
	 * </p>
	 * 
	 * @param code
	 *            a String representing the postal code (zip code) portion of
	 *            the business address for the contact.
	 */
	public void setBusinessAddressPostalCode(String code) {
		
		setProperty("BusinessAddressPostalCode", code);
	}
	
	/**
	 * Returns or sets a String representing the post office box number portion
	 * of the business address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the BusinessAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to the BusinessAddress property.
	 * </p>
	 * 
	 * @return a String representing the post office box number portion of the
	 *         business address for the contact.
	 */
	public String getBusinessAddressPostOfficeBox() {
		
		return getStringProperty("BusinessAddressPostOfficeBox");
	}
	
	/**
	 * Returns or sets a String representing the post office box number portion
	 * of the business address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the BusinessAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to the BusinessAddress property.
	 * </p>
	 * 
	 * @param boxNum
	 *            a String representing the post office box number portion of
	 *            the business address for the contact.
	 */
	public void setBusinessAddressPostOfficeBox(String boxNum) {
		
		setProperty("BusinessAddressPostOfficeBox", boxNum);
	}
	
	/**
	 * Returns or sets a String representing the state code portion of the
	 * business address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the BusinessAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to the BusinessAddress property.
	 * </p>
	 * 
	 * @return a String representing the state code portion of the business
	 *         address for the contact.
	 */
	public String getBusinessAddressState() {
		
		return getStringProperty("BusinessAddressState");
	}
	
	/**
	 * Returns or sets a String representing the state code portion of the
	 * business address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the BusinessAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to the BusinessAddress property.
	 * </p>
	 * 
	 * @param state
	 *            a String representing the state code portion of the business
	 *            address for the contact.
	 */
	public void setBusinessAddressState(String state) {
		
		setProperty("BusinessAddressState", state);
	}
	
	/**
	 * Returns or sets a String representing the street address portion of the
	 * business address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the BusinessAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to the BusinessAddress property.
	 * </p>
	 * 
	 * @return a String representing the street address portion of the business
	 *         address for the contact.
	 */
	public String getBusinessAddressStreet() {
		
		return getStringProperty("BusinessAddressStreet");
	}
	
	/**
	 * Returns or sets a String representing the street address portion of the
	 * business address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the BusinessAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to the BusinessAddress property.
	 * </p>
	 * 
	 * @param street
	 *            a String representing the street address portion of the
	 *            business address for the contact.
	 */
	public void setBusinessAddressStreet(String street) {
		
		setProperty("BusinessAddressStreet", street);
	}
	
	/**
	 * Returns or sets a String that represents the XML markup for the layout of
	 * the Electronic Business Card. Read/write.
	 * <p>
	 * For more information on the XML schema for Electronic Business Cards, see
	 * the Microsoft Outlook 2010 XML Schema Reference in the MSDN Library.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a String that represents the XML markup for the layout of the
	 *         Electronic Business Card.
	 */
	public String getBusinessCardLayoutXml() {
		
		return getStringProperty("BusinessCardLayoutXml");
	}
	
	/**
	 * Returns or sets a String that represents the XML markup for the layout of
	 * the Electronic Business Card. Read/write.
	 * <p>
	 * For more information on the XML schema for Electronic Business Cards, see
	 * the Microsoft Outlook 2010 XML Schema Reference in the MSDN Library.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param xml
	 *            a String that represents the XML markup for the layout of the
	 *            Electronic Business Card.
	 */
	public void setBusinessCardLayoutXml(String xml) {
		
		setProperty("BusinessCardLayoutXml", xml);
	}
	
	/**
	 * Returns a BusinessCardType constant that specifies the type of Electronic
	 * Business Card used by this contact. Read-only.
	 * <p>
	 * The Electronic Business Card can be either in Microsoft Office
	 * InterConnect format or Outlook format. An Electronic Business Card in
	 * InterConnect format cannot be modified through the Outlook object model.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a BusinessCardType constant that specifies the type of Electronic
	 *         Business Card used by this contact.
	 */
	public BusinessCardType getBusinessCardType() {
		
		return BusinessCardType.parse(getShortProperty("BusinessCardType"));
	}
	
	/**
	 * Returns or sets a String representing the business fax number for the
	 * contact. Read/write.
	 * 
	 * @return a String representing the business fax number for the contact.
	 */
	public String getBusinessFaxNumber() {
		
		return getStringProperty("BusinessFaxNumber");
	}
	
	/**
	 * Returns or sets a String representing the business fax number for the
	 * contact. Read/write.
	 * 
	 * @param fax
	 *            a String representing the business fax number for the contact.
	 */
	public void setBusinessFaxNumber(String fax) {
		
		setProperty("BusinessFaxNumber", fax);
	}
	
	/**
	 * Returns or sets a String representing the URL of the business Web page
	 * for the contact. Read/write.
	 * 
	 * @return a String representing the URL of the business Web page for the
	 *         contact.
	 */
	public String getBusinessHomePage() {
		
		return getStringProperty("BusinessHomePage");
	}
	
	/**
	 * Returns or sets a String representing the URL of the business Web page
	 * for the contact. Read/write.
	 * 
	 * @param url
	 *            a String representing the URL of the business Web page for the
	 *            contact.
	 */
	public void setBusinessHomePage(String url) {
		
		setProperty("BusinessHomePage", url);
	}
	
	/**
	 * Returns or sets a String representing the first business telephone number
	 * for the contact. Read/write.
	 * 
	 * @return a String representing the first business telephone number for the
	 *         contact.
	 */
	public String getBusinessTelephoneNumber() {
		
		return getStringProperty("BusinessTelephoneNumber");
	}
	
	/**
	 * Returns or sets a String representing the first business telephone number
	 * for the contact. Read/write.
	 * 
	 * @param tel
	 *            a String representing the first business telephone number for
	 *            the contact.
	 */
	public void setBusinessTelephoneNumber(String tel) {
		
		setProperty("BusinessTelephoneNumber", tel);
	}
	
	/**
	 * Returns or sets a String representing the callback telephone number for
	 * the contact. Read/write.
	 * 
	 * @return a String representing the callback telephone number for the
	 *         contact.
	 */
	public String getCallbackTelephoneNumber() {
		
		return getStringProperty("CallbackTelephoneNumber");
	}
	
	/**
	 * Returns or sets a String representing the callback telephone number for
	 * the contact. Read/write.
	 * 
	 * @param tel
	 *            a String representing the callback telephone number for the
	 *            contact.
	 */
	public void setCallbackTelephoneNumber(String tel) {
		
		setProperty("CallbackTelephoneNumber", tel);
	}
	
	/**
	 * Returns or sets a String representing the car telephone number for the
	 * contact. Read/write.
	 * 
	 * @return a String representing the car telephone number for the contact.
	 */
	public String getCarTelephoneNumber() {
		
		return getStringProperty("CarTelephoneNumber");
	}
	
	/**
	 * Returns or sets a String representing the car telephone number for the
	 * contact. Read/write.
	 * 
	 * @param tel
	 *            a String representing the car telephone number for the
	 *            contact.
	 */
	public void setCarTelephoneNumber(String tel) {
		
		setProperty("CarTelephoneNumber", tel);
	}
	
	/**
	 * Returns or sets a String representing the names of the children of the
	 * contact. Read/write.
	 * 
	 * @return a String representing the names of the children of the contact.
	 */
	public String getChildren() {
		
		return getStringProperty("Children");
	}
	
	/**
	 * Returns or sets a String representing the names of the children of the
	 * contact. Read/write.
	 * 
	 * @param kids
	 *            a String representing the names of the children of the
	 *            contact.
	 */
	public void setChildren(String kids) {
		
		setProperty("Children", kids);
	}
	
	/**
	 * Returns a String representing the concatenated company name and full name
	 * for the contact. Read-only.
	 * 
	 * @return a String representing the concatenated company name and full name
	 *         for the contact.
	 */
	public String getCompanyAndFullName() {
		
		return getStringProperty("CompanyAndFullName");
	}
	
	/**
	 * Returns a String representing the company name for the contact followed
	 * by the concatenated last name, first name, and middle name with no space
	 * between the last and first names. Read-only.
	 * <p>
	 * This property is parsed from the CompanyName, LastName, FirstName, and
	 * MiddleName properties. The LastName, FirstName, and MiddleName properties
	 * are themselves parsed from the FullName property. The value of this
	 * property is only filled when its associated property (FirstName,
	 * LastName, MiddleName, CompanyName, and Suffix) contain Asian (DBCS)
	 * characters. If the corresponding field does not contain Asian characters,
	 * the property will be empty.
	 * </p>
	 * 
	 * @return a String representing the company name for the contact followed
	 *         by the concatenated last name, first name, and middle name with
	 *         no space between the last and first names.
	 */
	public String getCompanyLastFirstNoSpace() {
		
		return getStringProperty("CompanyLastFirstNoSpace");
	}
	
	/**
	 * Returns a String representing the company name for the contact followed
	 * by the concatenated last name, first name, and middle name with spaces
	 * between the last, first, and middle names. Read-only.
	 * <p>
	 * This property is parsed from the CompanyName, LastName, FirstName, and
	 * MiddleName properties. The LastName, FirstName, and MiddleName properties
	 * are themselves parsed from the FullName property. The value of this
	 * property is only filled when its associated property (FirstName,
	 * LastName, MiddleName, CompanyName, and Suffix) contain Asian (DBCS)
	 * characters. If the corresponding field does not contain Asian characters,
	 * the property will be empty.
	 * </p>
	 * 
	 * @return a String representing the company name for the contact followed
	 *         by the concatenated last name, first name, and middle name with
	 *         spaces between the last, first, and middle names.
	 */
	public String getCompanyLastFirstSpaceOnly() {
		
		return getStringProperty("CompanyLastFirstSpaceOnly");
	}
	
	/**
	 * Returns or sets a String representing the company main telephone number
	 * for the contact. Read/write.
	 * 
	 * @return a String representing the company main telephone number for the
	 *         contact.
	 */
	public String getCompanyMainTelephoneNumber() {
		
		return getStringProperty("CompanyMainTelephoneNumber");
	}
	
	/**
	 * Returns or sets a String representing the company main telephone number
	 * for the contact. Read/write.
	 * 
	 * @param tel
	 *            a String representing the company main telephone number for
	 *            the contact.
	 */
	public void setCompanyMainTelephoneNumber(String tel) {
		
		setProperty("CompanyMainTelephoneNumber", tel);
	}
	
	/**
	 * Returns or sets a String representing the company name for the contact.
	 * Read/write.
	 * 
	 * @return a String representing the company name for the contact.
	 */
	public String getCompanyName() {
		
		return getStringProperty("CompanyName");
	}
	
	/**
	 * Returns or sets a String representing the company name for the contact.
	 * Read/write.
	 * 
	 * @param name
	 *            a String representing the company name for the contact.
	 */
	public void setCompanyName(String name) {
		
		setProperty("CompanyName", name);
	}
	
	/**
	 * Returns or sets a String representing the name of the computer network
	 * for the contact. Read/write.
	 * 
	 * @return a String representing the name of the computer network for the
	 *         contact.
	 */
	public String getComputerNetworkName() {
		
		return getStringProperty("ComputerNetworkName");
	}
	
	/**
	 * Returns or sets a String representing the name of the computer network
	 * for the contact. Read/write.
	 * 
	 * @param name
	 *            a String representing the name of the computer network for the
	 *            contact.
	 */
	public void setComputerNetworkName(String name) {
		
		setProperty("ComputerNetworkName", name);
	}
	
	/**
	 * Returns or sets a String representing the customer ID for the contact.
	 * Read/write.
	 * 
	 * @return a String representing the customer ID for the contact.
	 */
	public String getCustomerID() {
		
		return getStringProperty("CustomerID");
	}
	
	/**
	 * Returns or sets a String representing the customer ID for the contact.
	 * Read/write.
	 * 
	 * @param id
	 *            a String representing the customer ID for the contact.
	 */
	public void setCustomerID(String id) {
		
		setProperty("CustomerID", id);
	}
	
	/**
	 * Returns or sets a String representing the department name for the
	 * contact. Read/write.
	 * 
	 * @return a String representing the department name for the contact.
	 */
	/**
	 * @return
	 */
	public String getDepartment() {
		
		return getStringProperty("Department");
	}
	
	/**
	 * Returns or sets a String representing the department name for the
	 * contact. Read/write.
	 * 
	 * @param name
	 *            a String representing the department name for the contact.
	 */
	/**
	 * @param name
	 */
	public void setDepartment(String name) {
		
		setProperty("Department", name);
	}
	
	/**
	 * Returns or sets a String representing the e-mail address of the first
	 * e-mail entry for the contact. Read/write.
	 * 
	 * @return a String representing the e-mail address of the first e-mail
	 *         entry for the contact.
	 */
	public String getEmail1Address() {
		
		return getStringProperty("Email1Address");
	}
	
	/**
	 * Returns or sets a String representing the e-mail address of the first
	 * e-mail entry for the contact. Read/write.
	 * 
	 * @param val
	 *            a String representing the e-mail address of the first e-mail
	 *            entry for the contact.
	 */
	public void setEmail1Address(String val) {
		
		setProperty("Email1Address", val);
	}
	
	/**
	 * Returns or sets a String representing the address type (such as EX or
	 * SMTP) of the first e-mail entry for the contact. Read/write.
	 * <p>
	 * This is a free-form text field, but it must match the actual type of an
	 * existing e-mail transport.
	 * </p>
	 * 
	 * @return a String representing the address type (such as EX or SMTP) of
	 *         the first e-mail entry for the contact.
	 */
	public String getEmail1AddressType() {
		
		return getStringProperty("Email1AddressType");
	}
	
	/**
	 * Returns or sets a String representing the address type (such as EX or
	 * SMTP) of the first e-mail entry for the contact. Read/write.
	 * <p>
	 * This is a free-form text field, but it must match the actual type of an
	 * existing e-mail transport.
	 * </p>
	 * 
	 * @param val
	 *            a String representing the address type (such as EX or SMTP) of
	 *            the first e-mail entry for the contact.
	 */
	public void setEmail1AddressType(String val) {
		
		setProperty("Email1AddressType", val);
	}
	
	/**
	 * Returns a String representing the display name of the first e-mail
	 * address for the contact. Read/write.
	 * 
	 * @return a String representing the display name of the first e-mail
	 *         address for the contact.
	 */
	public String getEmail1DisplayName() {
		
		return getStringProperty("Email1DisplayName");
	}
	
	/**
	 * Returns a String representing the display name of the first e-mail
	 * address for the contact. Read/write.
	 * 
	 * @param val
	 *            a String representing the display name of the first e-mail
	 *            address for the contact.
	 */
	public void setEmail1DisplayName(String val) {
		
		setProperty("Email1DisplayName", val);
	}
	
	/**
	 * Returns a String representing the entry ID of the first e-mail address
	 * for the contact. Read-only.
	 * <p>
	 * This property corresponds to the MAPI named property
	 * dispidEmail1OriginalEntryID
	 * </p>
	 * <p>
	 * If you are getting this property in a Microsoft Visual Basic or Microsoft
	 * Visual Basic for Applications (VBA) solution, owing to some type issues,
	 * instead of directly referencing Email1EntryID, you should get the
	 * property through the PropertyAccessor object returned by the
	 * ContactItem.PropertyAccessor property, specifying the MAPI property
	 * PidLidEmail1OriginalEntryId property and its MAPI id namespace. The
	 * following code sample in VBA shows the workaround.
	 * </p>
	 * 
	 * @return a String representing the entry ID of the first e-mail address
	 *         for the contact.
	 */
	public String getEmail1EntryID() {
		
		return getStringProperty("Email1EntryID");
	}
	
	/**
	 * Returns or sets a String representing the e-mail address of the second
	 * e-mail entry for the contact. Read/write.
	 * 
	 * @return a String representing the e-mail address of the second e-mail
	 *         entry for the contact.
	 */
	public String getEmail2Address() {
		
		return getStringProperty("Email2Address");
	}
	
	/**
	 * Returns or sets a String representing the e-mail address of the second
	 * e-mail entry for the contact. Read/write.
	 * 
	 * @param val
	 *            a String representing the e-mail address of the second e-mail
	 *            entry for the contact.
	 */
	public void setEmail2Address(String val) {
		
		setProperty("Email2Address", val);
	}
	
	/**
	 * Returns or sets a String representing the address type (such as EX or
	 * SMTP) of the second e-mail entry for the contact. Read/write.
	 * <p>
	 * This is a free-form text field, but it must match the actual type of an
	 * existing e-mail transport.
	 * </p>
	 * 
	 * @return a String representing the address type (such as EX or SMTP) of
	 *         the second e-mail entry for the contact.
	 */
	public String getEmail2AddressType() {
		
		return getStringProperty("Email2AddressType");
	}
	
	/**
	 * Returns or sets a String representing the address type (such as EX or
	 * SMTP) of the second e-mail entry for the contact. Read/write.
	 * <p>
	 * This is a free-form text field, but it must match the actual type of an
	 * existing e-mail transport.
	 * </p>
	 * 
	 * @param val
	 *            a String representing the address type (such as EX or SMTP) of
	 *            the second e-mail entry for the contact.
	 */
	public void setEmail2AddressType(String val) {
		
		setProperty("Email2AddressType", val);
	}
	
	/**
	 * Returns a String representing the display name of the second e-mail entry
	 * for the contact. Read/write.
	 * <p>
	 * This property is set to the value of the FullName property by default.
	 * </p>
	 * 
	 * @return a String representing the display name of the second e-mail entry
	 *         for the contact.
	 */
	public String getEmail2DisplayName() {
		
		return getStringProperty("Email2DisplayName");
	}
	
	/**
	 * Returns a String representing the display name of the second e-mail entry
	 * for the contact. Read/write.
	 * <p>
	 * This property is set to the value of the FullName property by default.
	 * </p>
	 * 
	 * @param val
	 *            a String representing the display name of the second e-mail
	 *            entry for the contact.
	 */
	public void setEmail2DisplayName(String val) {
		
		setProperty("Email2DisplayName", val);
	}
	
	/**
	 * Returns a String representing the entry ID of the second e-mail entry for
	 * the contact. Read-only.
	 * <p>
	 * This property corresponds to the MAPI named property
	 * dispidEmail2OriginalEntryID.
	 * </p>
	 * <p>
	 * If you are getting this property in a Microsoft Visual Basic or Microsoft
	 * Visual Basic for Applications (VBA) solution, owing to some type issues,
	 * instead of directly referencing Email2EntryID, you should get the
	 * property through the PropertyAccessor object returned by the
	 * ContactItem.PropertyAccessor property, specifying the MAPI property
	 * PidLidEmail2OriginalEntryId property and its MAPI id namespace. The
	 * following code sample in VBA shows the workaround.
	 * </p>
	 * 
	 * @return a String representing the entry ID of the second e-mail entry for
	 *         the contact.
	 */
	public String getEmail2EntryID() {
		
		return getStringProperty("Email1EntryID");
	}
	
	/**
	 * Returns or sets a String representing the e-mail address of the third
	 * e-mail entry for the contact. Read/write.
	 * 
	 * @return a String representing the e-mail address of the third e-mail
	 *         entry for the contact.
	 */
	public String getEmail3Address() {
		
		return getStringProperty("Email3Address");
	}
	
	/**
	 * Returns or sets a String representing the e-mail address of the third
	 * e-mail entry for the contact. Read/write.
	 * 
	 * @param val
	 *            a String representing the e-mail address of the third e-mail
	 *            entry for the contact.
	 */
	public void setEmail3Address(String val) {
		
		setProperty("Email3Address", val);
	}
	
	/**
	 * Returns or sets a String representing the address type (such as EX or
	 * SMTP) of the third e-mail entry for the contact. Read/write.
	 * <p>
	 * This is a free-form text field, but it must match the actual type of an
	 * existing e-mail transport.
	 * </p>
	 * 
	 * @return a String representing the address type (such as EX or SMTP) of
	 *         the third e-mail entry for the contact.
	 */
	public String getEmail3AddressType() {
		
		return getStringProperty("Email3AddressType");
	}
	
	/**
	 * Returns or sets a String representing the address type (such as EX or
	 * SMTP) of the third e-mail entry for the contact. Read/write.
	 * <p>
	 * This is a free-form text field, but it must match the actual type of an
	 * existing e-mail transport.
	 * </p>
	 * 
	 * @param val
	 *            a String representing the address type (such as EX or SMTP) of
	 *            the third e-mail entry for the contact.
	 */
	public void setEmail3AddressType(String val) {
		
		setProperty("Email3AddressType", val);
	}
	
	/**
	 * Returns a String representing the display name of the third e-mail entry
	 * for the contact. Read/write.
	 * <p>
	 * This property is set to the value of the FullName property by default.
	 * </p>
	 * 
	 * @return a String representing the display name of the third e-mail entry
	 *         for the contact.
	 */
	public String getEmail3DisplayName() {
		
		return getStringProperty("Email3DisplayName");
	}
	
	/**
	 * Returns a String representing the display name of the third e-mail entry
	 * for the contact. Read/write.
	 * <p>
	 * This property is set to the value of the FullName property by default.
	 * </p>
	 * 
	 * @param val
	 *            a String representing the display name of the third e-mail
	 *            entry for the contact.
	 */
	public void setEmail3DisplayName(String val) {
		
		setProperty("Email3DisplayName", val);
	}
	
	/**
	 * Returns a String representing the entry ID of the third e-mail entry for
	 * the contact. Read-only.
	 * <p>
	 * This property corresponds to the MAPI named property
	 * dispidEmail3OriginalEntryID.
	 * </p>
	 * <p>
	 * If you are getting this property in a Microsoft Visual Basic or Microsoft
	 * Visual Basic for Applications (VBA) solution, owing to some type issues,
	 * instead of directly referencing Email3EntryID, you should get the
	 * property through the PropertyAccessor object returned by the
	 * ContactItem.PropertyAccessor property, specifying the MAPI property
	 * PidLidEmail3OriginalEntryId property and its MAPI id namespace. The
	 * following code sample in VBA shows the workaround.
	 * </p>
	 * 
	 * @return a String representing the entry ID of the third e-mail entry for
	 *         the contact.
	 */
	public String getEmail3EntryID() {
		
		return getStringProperty("Email1EntryID");
	}
	
	/**
	 * Returns or sets a String indicating the default keyword string assigned
	 * to the contact when it is filed. Read/write.
	 * 
	 * @return a String indicating the default keyword string assigned to the
	 *         contact when it is filed.
	 */
	public String getFileAs() {
		
		return getStringProperty("FileAs");
	}
	
	/**
	 * Returns or sets a String indicating the default keyword string assigned
	 * to the contact when it is filed. Read/write.
	 * 
	 * @param val
	 *            a String indicating the default keyword string assigned to the
	 *            contact when it is filed.
	 */
	public void setFileAs(String val) {
		
		setProperty("FileAs", val);
	}
	
	/**
	 * Returns or sets a String representing the first name for the contact.
	 * Read/write.
	 * <p>
	 * This property is parsed from the FullName property, but may be changed or
	 * entered independently should it be parsed incorrectly. Note that any such
	 * changes or entries to this property will be overwritten by any subsequent
	 * changes of entries to FullName.
	 * </p>
	 * 
	 * @return a String representing the first name for the contact.
	 */
	public String getFirstName() {
		
		return getStringProperty("FirstName");
	}
	
	/**
	 * Returns or sets a String representing the first name for the contact.
	 * Read/write.
	 * <p>
	 * This property is parsed from the FullName property, but may be changed or
	 * entered independently should it be parsed incorrectly. Note that any such
	 * changes or entries to this property will be overwritten by any subsequent
	 * changes of entries to FullName.
	 * </p>
	 * 
	 * @param name a String representing the first name for the contact.
	 */
	public void setFirstName(String name) {
		
		setProperty("FirstName", name);
	}
	
	/**
	 * Creates a new MailItem object containing contact information and,
	 * optionally, an Electronic Business Card (EBC) image based on the
	 * specified ContactItem object.
	 * <p>
	 * This method creates a new Outlook mail item based on the information
	 * stored in the ContactItem object. The information included in the Outlook
	 * mail item depends on the value of the BodyFormat property for the
	 * MailItem object:
	 * <ul>
	 * <li>olFormatPlain - A vCard (.vcf) file is created and added to the
	 * Attachments collection of the MailItem object.</li>
	 * 
	 * <li>olFormatRichText - A vCard file is created and added to the
	 * Attachments collection of the MailItem object.</li>
	 * 
	 * <li>olFormatHTML - An image of the Electronic Business Card is generated
	 * and included in the Body property of the MailItem object, and a vCard
	 * file is created and added to the Attachments collection of the MailItem
	 * object.</li>
	 * </ul>
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return a new MailItem object containing contact information and,
	 *         optionally, an Electronic Business Card (EBC) image based on the
	 *         specified ContactItem object.
	 */
	public MailItem forwardAsBusinessCard() {
		
		return new MailItem((IDispatch) invoke("ForwardAsBusinessCard").getValue());
	}
	
	/**
	 * Creates a MailItem and attaches the contact information in vCard format.
	 * <p>
	 * A MailItem object that represents the new mail item to which the contact
	 * information is attached.
	 * </p>
	 * </p>
	 * 
	 * @return a MailItem and attaches the contact information in vCard format.
	 */
	public MailItem forwardAsVcard() {
		
		return new MailItem((IDispatch) invoke("ForwardAsVcard").getValue());
	}
	
	/**
	 * Returns or sets a String representing the FTP site entry for the contact.
	 * Read/write.
	 * 
	 * @return a String representing the FTP site entry for the contact.
	 */
	public String getFTPSite() {
		
		return getStringProperty("FTPSite");
	}
	
	/**
	 * Returns or sets a String representing the FTP site entry for the contact.
	 * Read/write.
	 * 
	 * @param site
	 *            a String representing the FTP site entry for the contact.
	 */
	public void setFTPSite(String site) {
		
		setProperty("FTPSite", site);
	}
	
	/**
	 * Returns or sets a String specifying the whole, unparsed full name for the
	 * contact. Read/write.
	 * <p>
	 * This property is parsed into the FirstName , MiddleName , LastName, and
	 * Suffix properties, which may be changed or typed independently if they
	 * are parsed incorrectly. Any changes or entries to the FirstName,
	 * LastName, MiddleName, or Suffix properties will be overwritten by any
	 * subsequent changes or entries to FullName.
	 * </p>
	 * 
	 * @return a String specifying the whole, unparsed full name for the
	 *         contact.
	 */
	public String getFullName() {
		
		return getStringProperty("FullName");
	}
	
	/**
	 * Returns or sets a String specifying the whole, unparsed full name for the
	 * contact. Read/write.
	 * <p>
	 * This property is parsed into the FirstName , MiddleName , LastName, and
	 * Suffix properties, which may be changed or typed independently if they
	 * are parsed incorrectly. Any changes or entries to the FirstName,
	 * LastName, MiddleName, or Suffix properties will be overwritten by any
	 * subsequent changes or entries to FullName.
	 * </p>
	 * 
	 * @param name
	 *            a String specifying the whole, unparsed full name for the
	 *            contact.
	 */
	public void setFullName(String name) {
		
		setProperty("FullName", name);
	}
	
	/**
	 * Returns a String representing the full name and company of the contact by
	 * concatenating the values of the FullName and CompanyName properties.
	 * Read-only.
	 * 
	 * @return a String representing the full name and company of the contact by
	 *         concatenating the values of the FullName and CompanyName
	 *         properties.
	 */
	public String getFullNameAndCompany() {
		
		return getStringProperty("FullNameAndCompany");
	}
	
	/**
	 * Returns or sets a Gender constant indicating the gender of the contact.
	 * Read/write.
	 * 
	 * @return a Gender constant indicating the gender of the contact.
	 */
	public Gender getGender() {
		
		return Gender.parse(getShortProperty("Gender"));
	}
	
	/**
	 * Returns or sets a Gender constant indicating the gender of the contact.
	 * Read/write.
	 * 
	 * @param sex
	 *            a Gender constant indicating the gender of the contact.
	 */
	public void setGender(Gender sex) {
		
		setProperty("Gender", sex.value());
	}
	
	/**
	 * Returns or sets a String representing the government ID number for the
	 * contact. Read/write.
	 * 
	 * @return a String representing the government ID number for the contact.
	 */
	public String getGovernmentIDNumber() {
		
		return getStringProperty("GovernmentIDNumber");
	}
	
	/**
	 * Returns or sets a String representing the government ID number for the
	 * contact. Read/write.
	 * 
	 * @param id
	 *            a String representing the government ID number for the
	 *            contact.
	 */
	public void setGovernmentIDNumber(String id) {
		
		setProperty("GovernmentIDNumber", id);
	}
	
	/**
	 * Returns true if a Contacts item has a picture associated with it.
	 * Read-only
	 * 
	 * @return true if a Contacts item has a picture associated with it.
	 */
	public boolean hasPicture() {
		
		return getBooleanProperty("HasPicture");
	}
	
	/**
	 * Returns or sets a String representing the hobby for the contact.
	 * Read/write.
	 * 
	 * @return a String representing the hobby for the contact.
	 */
	public String getHobby() {
		
		return getStringProperty("Hobby");
	}
	
	/**
	 * Returns or sets a String representing the hobby for the contact.
	 * Read/write.
	 * 
	 * @param hobby
	 *            a String representing the hobby for the contact.
	 */
	public void setHobby(String hobby) {
		
		setProperty("Hobby", hobby);
	}
	
	/**
	 * Returns or sets a String representing the full, unparsed text of the home
	 * address for the contact. Read/write.
	 * 
	 * @return a String representing the full, unparsed text of the home address
	 *         for the contact.
	 */
	public String getHomeAddress() {
		
		return getStringProperty("HomeAddress");
	}
	
	/**
	 * Returns or sets a String representing the full, unparsed text of the home
	 * address for the contact. Read/write.
	 * 
	 * @param addr
	 *            a String representing the full, unparsed text of the home
	 *            address for the contact.
	 */
	public void setHomeAddress(String addr) {
		
		setProperty("HomeAddress", addr);
	}
	
	/**
	 * Returns or sets a String representing the city portion of the home
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the HomeAddress property, but may be changed
	 * or entered independently should it be parsed incorrectly. Note that any
	 * such changes or entries to this property will be overwritten by any
	 * subsequent changes or entries to HomeAddress.
	 * </p>
	 * 
	 * @return a String representing the city portion of the home address for
	 *         the contact.
	 */
	public String getHomeAddressCity() {
		
		return getStringProperty("HomeAddressCity");
	}
	
	/**
	 * Returns or sets a String representing the city portion of the home
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the HomeAddress property, but may be changed
	 * or entered independently should it be parsed incorrectly. Note that any
	 * such changes or entries to this property will be overwritten by any
	 * subsequent changes or entries to HomeAddress.
	 * </p>
	 * 
	 * @param city
	 *            a String representing the city portion of the home address for
	 *            the contact.
	 */
	public void setHomeAddressCity(String city) {
		
		setProperty("HomeAddressCity", city);
	}
	
	/**
	 * Returns or sets a String representing the country/region portion of the
	 * home address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the HomeAddress property, but may be changed
	 * or entered independently should it be parsed incorrectly. Note that any
	 * such changes or entries to this property will be overwritten by any
	 * subsequent changes or entries to HomeAddress.
	 * </p>
	 * 
	 * @return a String representing the country/region portion of the home
	 *         address for the contact.
	 */
	public String getHomeAddressCountry() {
		
		return getStringProperty("HomeAddressCountry");
	}
	
	/**
	 * Returns or sets a String representing the country/region portion of the
	 * home address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the HomeAddress property, but may be changed
	 * or entered independently should it be parsed incorrectly. Note that any
	 * such changes or entries to this property will be overwritten by any
	 * subsequent changes or entries to HomeAddress.
	 * </p>
	 * 
	 * @param country
	 *            a String representing the country/region portion of the home
	 *            address for the contact.
	 */
	public void setHomeAddressCountry(String country) {
		
		setProperty("HomeAddressCountry", country);
	}
	
	/**
	 * Returns or sets a String representing the postal code portion of the home
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the HomeAddress property, but may be changed
	 * or entered independently should it be parsed incorrectly. Note that any
	 * such changes or entries to this property will be overwritten by any
	 * subsequent changes or entries to HomeAddress.
	 * </p>
	 * 
	 * @return a String representing the postal code portion of the home address
	 *         for the contact.
	 */
	public String getHomeAddressPostalCode() {
		
		return getStringProperty("HomeAddressPostalCode");
	}
	
	/**
	 * Returns or sets a String representing the postal code portion of the home
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the HomeAddress property, but may be changed
	 * or entered independently should it be parsed incorrectly. Note that any
	 * such changes or entries to this property will be overwritten by any
	 * subsequent changes or entries to HomeAddress.
	 * </p>
	 * 
	 * @param code
	 *            a String representing the postal code portion of the home
	 *            address for the contact.
	 */
	public void setHomeAddressPostalCode(String code) {
		
		setProperty("HomeAddressPostalCode", code);
	}
	
	/**
	 * Returns or sets a String the post office box number portion of the home
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the HomeAddress property, but may be changed
	 * or entered independently should it be parsed incorrectly. Note that any
	 * such changes or entries to this property will be overwritten by any
	 * subsequent changes or entries to HomeAddress.
	 * </p>
	 * 
	 * @return a String the post office box number portion of the home address
	 *         for the contact.
	 */
	public String getHomeAddressPostOfficeBox() {
		
		return getStringProperty("HomeAddressPostOfficeBox");
	}
	
	/**
	 * Returns or sets a String the post office box number portion of the home
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the HomeAddress property, but may be changed
	 * or entered independently should it be parsed incorrectly. Note that any
	 * such changes or entries to this property will be overwritten by any
	 * subsequent changes or entries to HomeAddress.
	 * </p>
	 * 
	 * @param code
	 *            a String the post office box number portion of the home
	 *            address for the contact.
	 */
	public void setHomeAddressPostOfficeBox(String code) {
		
		setProperty("HomeAddressPostOfficeBox", code);
	}
	
	/**
	 * Returns or sets a String representing the state portion of the home
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the HomeAddress property, but may be changed
	 * or entered independently should it be parsed incorrectly. Note that any
	 * such changes or entries to this property will be overwritten by any
	 * subsequent changes or entries to HomeAddress.
	 * </p>
	 * 
	 * @return a String representing the state portion of the home address for
	 *         the contact.
	 */
	public String getHomeAddressState() {
		
		return getStringProperty("HomeAddressState");
	}
	
	/**
	 * Returns or sets a String representing the state portion of the home
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the HomeAddress property, but may be changed
	 * or entered independently should it be parsed incorrectly. Note that any
	 * such changes or entries to this property will be overwritten by any
	 * subsequent changes or entries to HomeAddress.
	 * </p>
	 * 
	 * @param state
	 *            a String representing the state portion of the home address
	 *            for the contact.
	 */
	public void setHomeAddressState(String state) {
		
		setProperty("HomeAddressState", state);
	}
	
	/**
	 * Returns or sets a String representing the street portion of the home
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the HomeAddress property, but may be changed
	 * or entered independently should it be parsed incorrectly. Note that any
	 * such changes or entries to this property will be overwritten by any
	 * subsequent changes or entries to HomeAddress.
	 * </p>
	 * 
	 * @return a String representing the street portion of the home address for
	 *         the contact.
	 */
	public String getHomeAddressStreet() {
		
		return getStringProperty("HomeAddressStreet");
	}
	
	/**
	 * Returns or sets a String representing the street portion of the home
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the HomeAddress property, but may be changed
	 * or entered independently should it be parsed incorrectly. Note that any
	 * such changes or entries to this property will be overwritten by any
	 * subsequent changes or entries to HomeAddress.
	 * </p>
	 * 
	 * @param street
	 *            a String representing the street portion of the home address
	 *            for the contact.
	 */
	public void setHomeAddressStreet(String street) {
		
		setProperty("HomeAddressStreet", street);
	}
	
	/**
	 * Returns or sets a String representing the home fax number for the
	 * contact. Read/write.
	 * 
	 * @return a String representing the home fax number for the contact.
	 */
	public String getHomeFaxNumber() {
		
		return getStringProperty("HomeFaxNumber");
	}
	
	/**
	 * Returns or sets a String representing the home fax number for the
	 * contact. Read/write.
	 * 
	 * @param num
	 *            a String representing the home fax number for the contact.
	 */
	public void setHomeFaxNumber(String num) {
		
		setProperty("HomeFaxNumber", num);
	}
	
	/**
	 * Returns or sets a String representing the first home telephone number for
	 * the contact. Read/write.
	 * 
	 * @return a String representing the first home telephone number for the
	 *         contact.
	 */
	public String getHomeTelephoneNumber() {
		
		return getStringProperty("HomeTelephoneNumber");
	}
	
	/**
	 * Returns or sets a String representing the first home telephone number for
	 * the contact. Read/write.
	 * 
	 * @param num
	 *            a String representing the first home telephone number for the
	 *            contact.
	 */
	public void setHomeTelephoneNumber(String num) {
		
		setProperty("HomeTelephoneNumber", num);
	}
	
	/**
	 * Returns or sets a String that represents a contact's Microsoft Instant
	 * Messenger address. Read/write.
	 * <p>
	 * Unlike the Recipients or To properties, there is no way to verify that
	 * the IMAddress property contains a valid address.
	 * </p>
	 * 
	 * @return a String that represents a contact's Microsoft Instant Messenger
	 *         address.
	 */
	public String getIMAddress() {
		
		return getStringProperty("IMAddress");
	}
	
	/**
	 * Returns or sets a String that represents a contact's Microsoft Instant
	 * Messenger address. Read/write.
	 * <p>
	 * Unlike the Recipients or To properties, there is no way to verify that
	 * the IMAddress property contains a valid address.
	 * </p>
	 * 
	 * @param addr
	 *            a String that represents a contact's Microsoft Instant
	 *            Messenger address.
	 */
	public void setIMAddress(String addr) {
		
		setProperty("IMAddress", addr);
	}
	
	/**
	 * Returns or sets a String representing the initials for the contact.
	 * Read/write.
	 * 
	 * @return a String representing the initials for the contact.
	 */
	public String getInitials() {
		
		return getStringProperty("Initials");
	}
	
	/**
	 * Returns or sets a String representing the initials for the contact.
	 * Read/write.
	 * 
	 * @param initials
	 *            a String representing the initials for the contact.
	 */
	public void setInitials(String initials) {
		
		setProperty("Initials", initials);
	}
	
	/**
	 * Returns or sets a String corresponding to the Address box on the Details
	 * tab for a contact. Read/write.
	 * <p>
	 * The Address box on the Details tab can contain the URL location of the
	 * user's free-busy information in vCard Free-Busy standard format.
	 * </p>
	 * 
	 * @return a String corresponding to the Address box on the Details tab for
	 *         a contact.
	 */
	public String getInternetFreeBusyAddress() {
		
		return getStringProperty("InternetFreeBusyAddress");
	}
	
	/**
	 * Returns or sets a String corresponding to the Address box on the Details
	 * tab for a contact. Read/write.
	 * <p>
	 * The Address box on the Details tab can contain the URL location of the
	 * user's free-busy information in vCard Free-Busy standard format.
	 * </p>
	 * 
	 * @param addr
	 *            a String corresponding to the Address box on the Details tab
	 *            for a contact.
	 */
	public void setInternetFreeBusyAddress(String addr) {
		
		setProperty("InternetFreeBusyAddress", addr);
	}
	
	/**
	 * Returns or sets a String representing the ISDN number for the contact.
	 * Read/write.
	 * 
	 * @return a String representing the ISDN number for the contact.
	 */
	public String getISDNNumber() {
		
		return getStringProperty("ISDNNumber");
	}
	
	/**
	 * Returns or sets a String representing the ISDN number for the contact.
	 * Read/write.
	 * 
	 * @param num
	 *            a String representing the ISDN number for the contact.
	 */
	public void setISDNNumber(String num) {
		
		setProperty("ISDNNumber", num);
	}
	
	/**
	 * Returns or sets a String representing the job title for the contact.
	 * Read/write.
	 * 
	 * @return a String representing the job title for the contact.
	 */
	public String getJobTitle() {
		
		return getStringProperty("JobTitle");
	}
	
	/**
	 * Returns or sets a String representing the job title for the contact.
	 * Read/write.
	 * 
	 * @param title
	 *            a String representing the job title for the contact.
	 */
	public void setJobTitle(String title) {
		
		setProperty("JobTitle", title);
	}
	
	/**
	 * Returns a boolean that indicates true if the transaction of the contact
	 * will be journalised. Read/write.
	 * <p>
	 * The default value is false.
	 * </p>
	 * 
	 * @return a boolean that indicates true if the transaction of the contact
	 *         will be journalised.
	 */
	public boolean hasJournal() {
		
		return getBooleanProperty("Journal");
	}
	
	/**
	 * Returns a boolean that indicates true if the transaction of the contact
	 * will be journalised. Read/write.
	 * <p>
	 * The default value is false.
	 * </p>
	 * 
	 * @param flag
	 *            a boolean that indicates true if the transaction of the
	 *            contact will be journalised.
	 */
	public void setJournal(boolean flag) {
		
		setProperty("Journal", flag);
	}
	
	/**
	 * Returns or sets a String that represents the language in which the
	 * contact writes messages. Read/write.
	 * <p>
	 * This property corresponds to the MAPI property PidTagLanguage.
	 * </p>
	 * 
	 * @return a String that represents the language in which the contact writes
	 *         messages.
	 */
	public String getLanguage() {
		
		return getStringProperty("Language");
	}
	
	/**
	 * Returns or sets a String that represents the language in which the
	 * contact writes messages. Read/write.
	 * <p>
	 * This property corresponds to the MAPI property PidTagLanguage.
	 * </p>
	 * 
	 * @param language
	 *            a String that represents the language in which the contact
	 *            writes messages.
	 */
	public void setLanguage(String language) {
		
		setProperty("Language", language);
	}
	
	/**
	 * Returns a String representing the last name, first name, middle name, and
	 * suffix of the contact. Read-only.
	 * <p>
	 * There is a comma between the last and first names and spaces between all
	 * the names and the suffix. This property is parsed from the LastName,
	 * FirstName, MiddleName and Suffix properties. The LastName, FirstName, and
	 * Suffix properties are themselves parsed from the FullName property. The
	 * value of this property is only filled when its associated property
	 * (FirstName, LastName, MiddleName, CompanyName, and Suffix) contain Asian
	 * (DBCS) characters. If the corresponding field does not contain Asian
	 * characters, the property will be empty.
	 * </p>
	 * 
	 * @return a String representing the last name, first name, middle name, and
	 *         suffix of the contact.
	 */
	public String getLastFirstAndSuffix() {
		
		return getStringProperty("LastFirstAndSuffix");
	}
	
	/**
	 * Returns a String representing the concatenated last name, first name, and
	 * middle name of the contact with no space between the last name and the
	 * first name. Read-only.
	 * <p>
	 * This property is parsed from the LastName, FirstName, and MiddleName
	 * properties. The LastName, FirstName, and MiddleName properties are
	 * themselves parsed from the FullName property. The value of this property
	 * is only filled when its associated property (FirstName, LastName,
	 * MiddleName, CompanyName, and Suffix) contain Asian (DBCS) characters. If
	 * the corresponding field does not contain Asian characters, the property
	 * will be empty.
	 * </p>
	 * 
	 * @return a String representing the concatenated last name, first name, and
	 *         middle name of the contact with no space between the last name
	 *         and the first name.
	 */
	public String getLastFirstNoSpace() {
		
		return getStringProperty("LastFirstNoSpace");
	}
	
	/**
	 * Returns a String that contains the last name, first name, and suffix of
	 * the user without a space. Read-only
	 * <p>
	 * This property is used only when the FirstName, LastName, and Suffix
	 * properties (the fields that define this property) contain Asian (DBCS)
	 * characters. Note that any such changes or entries to the FirstName,
	 * LastName, or Suffix properties will be overwritten by any subsequent
	 * changes or entries to FullName.
	 * </p>
	 * 
	 * @return a String that contains the last name, first name, and suffix of
	 *         the user without a space.
	 */
	public String getLastFirstNoSpaceAndSuffix() {
		
		return getStringProperty("LastFirstNoSpaceAndSuffix");
	}
	
	/**
	 * Returns a String representing the concatenated last name, first name, and
	 * middle name of the contact with no space between the last name and the
	 * first name. Read-only.
	 * <p>
	 * The company name for the contact is included after the middle name. This
	 * property is parsed from the CompanyName, LastName, FirstName, and
	 * MiddleName properties. The LastName, FirstName, and MiddleName properties
	 * are themselves parsed from the FullName property. The value of this
	 * property is only filled when its associated property (FirstName,
	 * LastName, MiddleName, CompanyName, and Suffix) contain Asian (DBCS)
	 * characters. If the corresponding field does not contain Asian characters,
	 * the property will be empty.
	 * </p>
	 * 
	 * @return a String representing the concatenated last name, first name, and
	 *         middle name of the contact with no space between the last name
	 *         and the first name.
	 */
	public String getLastFirstNoSpaceCompany() {
		
		return getStringProperty("LastFirstNoSpaceCompany");
	}
	
	/**
	 * Returns a String representing the concatenated last name, first name, and
	 * middle name of the contact with spaces between them. Read-only.
	 * <p>
	 * This property is parsed from the LastName, FirstName, and MiddleName
	 * properties. The LastName, FirstName, and MiddleName properties are
	 * themselves parsed from the FullName property. The value of this property
	 * is only filled when its associated property (FirstName, LastName,
	 * MiddleName, CompanyName, and Suffix) contain Asian (DBCS) characters. If
	 * the corresponding field does not contain Asian characters, the property
	 * will be empty.
	 * </p>
	 * 
	 * @return a String representing the concatenated last name, first name, and
	 *         middle name of the contact with spaces between them.
	 */
	public String getLastFirstSpaceOnly() {
		
		return getStringProperty("LastFirstSpaceOnly");
	}
	
	/**
	 * Returns a String representing the concatenated last name, first name, and
	 * middle name of the contact with spaces between them. Read-only.
	 * <p>
	 * The company name for the contact is included after the middle name. This
	 * property is parsed from the CompanyName, LastName, FirstName, and
	 * MiddleName properties. The LastName, FirstName, and MiddleName properties
	 * are themselves parsed from the FullName property. The value of this
	 * property is only filled when its associated property (FirstName,
	 * LastName, MiddleName, CompanyName, and Suffix) contain Asian (DBCS)
	 * characters. If the corresponding field does not contain Asian characters,
	 * the property will be empty.
	 * </p>
	 * 
	 * @return a String representing the concatenated last name, first name, and
	 *         middle name of the contact with spaces between them.
	 */
	public String getLastFirstSpaceOnlyCompany() {
		
		return getStringProperty("LastFirstSpaceOnlyCompany");
	}
	
	/**
	 * Returns or sets a String representing the last name for the contact.
	 * Read/write.
	 * <p>
	 * This property is parsed from the FullName property, but may be changed or
	 * entered independently should it be parsed incorrectly. Note that any such
	 * changes or entries to this property will be overwritten by any subsequent
	 * changes of entries to FullName.
	 * </p>
	 * 
	 * @return a String representing the last name for the contact.
	 */
	public String getLastName() {
		
		return getStringProperty("LastName");
	}
	
	/**
	 * Returns or sets a String representing the last name for the contact.
	 * Read/write.
	 * <p>
	 * This property is parsed from the FullName property, but may be changed or
	 * entered independently should it be parsed incorrectly. Note that any such
	 * changes or entries to this property will be overwritten by any subsequent
	 * changes of entries to FullName.
	 * </p>
	 * 
	 * @param name
	 *            a String representing the last name for the contact.
	 */
	public void setLastName(String name) {
		
		setProperty("LastName", name);
	}
	
	/**
	 * Returns a String representing the concatenated last name and first name
	 * for the contact. Read-only.
	 * <p>
	 * This property is parsed from the FirstName and LastName properties for
	 * the contact, which are themselves parsed from the FullName property.
	 * </p>
	 * 
	 * @return a String representing the concatenated last name and first name
	 *         for the contact.
	 */
	public String getLastNameAndFirstName() {
		
		return getStringProperty("LastNameAndFirstName");
	}
	
	/**
	 * Returns or sets a String representing the full, unparsed selected mailing
	 * address for the contact. Read/write.
	 * <p>
	 * This property replicates the property indicated by the
	 * SelectedMailingAddress property, which is one of the following
	 * OlMailingAddress constants: olBusiness, olHome, olNone, or olOther. While
	 * it can be changed or entered independently, any such changes or entries
	 * to this property will be overwritten by any subsequent changes or entries
	 * to the property indicated by SelectedMailingAddress.
	 * </p>
	 * 
	 * @return a String representing the full, unparsed selected mailing address
	 *         for the contact.
	 */
	public String getMailingAddress() {
		
		return getStringProperty("MailingAddress");
	}
	
	/**
	 * Returns or sets a String representing the full, unparsed selected mailing
	 * address for the contact. Read/write.
	 * <p>
	 * This property replicates the property indicated by the
	 * SelectedMailingAddress property, which is one of the following
	 * OlMailingAddress constants: olBusiness, olHome, olNone, or olOther. While
	 * it can be changed or entered independently, any such changes or entries
	 * to this property will be overwritten by any subsequent changes or entries
	 * to the property indicated by SelectedMailingAddress.
	 * </p>
	 * 
	 * @param addr
	 *            a String representing the full, unparsed selected mailing
	 *            address for the contact.
	 */
	public void setMailingAddress(String addr) {
		
		setProperty("MailingAddress", addr);
	}
	
	/**
	 * Returns or sets a String representing the city name portion of the
	 * selected mailing address of the contact. Read/write.
	 * <p>
	 * This property replicates the property indicated by the
	 * SelectedMailingAddress property, which is one of the following
	 * OlMailingAddress constants: olBusiness, olHome, olNone, or olOther. While
	 * it can be changed or entered independently, any such changes or entries
	 * to this property will be overwritten by any subsequent changes or entries
	 * to the property indicated by SelectedMailingAddress.
	 * </p>
	 * 
	 * @return a String representing the city name portion of the selected
	 *         mailing address of the contact.
	 */
	public String getMailingAddressCity() {
		
		return getStringProperty("MailingAddressCity");
	}
	
	/**
	 * Returns or sets a String representing the city name portion of the
	 * selected mailing address of the contact. Read/write.
	 * <p>
	 * This property replicates the property indicated by the
	 * SelectedMailingAddress property, which is one of the following
	 * OlMailingAddress constants: olBusiness, olHome, olNone, or olOther. While
	 * it can be changed or entered independently, any such changes or entries
	 * to this property will be overwritten by any subsequent changes or entries
	 * to the property indicated by SelectedMailingAddress.
	 * </p>
	 * 
	 * @param city
	 *            a String representing the city name portion of the selected
	 *            mailing address of the contact.
	 */
	public void setMailingAddressCity(String city) {
		
		setProperty("MailingAddressCity", city);
	}
	
	/**
	 * Returns or sets a String representing the country/region code portion of
	 * the selected mailing address of the contact. Read/write.
	 * <p>
	 * This property replicates the property indicated by the
	 * SelectedMailingAddress property, which is one of the following
	 * OlMailingAddress constants: olBusiness, olHome, olNone, or olOther. While
	 * it can be changed or entered independently, any such changes or entries
	 * to this property will be overwritten by any subsequent changes or entries
	 * to the property indicated by SelectedMailingAddress.
	 * </p>
	 * 
	 * @return a String representing the country/region code portion of the
	 *         selected mailing address of the contact.
	 */
	public String getMailingAddressCountry() {
		
		return getStringProperty("MailingAddressCountry");
	}
	
	/**
	 * Returns or sets a String representing the country/region code portion of
	 * the selected mailing address of the contact. Read/write.
	 * <p>
	 * This property replicates the property indicated by the
	 * SelectedMailingAddress property, which is one of the following
	 * OlMailingAddress constants: olBusiness, olHome, olNone, or olOther. While
	 * it can be changed or entered independently, any such changes or entries
	 * to this property will be overwritten by any subsequent changes or entries
	 * to the property indicated by SelectedMailingAddress.
	 * </p>
	 * 
	 * @param country
	 *            a String representing the country/region code portion of the
	 *            selected mailing address of the contact.
	 */
	public void setMailingAddressCountry(String country) {
		
		setProperty("MailingAddressCountry", country);
	}
	
	/**
	 * Returns or sets a String representing the postal code (zip code) portion
	 * of the selected mailing address of the contact. Read/write.
	 * <p>
	 * This property replicates the property indicated by the
	 * SelectedMailingAddress property, which is one of the following
	 * OlMailingAddress constants: olBusiness, olHome, olNone, or olOther. While
	 * it can be changed or entered independently, any such changes or entries
	 * to this property will be overwritten by any subsequent changes or entries
	 * to the property indicated by SelectedMailingAddress.
	 * </p>
	 * 
	 * @return a String representing the postal code (zip code) portion of the
	 *         selected mailing address of the contact.
	 */
	public String getMailingAddressPostalCode() {
		
		return getStringProperty("MailingAddressPostalCode");
	}
	
	/**
	 * Returns or sets a String representing the postal code (zip code) portion
	 * of the selected mailing address of the contact. Read/write.
	 * <p>
	 * This property replicates the property indicated by the
	 * SelectedMailingAddress property, which is one of the following
	 * OlMailingAddress constants: olBusiness, olHome, olNone, or olOther. While
	 * it can be changed or entered independently, any such changes or entries
	 * to this property will be overwritten by any subsequent changes or entries
	 * to the property indicated by SelectedMailingAddress.
	 * </p>
	 * 
	 * @param code
	 *            a String representing the postal code (zip code) portion of
	 *            the selected mailing address of the contact.
	 */
	public void setMailingAddressPostalCode(String code) {
		
		setProperty("MailingAddressPostalCode", code);
	}
	
	/**
	 * Returns or sets a String representing the post office box number portion
	 * of the selected mailing address of the contact. Read/write.
	 * <p>
	 * This property replicates the property indicated by the
	 * SelectedMailingAddress property, which is one of the following
	 * OlMailingAddress constants: olBusiness, olHome, olNone, or olOther. While
	 * it can be changed or entered independently, any such changes or entries
	 * to this property will be overwritten by any subsequent changes or entries
	 * to the property indicated by SelectedMailingAddress.
	 * </p>
	 * 
	 * @return a String representing the post office box number portion of the
	 *         selected mailing address of the contact.
	 */
	public String getMailingAddressPostOfficeBox() {
		
		return getStringProperty("MailingAddressPostOfficeBox");
	}
	
	/**
	 * Returns or sets a String representing the post office box number portion
	 * of the selected mailing address of the contact. Read/write.
	 * <p>
	 * This property replicates the property indicated by the
	 * SelectedMailingAddress property, which is one of the following
	 * OlMailingAddress constants: olBusiness, olHome, olNone, or olOther. While
	 * it can be changed or entered independently, any such changes or entries
	 * to this property will be overwritten by any subsequent changes or entries
	 * to the property indicated by SelectedMailingAddress.
	 * </p>
	 * 
	 * @param box
	 *            a String representing the post office box number portion of
	 *            the selected mailing address of the contact.
	 */
	public void setMailingAddressPostOfficeBox(String box) {
		
		setProperty("MailingAddressPostOfficeBox", box);
	}
	
	/**
	 * Returns or sets a String representing the state code portion for the
	 * selected mailing address of the contact. Read/write.
	 * <p>
	 * This property replicates the property indicated by the
	 * SelectedMailingAddress property, which is one of the following
	 * OlMailingAddress constants: olBusiness, olHome, olNone, or olOther. While
	 * it can be changed or entered independently, any such changes or entries
	 * to this property will be overwritten by any subsequent changes or entries
	 * to the property indicated by SelectedMailingAddress.
	 * </p>
	 * 
	 * @return a String representing the state code portion for the selected
	 *         mailing address of the contact.
	 */
	public String getMailingAddressState() {
		
		return getStringProperty("MailingAddressState");
	}
	
	/**
	 * Returns or sets a String representing the state code portion for the
	 * selected mailing address of the contact. Read/write.
	 * <p>
	 * This property replicates the property indicated by the
	 * SelectedMailingAddress property, which is one of the following
	 * OlMailingAddress constants: olBusiness, olHome, olNone, or olOther. While
	 * it can be changed or entered independently, any such changes or entries
	 * to this property will be overwritten by any subsequent changes or entries
	 * to the property indicated by SelectedMailingAddress.
	 * </p>
	 * 
	 * @param state
	 *            a String representing the state code portion for the selected
	 *            mailing address of the contact.
	 */
	public void setMailingAddressState(String state) {
		
		setProperty("MailingAddressState", state);
	}
	
	/**
	 * Returns or sets a String representing the street address portion of the
	 * selected mailing address of the contact. Read/write.
	 * <p>
	 * This property replicates the property indicated by the
	 * SelectedMailingAddress property, which is one of the following
	 * OlMailingAddress constants: olBusiness, olHome, olNone, or olOther. While
	 * it can be changed or entered independently, any such changes or entries
	 * to this property will be overwritten by any subsequent changes or entries
	 * to the property indicated by SelectedMailingAddress.
	 * </p>
	 * 
	 * @return a String representing the street address portion of the selected
	 *         mailing address of the contact.
	 */
	public String getMailingAddressStreet() {
		
		return getStringProperty("MailingAddressStreet");
	}
	
	/**
	 * Returns or sets a String representing the street address portion of the
	 * selected mailing address of the contact. Read/write.
	 * <p>
	 * This property replicates the property indicated by the
	 * SelectedMailingAddress property, which is one of the following
	 * OlMailingAddress constants: olBusiness, olHome, olNone, or olOther. While
	 * it can be changed or entered independently, any such changes or entries
	 * to this property will be overwritten by any subsequent changes or entries
	 * to the property indicated by SelectedMailingAddress.
	 * </p>
	 * 
	 * @param street
	 *            a String representing the street address portion of the
	 *            selected mailing address of the contact.
	 */
	public void setMailingAddressStreet(String street) {
		
		setProperty("MailingAddressStreet", street);
	}
	
	/**
	 * Returns or sets a String representing the manager name for the contact.
	 * Read/write.
	 * 
	 * @return a String representing the manager name for the contact.
	 */
	public String getManagerName() {
		
		return getStringProperty("ManagerName");
	}
	
	/**
	 * Returns or sets a String representing the manager name for the contact.
	 * Read/write.
	 * 
	 * @param name
	 *            a String representing the manager name for the contact.
	 */
	public void setManagerName(String name) {
		
		setProperty("ManagerName", name);
	}
	
	/**
	 * Returns or sets a String representing the middle name for the contact.
	 * Read/write.
	 * <p>
	 * This property is parsed from the FullName property, but may be changed or
	 * entered independently should it be parsed incorrectly. Note that any such
	 * changes or entries to this property will be overwritten by any subsequent
	 * changes of entries to FullName.
	 * </p>
	 * 
	 * @return a String representing the middle name for the contact.
	 */
	public String getMiddleName() {
		
		return getStringProperty("MiddleName");
	}
	
	/**
	 * Returns or sets a String representing the middle name for the contact.
	 * Read/write.
	 * <p>
	 * This property is parsed from the FullName property, but may be changed or
	 * entered independently should it be parsed incorrectly. Note that any such
	 * changes or entries to this property will be overwritten by any subsequent
	 * changes of entries to FullName.
	 * </p>
	 * 
	 * @param name
	 *            a String representing the middle name for the contact.
	 */
	public void setMiddleName(String name) {
		
		setProperty("MiddleName", name);
	}
	
	/**
	 * Returns or sets a String representing the mobile telephone number for the
	 * contact. Read/write.
	 * 
	 * @return a String representing the mobile telephone number for the
	 *         contact.
	 */
	public String getMobileTelephoneNumber() {
		
		return getStringProperty("MobileTelephoneNumber");
	}
	
	/**
	 * Returns or sets a String representing the mobile telephone number for the
	 * contact. Read/write.
	 * 
	 * @param num
	 *            a String representing the mobile telephone number for the
	 *            contact.
	 */
	public void setMobileTelephoneNumber(String num) {
		
		setProperty("MobileTelephoneNumber", num);
	}
	
	/**
	 * Returns or sets a String indicating the user's Microsoft NetMeeting ID,
	 * or alias. Read/write.
	 * 
	 * @return Returns or sets a String indicating the user's Microsoft
	 *         NetMeeting ID, or alias. Read/write.
	 */
	public String getNetMeetingAlias() {
		
		return getStringProperty("NetMeetingAlias");
	}
	
	/**
	 * Returns or sets Returns or sets a String indicating the user's Microsoft
	 * NetMeeting ID, or alias. Read/write. Read/write.
	 * 
	 * @param alias
	 *            Returns or sets a String indicating the user's Microsoft
	 *            NetMeeting ID, or alias. Read/write.
	 */
	public void setNetMeetingAlias(String alias) {
		
		setProperty("NetMeetingAlias", alias);
	}
	
	/**
	 * Returns or sets a String specifying the name of the Microsoft NetMeeting
	 * server being used for an online meeting. Read/write.
	 * 
	 * @return a String specifying the name of the Microsoft NetMeeting server
	 *         being used for an online meeting.
	 */
	public String getNetMeetingServer() {
		
		return getStringProperty("NetMeetingServer");
	}
	
	/**
	 * Returns or sets a String specifying the name of the Microsoft NetMeeting
	 * server being used for an online meeting. Read/write.
	 * 
	 * @param server
	 *            a String specifying the name of the Microsoft NetMeeting
	 *            server being used for an online meeting.
	 */
	public void setNetMeetingServer(String server) {
		
		setProperty("NetMeetingServer", server);
	}
	
	/**
	 * Returns or sets a String representing the nickname for the contact.
	 * Read/write.
	 * 
	 * @return a String representing the nickname for the contact.
	 */
	public String getNickName() {
		
		return getStringProperty("NickName");
	}
	
	/**
	 * Returns or sets a String representing the nickname for the contact.
	 * Read/write.
	 * 
	 * @param name
	 *            a String representing the nickname for the contact.
	 */
	public void setNickName(String name) {
		
		setProperty("NickName", name);
	}
	
	/**
	 * Returns or sets a String specifying the specific office location (for
	 * example, Building 1 Room 1 or Suite 123) for the contact. Read/write.
	 * 
	 * @return a String specifying the specific office location (for example,
	 *         Building 1 Room 1 or Suite 123) for the contact.
	 */
	public String getOfficeLocation() {
		
		return getStringProperty("OfficeLocation");
	}
	
	/**
	 * Returns or sets a String specifying the specific office location (for
	 * example, Building 1 Room 1 or Suite 123) for the contact. Read/write.
	 * 
	 * @param location
	 *            a String specifying the specific office location (for example,
	 *            Building 1 Room 1 or Suite 123) for the contact.
	 */
	public void setOfficeLocation(String location) {
		
		setProperty("OfficeLocation", location);
	}
	
	/**
	 * Returns or sets a String representing the organisational ID number for
	 * the contact. Read/write.
	 * 
	 * @return a String representing the organisational ID number for the
	 *         contact.
	 */
	public String getOrganizationalIDNumber() {
		
		return getStringProperty("OrganizationalIDNumber");
	}
	
	/**
	 * Returns or sets a String representing the organisational ID number for
	 * the contact. Read/write.
	 * 
	 * @param id
	 *            a String representing the organisational ID number for the
	 *            contact.
	 */
	public void setOrganizationalIDNumber(String id) {
		
		setProperty("OrganizationalIDNumber", id);
	}
	
	/**
	 * Returns or sets a String representing the other address for the contact.
	 * Read/write.
	 * <p>
	 * This property contains the full, unparsed other address for the contact.
	 * </p>
	 * 
	 * @return a String representing the other address for the contact.
	 */
	public String getOtherAddress() {
		
		return getStringProperty("OtherAddress");
	}
	
	/**
	 * Returns or sets a String representing the other address for the contact.
	 * Read/write.
	 * <p>
	 * This property contains the full, unparsed other address for the contact.
	 * </p>
	 * 
	 * @param addr
	 *            a String representing the other address for the contact.
	 */
	public void setOtherAddress(String addr) {
		
		setProperty("OtherAddress", addr);
	}
	
	/**
	 * Returns or sets a String representing the city portion of the other
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the OtherAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to OtherAddress.
	 * </p>
	 * 
	 * @return a String representing the city portion of the other address for
	 *         the contact.
	 */
	public String getOtherAddressCity() {
		
		return getStringProperty("OtherAddressCity");
	}
	
	/**
	 * Returns or sets a String representing the city portion of the other
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the OtherAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to OtherAddress.
	 * </p>
	 * 
	 * @param city
	 *            a String representing the city portion of the other address
	 *            for the contact.
	 */
	public void setOtherAddressCity(String city) {
		
		setProperty("OtherAddressCity", city);
	}
	
	/**
	 * Returns or sets a String representing the country/region portion of the
	 * other address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the OtherAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to OtherAddress.
	 * </p>
	 * 
	 * @return a String representing the country/region portion of the other
	 *         address for the contact.
	 */
	public String getOtherAddressCountry() {
		
		return getStringProperty("OtherAddressCountry");
	}
	
	/**
	 * Returns or sets a String representing the country/region portion of the
	 * other address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the OtherAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to OtherAddress.
	 * </p>
	 * 
	 * @param country
	 *            a String representing the country/region portion of the other
	 *            address for the contact.
	 */
	public void setOtherAddressCountry(String country) {
		
		setProperty("OtherAddressCountry", country);
	}
	
	/**
	 * Returns or sets a String representing the postal code portion of the
	 * other address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the OtherAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to OtherAddress.
	 * </p>
	 * 
	 * @return a String representing the postal code portion of the other
	 *         address for the contact.
	 */
	public String getOtherAddressPostalCode() {
		
		return getStringProperty("OtherAddressPostalCode");
	}
	
	/**
	 * Returns or sets a String representing the postal code portion of the
	 * other address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the OtherAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to OtherAddress.
	 * </p>
	 * 
	 * @param code
	 *            a String representing the postal code portion of the other
	 *            address for the contact.
	 */
	public void setOtherAddressPostalCode(String code) {
		
		setProperty("OtherAddressPostalCode", code);
	}
	
	/**
	 * Returns or sets a String representing the post office box portion of the
	 * other address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the OtherAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to OtherAddress.
	 * </p>
	 * 
	 * @return a String representing the post office box portion of the other
	 *         address for the contact.
	 */
	public String getOtherAddressPostOfficeBox() {
		
		return getStringProperty("OtherAddressPostOfficeBox");
	}
	
	/**
	 * Returns or sets a String representing the post office box portion of the
	 * other address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the OtherAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to OtherAddress.
	 * </p>
	 * 
	 * @param box
	 *            a String representing the post office box portion of the other
	 *            address for the contact.
	 */
	public void setOtherAddressPostOfficeBox(String box) {
		
		setProperty("OtherAddressPostOfficeBox", box);
	}
	
	/**
	 * Returns or sets a String representing the state portion of the other
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the OtherAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to OtherAddress.
	 * </p>
	 * 
	 * @return a String representing the state portion of the other address for
	 *         the contact.
	 */
	public String getOtherAddressState() {
		
		return getStringProperty("OtherAddressState");
	}
	
	/**
	 * Returns or sets a String representing the state portion of the other
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the OtherAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to OtherAddress.
	 * </p>
	 * 
	 * @param state
	 *            a String representing the state portion of the other address
	 *            for the contact.
	 */
	public void setOtherAddressState(String state) {
		
		setProperty("OtherAddressState", state);
	}
	
	/**
	 * Returns or sets a String representing the street portion of the other
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the OtherAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to OtherAddress.
	 * </p>
	 * 
	 * @return a String representing the street portion of the other address for
	 *         the contact.
	 */
	public String getOtherAddressStreet() {
		
		return getStringProperty("OtherAddressStreet");
	}
	
	/**
	 * Returns or sets a String representing the street portion of the other
	 * address for the contact. Read/write.
	 * <p>
	 * This property is parsed from the OtherAddress property, but may be
	 * changed or entered independently should it be parsed incorrectly. Note
	 * that any such changes or entries to this property will be overwritten by
	 * any subsequent changes or entries to OtherAddress.
	 * </p>
	 * 
	 * @param street
	 *            a String representing the street portion of the other address
	 *            for the contact.
	 */
	public void setOtherAddressStreet(String street) {
		
		setProperty("OtherAddressStreet", street);
	}
	
	/**
	 * Returns or sets a String representing the other fax number for the
	 * contact. Read/write.
	 * 
	 * @return a String representing the other fax number for the contact.
	 */
	public String getOtherFaxNumber() {
		
		return getStringProperty("OtherFaxNumber");
	}
	
	/**
	 * Returns or sets a String representing the other fax number for the
	 * contact. Read/write.
	 * 
	 * @param num
	 *            a String representing the other fax number for the contact.
	 */
	public void setOtherFaxNumber(String num) {
		
		setProperty("OtherFaxNumber", num);
	}
	
	/**
	 * Returns or sets a String representing the other telephone number for the
	 * contact. Read/write.
	 * 
	 * @return a String representing the other telephone number for the contact.
	 */
	public String getOtherTelephoneNumber() {
		
		return getStringProperty("OtherTelephoneNumber");
	}
	
	/**
	 * Returns or sets a String representing the other telephone number for the
	 * contact. Read/write.
	 * 
	 * @param num
	 *            a String representing the other telephone number for the
	 *            contact.
	 */
	public void setOtherTelephoneNumber(String num) {
		
		setProperty("OtherTelephoneNumber", num);
	}
	
	/**
	 * Returns or sets a String representing the pager number for the contact.
	 * Read/write.
	 * 
	 * @return a String representing the pager number for the contact.
	 */
	public String getPagerNumber() {
		
		return getStringProperty("PagerNumber");
	}
	
	/**
	 * Returns or sets a String representing the pager number for the contact.
	 * Read/write.
	 * 
	 * @param num
	 *            a String representing the pager number for the contact.
	 */
	public void setPagerNumber(String num) {
		
		setProperty("PagerNumber", num);
	}
	
	/**
	 * Returns or sets a String representing the URL of the personal Web page
	 * for the contact. Read/write.
	 * 
	 * @return a String representing the URL of the personal Web page for the
	 *         contact.
	 */
	public String getPersonalHomePage() {
		
		return getStringProperty("PersonalHomePage");
	}
	
	/**
	 * Returns or sets a String representing the URL of the personal Web page
	 * for the contact. Read/write.
	 * 
	 * @param url
	 *            a String representing the URL of the personal Web page for the
	 *            contact.
	 */
	public void setPersonalHomePage(String url) {
		
		setProperty("PersonalHomePage", url);
	}
	
	/**
	 * Returns or sets a String specifying the primary telephone number for the
	 * contact. Read/write.
	 * 
	 * @return a String specifying the primary telephone number for the contact.
	 */
	public String getPrimaryTelephoneNumber() {
		
		return getStringProperty("PrimaryTelephoneNumber");
	}
	
	/**
	 * Returns or sets a String specifying the primary telephone number for the
	 * contact. Read/write.
	 * 
	 * @param num
	 *            a String specifying the primary telephone number for the
	 *            contact.
	 */
	public void setPrimaryTelephoneNumber(String num) {
		
		setProperty("PrimaryTelephoneNumber", num);
	}
	
	/**
	 * Returns or sets a String indicating the profession for the contact.
	 * Read/write.
	 * 
	 * @return a String indicating the profession for the contact.
	 */
	public String getProfession() {
		
		return getStringProperty("Profession");
	}
	
	/**
	 * Returns or sets a String indicating the profession for the contact.
	 * Read/write.
	 * 
	 * @param prof
	 *            a String indicating the profession for the contact.
	 */
	public void setProfession(String prof) {
		
		setProperty("Profession", prof);
	}
	
	/**
	 * Returns or sets a String indicating the radio telephone number for the
	 * contact. Read/write.
	 * 
	 * @return a String indicating the radio telephone number for the contact.
	 */
	public String getRadioTelephoneNumber() {
		
		return getStringProperty("RadioTelephoneNumber");
	}
	
	/**
	 * Returns or sets a String indicating the radio telephone number for the
	 * contact. Read/write.
	 * 
	 * @param num
	 *            a String indicating the radio telephone number for the
	 *            contact.
	 */
	public void setRadioTelephoneNumber(String num) {
		
		setProperty("RadioTelephoneNumber", num);
	}
	
	/**
	 * Returns or sets a String specifying the referral name entry for the
	 * contact. Read/write.
	 * 
	 * @return a String specifying the referral name entry for the contact.
	 */
	public String getReferredBy() {
		
		return getStringProperty("ReferredBy");
	}
	
	/**
	 * Returns or sets a String specifying the referral name entry for the
	 * contact. Read/write.
	 * 
	 * @param name
	 *            a String specifying the referral name entry for the contact.
	 */
	public void setReferredBy(String name) {
		
		setProperty("ReferredBy", name);
	}
	
	/**
	 * Removes a picture from a Contacts item.
	 */
	public void removePicture() {
		
		invokeNoReply("RemovePicture");
	}
	
	/**
	 * Resets the Electronic Business Card on the contact item to the default
	 * business card, deleting any custom layout and logo on the Electronic
	 * Business Card.
	 * <p>
	 * For contacts with a Microsoft Office InterConnect card type, this will
	 * reset the contact to using an Outlook card type.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 */
	public void resetBusinessCard() {
		
		invokeNoReply("ResetBusinessCard");
	}
	
	/**
	 * Saves an image of the business card generated from the specified
	 * ContactItem object.
	 * <p>
	 * This method generates an image, as a Portable Network Graphics (.png)
	 * file, of the business card generated from the specified ContactItem
	 * object. If the path and file name specified in Path cannot be resolved,
	 * an error occurs.
	 * </p>
	 * 
	 * @param path
	 *            The fully qualified path and file name of the image to be
	 *            saved.
	 */
	public void saveBusinessCardImage(String path) {
		
		invokeNoReply("SaveBusinessCardImage", newVariant(path));
	}
	
	/**
	 * Returns or sets a MailingAddress constant indicating the type of the
	 * mailing address for the contact. Read/write.
	 * 
	 * @return a MailingAddress constant indicating the type of the mailing
	 *         address for the contact.
	 */
	public MailingAddressType getSelectedMailingAddress() {
		
		return MailingAddressType.parse(getShortProperty("SelectedMailingAddress"));
	}
	
	/**
	 * Returns or sets a MailingAddress constant indicating the type of the
	 * mailing address for the contact. Read/write.
	 * 
	 * @param addrType
	 *            a MailingAddress constant indicating the type of the mailing
	 *            address for the contact.
	 */
	public void setSelectedMailingAddress(MailingAddressType addrType) {
		
		setProperty("SelectedMailingAddress", addrType.value());
	}
	
	/**
	 * Displays the electronic business card (EBC) editor dialog box for the
	 * ContactItem object.
	 * <p>
	 * Calling this method retrieves the data for the specified ContactItem
	 * object and then modally displays that data in the EBC editor dialog box.
	 * An error occurs if the data cannot be retrieved.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 */
	public void showBusinessCardEditor() {
		
		invokeNoReply("ShowBusinessCardEditor");
	}
	
	/**
	 * Displays the Check Phone Number dialog box for a specified telephone
	 * number contained by a ContactItem object.
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param num
	 *            The type of telephone number to be checked.
	 */
	public void showCheckPhoneDialog(ContactPhoneNumber num) {
		
		invokeNoReply("ShowCheckPhoneDialog", newVariant(num.value()));
	}
	
	/**
	 * Returns or sets a String indicating the spouse/partner name entry for the
	 * contact. Read/write.
	 * 
	 * @return a String indicating the spouse/partner name entry for the
	 *         contact.
	 */
	public String getSpouse() {
		
		return getStringProperty("Spouse");
	}
	
	/**
	 * Returns or sets a String indicating the spouse/partner name entry for the
	 * contact. Read/write.
	 * 
	 * @param name
	 *            a String indicating the spouse/partner name entry for the
	 *            contact.
	 */
	public void setSpouse(String name) {
		
		setProperty("Spouse", name);
	}
	
	/**
	 * Returns or sets a String indicating the name suffix (such as Jr., III, or
	 * Ph.D.) for the specified contact. Read/write.
	 * <p>
	 * The LastName , FirstName , MiddleName , and Suffix properties are parsed
	 * from the FullName property.
	 * </p>
	 * 
	 * @return a String indicating the name suffix (such as Jr., III, or Ph.D.)
	 *         for the specified contact.
	 */
	public String getSuffix() {
		
		return getStringProperty("Suffix");
	}
	
	/**
	 * Returns or sets a String indicating the name suffix (such as Jr., III, or
	 * Ph.D.) for the specified contact. Read/write.
	 * <p>
	 * The LastName , FirstName , MiddleName , and Suffix properties are parsed
	 * from the FullName property.
	 * </p>
	 * 
	 * @param val
	 *            a String indicating the name suffix (such as Jr., III, or
	 *            Ph.D.) for the specified contact.
	 */
	public void setSuffix(String val) {
		
		setProperty("Suffix", val);
	}
	
	/**
	 * Returns or sets a String indicating the telex number for the contact.
	 * Read/write.
	 * 
	 * @return a String indicating the telex number for the contact.
	 */
	public String getTelexNumber() {
		
		return getStringProperty("TelexNumber");
	}
	
	/**
	 * Returns or sets a String indicating the telex number for the contact.
	 * Read/write.
	 * 
	 * @param num
	 *            a String indicating the telex number for the contact.
	 */
	public void setTelexNumber(String num) {
		
		setProperty("TelexNumber", num);
	}
	
	/**
	 * Returns or sets a String indicating the title for the contact.
	 * Read/write.
	 * 
	 * @return a String indicating the title for the contact.
	 */
	public String getTitle() {
		
		return getStringProperty("Title");
	}
	
	/**
	 * Returns or sets a String indicating the title for the contact.
	 * Read/write.
	 * 
	 * @param ttl
	 *            a String indicating the title for the contact.
	 */
	public void setTitle(String ttl) {
		
		setProperty("Title", ttl);
	}
	
	/**
	 * Returns or sets a String specifying the first field on the Contacts form
	 * intended for miscellaneous use for the contact. Read/write.
	 * <p>
	 * The properties ContactItem.User1, ContactItem.User2, ContactItem.User3,
	 * and ContactItem.User4 map to the fields User Field 1, User Field 2, User
	 * Field 3, and User Field 4 on the Contacts form respectively. These
	 * properties are explicit built-in String properties; users can use these
	 * fields for miscellaneous purposes for the contact.
	 * </p>
	 * 
	 * @return a String specifying the first field on the Contacts form intended
	 *         for miscellaneous use for the contact.
	 */
	public String getUser1() {
		
		return getStringProperty("User1");
	}
	
	/**
	 * Returns or sets a String specifying the first field on the Contacts form
	 * intended for miscellaneous use for the contact. Read/write.
	 * <p>
	 * The properties ContactItem.User1, ContactItem.User2, ContactItem.User3,
	 * and ContactItem.User4 map to the fields User Field 1, User Field 2, User
	 * Field 3, and User Field 4 on the Contacts form respectively. These
	 * properties are explicit built-in String properties; users can use these
	 * fields for miscellaneous purposes for the contact.
	 * </p>
	 * 
	 * @param usr
	 *            a String specifying the first field on the Contacts form
	 *            intended for miscellaneous use for the contact.
	 */
	public void setUser1(String usr) {
		
		setProperty("User1", usr);
	}
	
	/**
	 * Returns or sets a String specifying the first field on the Contacts form
	 * intended for miscellaneous use for the contact. Read/write.
	 * <p>
	 * The properties ContactItem.User1, ContactItem.User2, ContactItem.User3,
	 * and ContactItem.User4 map to the fields User Field 1, User Field 2, User
	 * Field 3, and User Field 4 on the Contacts form respectively. These
	 * properties are explicit built-in String properties; users can use these
	 * fields for miscellaneous purposes for the contact.
	 * </p>
	 * 
	 * @return a String specifying the first field on the Contacts form intended
	 *         for miscellaneous use for the contact.
	 */
	public String getUser2() {
		
		return getStringProperty("User2");
	}
	
	/**
	 * Returns or sets a String specifying the first field on the Contacts form
	 * intended for miscellaneous use for the contact. Read/write.
	 * <p>
	 * The properties ContactItem.User1, ContactItem.User2, ContactItem.User3,
	 * and ContactItem.User4 map to the fields User Field 1, User Field 2, User
	 * Field 3, and User Field 4 on the Contacts form respectively. These
	 * properties are explicit built-in String properties; users can use these
	 * fields for miscellaneous purposes for the contact.
	 * </p>
	 * 
	 * @param usr
	 *            a String specifying the first field on the Contacts form
	 *            intended for miscellaneous use for the contact.
	 */
	public void setUser2(String usr) {
		
		setProperty("User2", usr);
	}
	
	/**
	 * Returns or sets a String specifying the first field on the Contacts form
	 * intended for miscellaneous use for the contact. Read/write.
	 * <p>
	 * The properties ContactItem.User1, ContactItem.User2, ContactItem.User3,
	 * and ContactItem.User4 map to the fields User Field 1, User Field 2, User
	 * Field 3, and User Field 4 on the Contacts form respectively. These
	 * properties are explicit built-in String properties; users can use these
	 * fields for miscellaneous purposes for the contact.
	 * </p>
	 * 
	 * @return a String specifying the first field on the Contacts form intended
	 *         for miscellaneous use for the contact.
	 */
	public String getUser3() {
		
		return getStringProperty("User3");
	}
	
	/**
	 * Returns or sets a String specifying the first field on the Contacts form
	 * intended for miscellaneous use for the contact. Read/write.
	 * <p>
	 * The properties ContactItem.User1, ContactItem.User2, ContactItem.User3,
	 * and ContactItem.User4 map to the fields User Field 1, User Field 2, User
	 * Field 3, and User Field 4 on the Contacts form respectively. These
	 * properties are explicit built-in String properties; users can use these
	 * fields for miscellaneous purposes for the contact.
	 * </p>
	 * 
	 * @param usr
	 *            a String specifying the first field on the Contacts form
	 *            intended for miscellaneous use for the contact.
	 */
	public void setUser3(String usr) {
		
		setProperty("User3", usr);
	}
	
	/**
	 * Returns or sets a String specifying the first field on the Contacts form
	 * intended for miscellaneous use for the contact. Read/write.
	 * <p>
	 * The properties ContactItem.User1, ContactItem.User2, ContactItem.User3,
	 * and ContactItem.User4 map to the fields User Field 1, User Field 2, User
	 * Field 3, and User Field 4 on the Contacts form respectively. These
	 * properties are explicit built-in String properties; users can use these
	 * fields for miscellaneous purposes for the contact.
	 * </p>
	 * 
	 * @return a String specifying the first field on the Contacts form intended
	 *         for miscellaneous use for the contact.
	 */
	public String getUser4() {
		
		return getStringProperty("User4");
	}
	
	/**
	 * Returns or sets a String specifying the first field on the Contacts form
	 * intended for miscellaneous use for the contact. Read/write.
	 * <p>
	 * The properties ContactItem.User1, ContactItem.User2, ContactItem.User3,
	 * and ContactItem.User4 map to the fields User Field 1, User Field 2, User
	 * Field 3, and User Field 4 on the Contacts form respectively. These
	 * properties are explicit built-in String properties; users can use these
	 * fields for miscellaneous purposes for the contact.
	 * </p>
	 * 
	 * @param usr
	 *            a String specifying the first field on the Contacts form
	 *            intended for miscellaneous use for the contact.
	 */
	public void setUser4(String usr) {
		
		setProperty("User4", usr);
	}
	
	/**
	 * Returns or sets a String indicating the URL (Uniform Resource Locator
	 * (URL): An address that specifies a protocol (such as HTTP or FTP) and a
	 * location of an object, document, World Wide Web page, or other
	 * destination on the Internet or an intranet, for example:
	 * http://www.microsoft.com/.) of the Web page for the contact. Read/write.
	 * 
	 * @return a String indicating the URL.
	 */
	public String getWebPage() {
		
		return getStringProperty("WebPage");
	}
	
	/**
	 * Returns or sets a String indicating the URL (Uniform Resource Locator
	 * (URL): An address that specifies a protocol (such as HTTP or FTP) and a
	 * location of an object, document, World Wide Web page, or other
	 * destination on the Internet or an intranet, for example:
	 * http://www.microsoft.com/.) of the Web page for the contact. Read/write.
	 * 
	 * @param url
	 *            a String indicating the URL.
	 */
	public void setWebPage(String url) {
		
		setProperty("WebPage", url);
	}
	
	/**
	 * Returns or sets a String indicating the Japanese phonetic rendering
	 * (yomigana) of the company name for the contact. Read/write.
	 * 
	 * @return a String indicating the Japanese phonetic rendering (yomigana) of
	 *         the company name for the contact.
	 */
	public String getYomiCompanyName() {
		
		return getStringProperty("YomiCompanyName");
	}
	
	/**
	 * Returns or sets a String indicating the Japanese phonetic rendering
	 * (yomigana) of the company name for the contact. Read/write.
	 * 
	 * @param val
	 *            a String indicating the Japanese phonetic rendering (yomigana)
	 *            of the company name for the contact.
	 */
	public void setYomiCompanyName(String val) {
		
		setProperty("YomiCompanyName", val);
	}
	
	/**
	 * Returns or sets a String indicating the Japanese phonetic rendering
	 * (yomigana) of the first name for the contact. Read/write.
	 * 
	 * @return a String indicating the Japanese phonetic rendering (yomigana) of
	 *         the first name for the contact.
	 */
	public String getYomiFirstName() {
		
		return getStringProperty("YomiFirstName");
	}
	
	/**
	 * Returns or sets a String indicating the Japanese phonetic rendering
	 * (yomigana) of the first name for the contact. Read/write.
	 * 
	 * @param val
	 *            a String indicating the Japanese phonetic rendering (yomigana)
	 *            of the first name for the contact.
	 */
	public void setYomiFirstName(String val) {
		
		setProperty("YomiFirstName", val);
	}
	
	/**
	 * Returns or sets a String indicating the Japanese phonetic rendering
	 * (yomigana) of the last name for the contact. Read/write.
	 * 
	 * @return a String indicating the Japanese phonetic rendering (yomigana) of
	 *         the last name for the contact.
	 */
	public String getYomiLastName() {
		
		return getStringProperty("YomiLastName");
	}
	
	/**
	 * Returns or sets a String indicating the Japanese phonetic rendering
	 * (yomigana) of the last name for the contact. Read/write.
	 * 
	 * @param val
	 *            a String indicating the Japanese phonetic rendering (yomigana)
	 *            of the last name for the contact.
	 */
	public void setYomiLastName(String val) {
		
		setProperty("YomiLastName", val);
	}
}
