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
 * There are an inordinate amount of
 * methods and properties that are shared between multiple Item types. However,
 * the commonality model is complex and there is a fine line between reducing
 * duplication and having a ridiculous depth to the inheritance model. The
 * compromise struck here was to have four levels of base class. Item classes
 * inherit from the most appropriate level.
 * 
 * @author Ian Darby
 * 
 * @see BaseItemLevel3
 * @see MailItem
 * @see SharingItem
 */
public class BaseItemLevel4 extends BaseItemLevel3 {

	/**
	 * Constructor scope is restricted to inheritance and package as it should
	 * not be used directly by user applications. It is only intended to be used
	 * from within factory methods and properties of the Outlook object model
	 * itself. It may also be called from unit tests which may supply a mock
	 * version of the IDispatch object.
	 * 
	 * @param iDisp
	 *            the IDispatch object which is the underlying Actions object
	 *            within the Outlook object model. All methods and properties of
	 *            this wrapper class ultimately delegate to IDispatch.
	 */
	protected BaseItemLevel4(IDispatch iDisp) {
		super(iDisp);
	}
	
	/**
	 * Appends contact information based on the Electronic Business Card (EBC)
	 * associated with the specified ContactItem object to the MailItem object.
	 * <p>
	 * This method adds contact information, generated from the information
	 * stored in the ContactItem object, to the existing MailItem object. The
	 * information included depends on the value of the BodyFormat property for
	 * the MailItem object:
	 * <ul>
	 * <li>olFormatPlain - A vCard (.vcf) file is created and added to the
	 * Attachments collection of the MailItem object.</li>
	 * 
	 * <li>olFormatRichText - A vCard (.vcf) file is created and added to the
	 * Attachments collection of the MailItem object.</li>
	 * 
	 * <li>olFormatHTML - An image of the business card is generated and included
	 * in the Body property of the MailItem object, and a vCard (.vcf) file is
	 * created and added to the Attachments collection of the MailItem object.</li>
	 * </ul>
	 * </p>
	 * <p>
	 * The attached vCard file contains only the contact information included in
	 * the Electronic Business Card associated with the ContactItem object. Any
	 * contact information not displayed in the Electronic Business Card is
	 * excluded from the vCard file.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param contact
	 *            The contact item from which to obtain the business card
	 *            information.
	 */
	public void addBusinessCard(ContactItem contact) {
		
		invoke("AddBusinessCard", newVariant(contact.getIDispatch()));
	}

	/**
	 * Returns true if the mail message can be forwarded. Read/write.
	 * 
	 * @return true if the mail message can be forwarded.
	 */
	public boolean isAlternateRecipientAllowed() {
		
		return getBooleanProperty("AlternateRecipientAllowed");
	}
	
	/**
	 * Set to true if the mail message can be forwarded. Read/write.
	 * 
	 * @param canForwardItem
	 *            true if the mail message can be forwarded.
	 */
	public void setAlternateRecipientAllowed(boolean canForwardItem) {
		
		setProperty("VotingResponse", canForwardItem);
	}
	
	/**
	 * A boolean value that returns true if the item was automatically
	 * forwarded. Read/write.
	 * 
	 * @return boolean value that returns true if the item was automatically
	 *         forwarded.
	 */
	public boolean isAutoForwarded() {
		
		return getBooleanProperty("AutoForwarded");
	}
	
	/**
	 * A boolean value that returns true if the item was automatically
	 * forwarded. Read/write.
	 * 
	 * @param flag
	 *            boolean value that returns true if the item was automatically
	 *            forwarded.
	 */
	public void setAutoForwarded(boolean flag) {
		
		setProperty("AutoForwarded", flag);
	}
	
	/**
	 * Returns a String representing the display list of blind carbon copy (BCC)
	 * names for an Item. Read/write.
	 * <p>
	 * This property contains the display names only. The Recipients collection
	 * should be used to modify the BCC recipients.
	 * </p>
	 * 
	 * @return a String representing the display list of blind carbon copy (BCC)
	 *         names for an Item.
	 */
	public String getBcc() {
		
		return getStringProperty("BCC");
	}
	
	/**
	 * Returns a String representing the display list of blind carbon copy (BCC)
	 * names for an Item. Read/write.
	 * <p>
	 * This property contains the display names only. The Recipients collection
	 * should be used to modify the BCC recipients.
	 * </p>
	 * 
	 * @param bcc
	 *            a String representing the display list of blind carbon copy
	 *            (BCC) names for an Item.
	 */
	public void setBcc(String bcc) {
		
		setProperty("BCC", bcc);
	}
	
	/**
	 * Returns or sets a BodyFormat constant indicating the format of the body
	 * text. Read/write.
	 * <p>
	 * The body text format determines the standard used to display the text of
	 * the message. Microsoft Outlook provides three body text format options:
	 * Plain Text, Rich Text (RTF), and HTML.
	 * </p>
	 * <p>
	 * All text formatting will be lost when the BodyFormat property is switched
	 * from RTF to HTML and vice-versa.
	 * </p>
	 * 
	 * @return BodyFormat constant indicating the format of the body text.
	 */
	public MailBodyFormat getBodyFormat() {
		
		return MailBodyFormat.parse(getIntProperty("BodyFormat"));
	}
	
	/**
	 * Returns or sets a BodyFormat constant indicating the format of the body
	 * text. Read/write.
	 * <p>
	 * The body text format determines the standard used to display the text of
	 * the message. Microsoft Outlook provides three body text format options:
	 * Plain Text, Rich Text (RTF), and HTML.
	 * </p>
	 * <p>
	 * All text formatting will be lost when the BodyFormat property is switched
	 * from RTF to HTML and vice-versa.
	 * </p>
	 * 
	 * @param fmt
	 *            BodyFormat constant indicating the format of the body text.
	 */
	public void setBodyFormat(MailBodyFormat fmt) {
		
		setProperty("BodyFormat", fmt.value());
	}
	
	/**
	 * Returns a String representing the display list of carbon copy (CC) names
	 * for an Item. Read/write.
	 * <p>
	 * This property contains the display names only. The Recipients collection
	 * should be used to modify the CC recipients.
	 * </p>
	 * 
	 * @return a String representing the display list of carbon copy (CC) names
	 *         for an Item.
	 */
	public String getCc() {
		
		return getStringProperty("CC");
	}
	
	/**
	 * Returns a String representing the display list of carbon copy (CC) names
	 * for an Item. Read/write.
	 * <p>
	 * This property contains the display names only. The Recipients collection
	 * should be used to modify the CC recipients.
	 * </p>
	 * 
	 * @param cc
	 *            a String representing the display list of carbon copy (CC)
	 *            names for an Item.
	 */
	public void setCc(String cc) {
		
		setProperty("CC", cc);
	}
	
	/**
	 * Clears the index of the conversation thread for the mail message.
	 */
	public void clearConversationIndex() {
		
		invokeNoReply("ClearConversationIndex");
	}
	
	/**
	 * Returns or sets a Date indicating the date and time the mail message is
	 * to be delivered. Read/write.
	 * <p>
	 * This property corresponds to the MAPI property
	 * PidTagDeferredDeliveryTime.
	 * </p>
	 * 
	 * @return a Date indicating the date and time the mail message is to be
	 *         delivered.
	 */
	public Date getDeferredDeliveryTime() {
		
		return getDateProperty("DeferredDeliveryTime");
	}
	
	/**
	 * Returns or sets a Date indicating the date and time the mail message is
	 * to be delivered. Read/write.
	 * <p>
	 * This property corresponds to the MAPI property
	 * PidTagDeferredDeliveryTime.
	 * </p>
	 * 
	 * @param deliverAt
	 *            a Date indicating the date and time the mail message is to be
	 *            delivered.
	 */
	public void setDeferredDeliveryTime(Date deliverAt) {
		
		setProperty("DeferredDeliveryTime", deliverAt);
	}
	
	/**
	 * Returns or sets a boolean value that is true if a copy of the mail
	 * message is not saved upon being sent, and false if a copy is saved.
	 * Read/write.
	 * 
	 * @return a boolean value that is true if a copy of the mail message is not
	 *         saved upon being sent, and false if a copy is saved.
	 */
	public boolean isDeleteAfterSubmit() {
		
		return getBooleanProperty("DeleteAfterSubmit");
	}
	
	/**
	 * Returns or sets a boolean value that is true if a copy of the mail
	 * message is not saved upon being sent, and false if a copy is saved.
	 * Read/write.
	 * 
	 * @param flag
	 *            a boolean value that is true if a copy of the mail message is
	 *            not saved upon being sent, and false if a copy is saved.
	 */
	public void setDeleteAfterSubmit(boolean flag) {
		
		setProperty("DeleteAfterSubmit", flag);
	}
	
	/**
	 * Returns or sets a Date indicating the date and time at which the item
	 * becomes invalid and can be deleted. Read/write.
	 * 
	 * @return a Date indicating the date and time at which the item becomes
	 *         invalid and can be deleted.
	 */
	public Date getExpiryTime() {
		
		return getDateProperty("ExpiryTime");
	}
	
	/**
	 * Returns or sets a Date indicating the date and time at which the item
	 * becomes invalid and can be deleted. Read/write.
	 * 
	 * @param expireAt
	 *            a Date indicating the date and time at which the item becomes
	 *            invalid and can be deleted.
	 */
	public void setExpiryTime(Date expireAt) {
		
		setProperty("ExpiryTime", expireAt);
	}
	
	/**
	 * Returns or sets a String that indicates the requested action for a mail
	 * item. Read/write.
	 * <p>
	 * By default, a mail item is not marked with any flag and the default value
	 * for this property is the empty string. You can set the value of
	 * FlagRequest either through the user interface or programmatically. When
	 * you mark a mail item with a flag through the user interface, the default
	 * value of FlagRequest is "Follow up".
	 * </p>
	 * 
	 * @return a String that indicates the requested action for a mail item.
	 */
	public String getFlagRequest() {
		
		return getStringProperty("FlagRequest");
	}
	
	/**
	 * Returns or sets a String that indicates the requested action for a mail
	 * item. Read/write.
	 * <p>
	 * By default, a mail item is not marked with any flag and the default value
	 * for this property is the empty string. You can set the value of
	 * FlagRequest either through the user interface or programmatically. When
	 * you mark a mail item with a flag through the user interface, the default
	 * value of FlagRequest is "Follow up".
	 * </p>
	 * 
	 * @param flagText
	 *            a String that indicates the requested action for a mail item.
	 */
	public void setFlagRequest(String flagText) {
		
		setProperty("FlagRequest", flagText);
	}
	
	/**
	 * Returns or sets a String representing the HTML body of the specified item
	 * (item: An item is the basic element that holds information in Outlook
	 * (similar to a file in other programs). Items include e-mail messages,
	 * appointments, contacts, tasks, journal entries, notes, posted items, and
	 * documents.). Read/write.
	 * <p>
	 * The HTMLBody property should be an HTML syntax string.
	 * </p>
	 * <p>
	 * Setting the HTMLBody property will always update the Body property
	 * immediately.
	 * </p>
	 * 
	 * @return a String representing the HTML body of the specified item.
	 */
	public String getHtmlBody() {
		
		return getStringProperty("HTMLBody");
	}
	
	/**
	 * Returns or sets a String representing the HTML body of the specified item
	 * (item: An item is the basic element that holds information in Outlook
	 * (similar to a file in other programs). Items include e-mail messages,
	 * appointments, contacts, tasks, journal entries, notes, posted items, and
	 * documents.). Read/write.
	 * <p>
	 * The HTMLBody property should be an HTML syntax string.
	 * </p>
	 * <p>
	 * Setting the HTMLBody property will always update the Body property
	 * immediately.
	 * </p>
	 * 
	 * @param htmlBody
	 *            a String representing the HTML body of the specified item
	 */
	public void setHtmlBody(String htmlBody) {
		
		setProperty("HTMLBody", htmlBody);
	}
	
	/**
	 * Returns or sets an int that determines the Internet code page used by the
	 * item. Read/write.
	 * <p>
	 * The Internet code page defines the text encoding scheme used by the item.
	 * </p>
	 * <p>
	 * The following table lists the values that are supported by the
	 * InternetCodePage property.
	 * </p>
	 * <p>
	 * Name Character Set Code Page Arabic (ISO) iso-8859-6 28596 Arabic
	 * (Windows) windows-1256 1256 Baltic (ISO) iso-8859-4 28594 Baltic
	 * (Windows) windows-1257 1257 Central European (ISO) iso-8859-2 28592
	 * Central European (Windows) windows-1250 1250 Chinese Simplified (GB2312)
	 * gb2312 936 Chinese Simplified (HZ) hz-gb-2312 52936 Chinese Traditional
	 * (Big5) big5 950 Cyrillic (ISO) iso-8859-5 28595 Cyrillic (KOI8-R) koi8-r
	 * 20866 Cyrillic (KOI8-U) koi8-u 21866 Cyrillic (Windows) windows-1251 1251
	 * Greek (ISO) iso-8859-7 28597 Greek (Windows) windows-1253 1253 Hebrew
	 * (ISO-Logical) iso-8859-8-i 38598 Hebrew (Windows) windows-1255 1255
	 * Japanese (EUC) euc-jp 51932 Japanese (JIS) iso-2022-jp 50220 Japanese
	 * (JIS-Allow 1 byte Kana) csISO2022JP 50221 Japanese (Shift-JIS)
	 * iso-2022-jp 932 Korean ks_c_5601-1987 949 Korean (EUC) euc-kr 51949 Latin
	 * 3 (ISO) iso-8859-3 28593 Latin 9 (ISO) iso-8859-15 28605 Thai (Windows)
	 * windows-874 874 Turkish (ISO) iso-8859-9 28599 Turkish (Windows)
	 * windows-1254 1254 Unicode (UTF-7) utf-7 65000 Unicode (UTF-8) utf-8 65001
	 * US-ASCII us-ascii 20127 Vietnamese (Windows) windows-1258 1258 Western
	 * European (ISO) iso-8859-1 28591 Western European (Windows) Windows-1252
	 * 1252
	 * </p>
	 * <p>
	 * The following table lists the code pages Microsoft recommends that you
	 * use for the best compatibility with older e-mail systems.
	 * </p>
	 * <p>
	 * Name Character Set Code Page Arabic (Windows) windows-1256 1256 Baltic
	 * (ISO) iso-8859-4 28594 Central European (ISO) iso-8859-2 28592 Chinese
	 * Simplified (GB2312) gb2312 936 Chinese Traditional (Big5) big5 950
	 * Cyrillic (KOI8-R) koi8-r 20866 Cyrillic (Windows) windows-1251 1251 Greek
	 * (ISO) iso-8859-7 28597 Hebrew (Windows) windows-1255 1255 Japanese (JIS)
	 * iso-2022-jp 50220 Korean ks_c_5601-1987 949 Thai (Windows) windows-874
	 * 874 Turkish (ISO) iso-8859-9 28599 Unicode (UTF-8) utf-8 65001 US-ASCII
	 * us-ascii 20127 Vietnamese (Windows) windows-1258 1258 Western European
	 * (ISO) iso-8859-1 28591
	 * </p>
	 * 
	 * @return an int that determines the Internet code page used by the item.
	 */
	public int getInternetCodePage() {
		
		return getIntProperty("InternetCodePage");
	}
	
	/**
	 * Returns or sets an int that determines the Internet code page used by the
	 * item. Read/write.
	 * <p>
	 * The Internet code page defines the text encoding scheme used by the item.
	 * </p>
	 * <p>
	 * The following table lists the values that are supported by the
	 * InternetCodePage property.
	 * </p>
	 * <p>
	 * Name Character Set Code Page Arabic (ISO) iso-8859-6 28596 Arabic
	 * (Windows) windows-1256 1256 Baltic (ISO) iso-8859-4 28594 Baltic
	 * (Windows) windows-1257 1257 Central European (ISO) iso-8859-2 28592
	 * Central European (Windows) windows-1250 1250 Chinese Simplified (GB2312)
	 * gb2312 936 Chinese Simplified (HZ) hz-gb-2312 52936 Chinese Traditional
	 * (Big5) big5 950 Cyrillic (ISO) iso-8859-5 28595 Cyrillic (KOI8-R) koi8-r
	 * 20866 Cyrillic (KOI8-U) koi8-u 21866 Cyrillic (Windows) windows-1251 1251
	 * Greek (ISO) iso-8859-7 28597 Greek (Windows) windows-1253 1253 Hebrew
	 * (ISO-Logical) iso-8859-8-i 38598 Hebrew (Windows) windows-1255 1255
	 * Japanese (EUC) euc-jp 51932 Japanese (JIS) iso-2022-jp 50220 Japanese
	 * (JIS-Allow 1 byte Kana) csISO2022JP 50221 Japanese (Shift-JIS)
	 * iso-2022-jp 932 Korean ks_c_5601-1987 949 Korean (EUC) euc-kr 51949 Latin
	 * 3 (ISO) iso-8859-3 28593 Latin 9 (ISO) iso-8859-15 28605 Thai (Windows)
	 * windows-874 874 Turkish (ISO) iso-8859-9 28599 Turkish (Windows)
	 * windows-1254 1254 Unicode (UTF-7) utf-7 65000 Unicode (UTF-8) utf-8 65001
	 * US-ASCII us-ascii 20127 Vietnamese (Windows) windows-1258 1258 Western
	 * European (ISO) iso-8859-1 28591 Western European (Windows) Windows-1252
	 * 1252
	 * </p>
	 * <p>
	 * The following table lists the code pages Microsoft recommends that you
	 * use for the best compatibility with older e-mail systems.
	 * </p>
	 * <p>
	 * Name Character Set Code Page Arabic (Windows) windows-1256 1256 Baltic
	 * (ISO) iso-8859-4 28594 Central European (ISO) iso-8859-2 28592 Chinese
	 * Simplified (GB2312) gb2312 936 Chinese Traditional (Big5) big5 950
	 * Cyrillic (KOI8-R) koi8-r 20866 Cyrillic (Windows) windows-1251 1251 Greek
	 * (ISO) iso-8859-7 28597 Hebrew (Windows) windows-1255 1255 Japanese (JIS)
	 * iso-2022-jp 50220 Korean ks_c_5601-1987 949 Thai (Windows) windows-874
	 * 874 Turkish (ISO) iso-8859-9 28599 Unicode (UTF-8) utf-8 65001 US-ASCII
	 * us-ascii 20127 Vietnamese (Windows) windows-1258 1258 Western European
	 * (ISO) iso-8859-1 28591
	 * </p>
	 * 
	 * @param codePage
	 *            an int that determines the Internet code page used by the
	 *            item.
	 */
	public void setInternetCodePage(int codePage) {
		
		setProperty("InternetCodePage", codePage);
	}
	
	/**
	 * Returns or sets a boolean value that determines whether the originator of
	 * the meeting item or mail message will receive a delivery report.
	 * Read/write.
	 * <p>
	 * Each transport provider that handles your message sends you a single
	 * delivery notification containing the names and addresses of each
	 * recipient to whom it was delivered. Delivery does not imply that the
	 * message has been read. True if the originator requested a delivery
	 * receipt on the message.
	 * </p>
	 * <p>
	 * The OriginatorDeliveryReportRequested property corresponds to the MAPI
	 * property PidTagOriginatorDeliveryReportRequested.
	 * </p>
	 * 
	 * @return a boolean value that determines whether the originator of the
	 *         meeting item or mail message will receive a delivery report.
	 */
	public boolean isOriginatorDeliveryReportRequested() {
		
		return getBooleanProperty("OriginatorDeliveryReportRequested");
	}
	
	/**
	 * Returns or sets a boolean value that determines whether the originator of
	 * the meeting item or mail message will receive a delivery report.
	 * Read/write.
	 * <p>
	 * Each transport provider that handles your message sends you a single
	 * delivery notification containing the names and addresses of each
	 * recipient to whom it was delivered. Delivery does not imply that the
	 * message has been read. True if the originator requested a delivery
	 * receipt on the message.
	 * </p>
	 * <p>
	 * The OriginatorDeliveryReportRequested property corresponds to the MAPI
	 * property PidTagOriginatorDeliveryReportRequested.
	 * </p>
	 * 
	 * @param flag
	 *            a boolean value that determines whether the originator of the
	 *            meeting item or mail message will receive a delivery report.
	 */
	public void setOriginatorDeliveryReportRequested(boolean flag) {
		
		setProperty("OriginatorDeliveryReportRequested", flag);
	}
	
	/**
	 * Sets or returns a Permission constant that determines what permissions to
	 * grant to the recipients of the e-mail item. Read/write.
	 * <p>
	 * The Permission property should be synchronized with the
	 * PermissionTemplateGuid property to accurately reflect the permission
	 * status of the MailItem. Setting the PermissionTemplateGuid property to a
	 * valid GUID also sets the Permission property to
	 * OlPermission.olPermissionTemplate.
	 * </p>
	 * <p>
	 * When no Information Rights Management (IRM) has been set up, (in which
	 * case the Permission property is OlPermission.olUnrestricted), or the
	 * restriction is not to forward the MailItem, (in which case the Permission
	 * property is OlPermission.olDoNotForward), the value of the
	 * PermissionTemplateGuid property should be an empty string.
	 * </p>
	 * <p>
	 * Although you can view content that is protected by IRM on any computer
	 * that is running the 2007 Microsoft Office system or a later version, you
	 * must have Microsoft Office Professional Edition 2003, Microsoft Office
	 * Outlook 2007, or a later version of Outlook to create or send an e-mail
	 * that is protected by IRM.
	 * </p>
	 * 
	 * @return a Permission constant that determines what permissions to grant
	 *         to the recipients of the e-mail item.
	 */
	public Permission getPermission() {
		
		return Permission.parse(getShortProperty("Permission"));
	}
	
	/**
	 * Sets or returns a Permission constant that determines what permissions to
	 * grant to the recipients of the e-mail item. Read/write.
	 * <p>
	 * The Permission property should be synchronized with the
	 * PermissionTemplateGuid property to accurately reflect the permission
	 * status of the MailItem. Setting the PermissionTemplateGuid property to a
	 * valid GUID also sets the Permission property to
	 * OlPermission.olPermissionTemplate.
	 * </p>
	 * <p>
	 * When no Information Rights Management (IRM) has been set up, (in which
	 * case the Permission property is OlPermission.olUnrestricted), or the
	 * restriction is not to forward the MailItem, (in which case the Permission
	 * property is OlPermission.olDoNotForward), the value of the
	 * PermissionTemplateGuid property should be an empty string.
	 * </p>
	 * <p>
	 * Although you can view content that is protected by IRM on any computer
	 * that is running the 2007 Microsoft Office system or a later version, you
	 * must have Microsoft Office Professional Edition 2003, Microsoft Office
	 * Outlook 2007, or a later version of Outlook to create or send an e-mail
	 * that is protected by IRM.
	 * </p>
	 * 
	 * @param permission
	 *            a Permission constant that determines what permissions to
	 *            grant to the recipients of the e-mail item.
	 */
	public void setPermission(Permission permission) {
		
		setProperty("Permission", permission.value());
	}
	
	/**
	 * Sets or returns a PermissionService constant that determines the
	 * permission service that will be used when sending a message protected by
	 * Information Rights Management (IRM). Read/write.
	 * <p>
	 * This property is useful only if you have more than one permission
	 * identity for a particular SMTP address.
	 * </p>
	 * <p>
	 * While you can view content that is protected by IRM on any computer
	 * running the 2007 Microsoft Office system or a later version, you must
	 * have Microsoft Office Professional Edition 2003, Microsoft Office Outlook
	 * 2007, or a later version of Outlook to create or send an e-mail that is
	 * protected by IRM.
	 * </p>
	 * 
	 * @return a PermissionService constant that determines the permission
	 *         service that will be used when sending a message protected by
	 *         Information Rights Management (IRM).
	 */
	public PermissionService getPermissionService() {
		
		return PermissionService.parse(getShortProperty("PermissionService"));
	}
	
	/**
	 * Sets or returns a PermissionService constant that determines the
	 * permission service that will be used when sending a message protected by
	 * Information Rights Management (IRM). Read/write.
	 * <p>
	 * This property is useful only if you have more than one permission
	 * identity for a particular SMTP address.
	 * </p>
	 * <p>
	 * While you can view content that is protected by IRM on any computer
	 * running the 2007 Microsoft Office system or a later version, you must
	 * have Microsoft Office Professional Edition 2003, Microsoft Office Outlook
	 * 2007, or a later version of Outlook to create or send an e-mail that is
	 * protected by IRM.
	 * </p>
	 * 
	 * @param permission
	 *            a PermissionService constant that determines the permission
	 *            service that will be used when sending a message protected by
	 *            Information Rights Management (IRM).
	 */
	public void setPermissionService(PermissionService permission) {
		
		setProperty("PermissionService", permission.value());
	}
	
	/**
	 * Returns or sets the GUID of the template file to apply to the MailItem in
	 * order to specify Information Rights Management (IRM) permissions.
	 * Read/write.
	 * <p>
	 * This property complements the IRM properties on a MailItem object; that
	 * is, the Permission property and the PermissionService properties.
	 * </p>
	 * <p>
	 * In particular, the PermissionTemplateGuid property should be synchronized
	 * with the Permission property to accurately reflect the permission status
	 * of the MailItem. Setting the PermissionTemplateGuid property to a valid
	 * GUID should also incur setting the Permission property to
	 * OlPermission.olPermissionTemplate.
	 * </p>
	 * <p>
	 * An empty string value for the PermissionTemplateGuid property means that
	 * there is no permission template file specified for the MailItem. For
	 * example, if no IRM has been set up (in which case the Permission property
	 * is OlPermission.olUnrestricted), or the restriction is not to forward the
	 * MailItem (in which case the Permission property is
	 * OlPermission.olDoNotForward).
	 * </p>
	 * <p>
	 * If you attempt to set the PermissionTemplateGuid property for a received
	 * message (that is, the Sent property of the MailItem is True), Microsoft
	 * Outlook returns an error.
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @return the GUID of the template file to apply to the MailItem in order
	 *         to specify Information Rights Management (IRM) permissions.
	 */
	public String getPermissionTemplateGuid() {
		
		return getStringProperty("PermissionTemplateGuid");
	}
	
	/**
	 * Returns or sets the GUID of the template file to apply to the MailItem in
	 * order to specify Information Rights Management (IRM) permissions.
	 * Read/write.
	 * <p>
	 * This property complements the IRM properties on a MailItem object; that
	 * is, the Permission property and the PermissionService properties.
	 * </p>
	 * <p>
	 * In particular, the PermissionTemplateGuid property should be synchronized
	 * with the Permission property to accurately reflect the permission status
	 * of the MailItem. Setting the PermissionTemplateGuid property to a valid
	 * GUID should also incur setting the Permission property to
	 * OlPermission.olPermissionTemplate.
	 * </p>
	 * <p>
	 * An empty string value for the PermissionTemplateGuid property means that
	 * there is no permission template file specified for the MailItem. For
	 * example, if no IRM has been set up (in which case the Permission property
	 * is OlPermission.olUnrestricted), or the restriction is not to forward the
	 * MailItem (in which case the Permission property is
	 * OlPermission.olDoNotForward).
	 * </p>
	 * <p>
	 * If you attempt to set the PermissionTemplateGuid property for a received
	 * message (that is, the Sent property of the MailItem is True), Microsoft
	 * Outlook returns an error.
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @param guid
	 *            the GUID of the template file to apply to the MailItem in
	 *            order to specify Information Rights Management (IRM)
	 *            permissions.
	 */
	public void setPermissionTemplateGuid(String guid) {
		
		setProperty("PermissionTemplateGuid", guid);
	}
	
	/**
	 * Returns a boolean value that indicates true if a read receipt has been
	 * requested by the sender.
	 * <p>
	 * This property corresponds to the MAPI property
	 * PidTagReadReceiptRequested. Read/write for e-mail items that have been
	 * created but have not been sent or posted; read-only for sent e-mail
	 * items.
	 * </p>
	 * 
	 * @return a boolean value that indicates true if a read receipt has been
	 *         requested by the sender.
	 */
	public boolean isReadReceiptRequested() {
		
		return getBooleanProperty("ReadReceiptRequested");
	}
	
	/**
	 * Returns a boolean value that indicates true if a read receipt has been
	 * requested by the sender.
	 * <p>
	 * This property corresponds to the MAPI property
	 * PidTagReadReceiptRequested. Read/write for e-mail items that have been
	 * created but have not been sent or posted; read-only for sent e-mail
	 * items.
	 * </p>
	 * 
	 * @param flag
	 *            a boolean value that indicates true if a read receipt has been
	 *            requested by the sender.
	 */
	public void setReadReceiptRequested(boolean flag) {
		
		setProperty("ReadReceiptRequested", flag);
	}
	
	/**
	 * Returns a String representing the EntryID for the true recipient as set
	 * by the transport provider delivering the mail message. Read-only.
	 * <p>
	 * This property corresponds to the MAPI property PidTagReceivedByEntryId.
	 * </p>
	 * <p>
	 * If you are getting this property in a Microsoft Visual Basic or Microsoft
	 * Visual Basic for Applications (VBA) solution, owing to some type issues,
	 * instead of directly referencing ReceivedByEntryID, you should get the
	 * property through the PropertyAccessor object returned by the
	 * MailItem.PropertyAccessor property, specifying the
	 * PidTagReceivedByEntryId property and its MAPI proptag namespace. The
	 * following code sample in VBA shows the workaround.
	 * </p>
	 * 
	 * @return a String representing the EntryID for the true recipient as set
	 *         by the transport provider delivering the mail message.
	 */
	public String getReceivedByEntryID() {
		
		return getStringProperty("ReceivedByEntryID");
	}
	
	/**
	 * Returns a String representing the display name of the true recipient for
	 * the mail message. Read-only.
	 * <p>
	 * This property corresponds to the MAPI property PidTagReceivedByName.
	 * </p>
	 * 
	 * @return a String representing the display name of the true recipient for
	 *         the mail message.
	 */
	public String getReceivedByName() {
		
		return getStringProperty("ReceivedByName");
	}
	
	/**
	 * Returns a String representing the EntryID of the user delegated to
	 * represent the recipient for the mail message. Read-only.
	 * <p>
	 * This property corresponds to the MAPI property
	 * PidTagReceivedRepresentingEntryId.
	 * </p>
	 * <p>
	 * If you are getting this property in a Microsoft Visual Basic or Microsoft
	 * Visual Basic for Applications (VBA) solution, owing to some type issues,
	 * instead of directly referencing ReceivedOnBehalfOfEntryID, you should get
	 * the property through the PropertyAccessor object returned by the
	 * MailItem.PropertyAccessor property, specifying the MAPI property
	 * PidTagReceivedRepresentingEntryId property and its MAPI proptag
	 * namespace. The following code sample in VBA shows the workaround.
	 * </p>
	 * 
	 * @return a String representing the EntryID of the user delegated to
	 *         represent the recipient for the mail message.
	 */
	public String getReceivedOnBehalfOfEntryID() {
		
		return getStringProperty("ReceivedOnBehalfOfEntryID");
	}
	
	/**
	 * Returns a String representing the display name of the user delegated to
	 * represent the recipient for the mail message. Read-only.
	 * <p>
	 * This property corresponds to the MAPI property
	 * PidTagReceivedRepresentingName.
	 * </p>
	 * 
	 * @return a String representing the display name of the user delegated to
	 *         represent the recipient for the mail message.
	 */
	public String getReceivedOnBehalfOfName() {
		
		return getStringProperty("ReceivedOnBehalfOfName");
	}
	
	/**
	 * Returns a Date indicating the date and time at which the item was
	 * received. Read-only.
	 * 
	 * @return a Date indicating the date and time at which the item was
	 *         received.
	 */
	public Date getReceivedTime() {
		
		return getDateProperty("ReceivedTime");
	}
	
	/**
	 * Returns a boolean that indicates true if the recipient cannot forward the
	 * mail message. Read/write.
	 * 
	 * @return a boolean that indicates true if the recipient cannot forward the
	 *         mail message.
	 */
	public boolean isRecipientReassignmentProhibited() {
		
		return getBooleanProperty("RecipientReassignmentProhibited");
	}
	
	/**
	 * Returns a boolean that indicates true if the recipient cannot forward the
	 * mail message. Read/write.
	 * 
	 * @param flag
	 *            a boolean that indicates true if the recipient cannot forward
	 *            the mail message.
	 */
	public void setRecipientReassignmentProhibited(boolean flag) {
		
		setProperty("RecipientReassignmentProhibited", flag);
	}
	
	/**
	 * Returns a Recipients collection that represents all the recipients for
	 * the Outlook item (item: An item is the basic element that holds
	 * information in Outlook (similar to a file in other programs). Items
	 * include e-mail messages, appointments, contacts, tasks, journal entries,
	 * notes, posted items, and documents.). Read-only.
	 * <p>
	 * A recipient can be specified by a string representing the recipient's
	 * display name, alias, or full SMTP e-mail address.
	 * </p>
	 * 
	 * @return a Recipients collection that represents all the recipients for
	 *         the Outlook item.
	 */
	public Recipients getRecipients() {
		
		return new Recipients(getAutomationProperty("Recipients"));
	}
	
	/**
	 * Returns or sets a RemoteStatus constant specifying the remote status of
	 * the mail message. Read/write.
	 * 
	 * @return a RemoteStatus constant specifying the remote status of the mail
	 *         message.
	 */
	public RemoteStatus getRemoteStatus() {
		
		return RemoteStatus.parse(getShortProperty("RemoteStatus"));
	}
	
	/**
	 * Returns or sets a RemoteStatus constant specifying the remote status of
	 * the mail message. Read/write.
	 * 
	 * @param status
	 *            a RemoteStatus constant specifying the remote status of the
	 *            mail message.
	 */
	public void setRemoteStatus(RemoteStatus status) {
		
		setProperty("RemoteStatus", status.value());
	}
	
	/**
	 * Returns a semicolon-delimited String list of reply recipients for the
	 * mail message. Read-only.
	 * <p>
	 * This property only contains the display names for the reply recipients.
	 * The reply recipients list should be set by using the ReplyRecipients
	 * collection.
	 * </p>
	 * 
	 * @return a semicolon-delimited String list of reply recipients for the
	 *         mail message.
	 */
	public String getReplyRecipientNames() {
		
		return getStringProperty("ReplyRecipientNames");
	}
	
	/**
	 * Returns a Recipients collection that represents all the reply recipient
	 * objects for the Outlook item. Read-only.
	 * 
	 * @return a Recipients collection that represents all the reply recipient
	 *         objects for the Outlook item.
	 */
	public Recipients getReplyRecipients() {
		
		return new Recipients(getAutomationProperty("ReplyRecipients"));
	}
	
	/**
	 * Returns a Date that specifies the date when the MailItem object expires,
	 * after which the Messaging Records Management (MRM) Assistant will delete
	 * the item. Read-only.
	 * <p>
	 * A retention policy is enabled and disabled by an administrator for a
	 * Microsoft Exchange Server on a mailbox level. This feature is available
	 * only on an Exchange Server 2010 mailbox with MRM 2.0 enabled.
	 * </p>
	 * <p>
	 * Microsoft Outlook calculates the value of this property based on the item
	 * retention start date and the retention period, if Outlook is in cache or
	 * offline mode. The Exchange Server specifies the value if Outlook is in
	 * online mode.
	 * </p>
	 * <p>
	 * In general, the retention start date for the item is determined as
	 * follows:
	 * <ul>
	 * <li>Received or sent items: the retention start date is the received
	 * date.</li>
	 * 
	 * <li>Nonrecurring calendar items: the retention start date is the
	 * appointment end date.</li>
	 * 
	 * <li>Recurring calendar items: the retention start date is the end date of
	 * last recurrence. If there is no end date, the item never expires.</li>
	 * </ul>
	 * </p>
	 * <p>
	 * Added in Outlook 2010.
	 * </p>
	 * 
	 * @return a Date that specifies the date when the MailItem object expires,
	 *         after which the Messaging Records Management (MRM) Assistant will
	 *         delete the item.
	 */
	public Date getRetentionExpirationDate() {
		
		return getDateProperty("RetentionExpirationDate");
	}
	
	/**
	 * Returns a String that specifies the name of the retention policy.
	 * Read-only.
	 * <p>
	 * Retention is enabled and disabled by an administrator for a Microsoft
	 * Exchange Server on a mailbox level. The feature is available only on an
	 * Exchange Server 2010 mailbox with Messaging Records Management (MRM) 2.0
	 * enabled. An example of a retention policy name is
	 * "Define time interval for expiration Quick Searches".
	 * </p>
	 * 
	 * @return a String that specifies the name of the retention policy.
	 */
	public String getRetentionPolicyName() {
		
		return getStringProperty("RetentionPolicyName");
	}
	
	/**
	 * Returns or sets a Folder object that represents the folder in which a
	 * copy of the e-mail message will be saved after being sent. Read/write.
	 * 
	 * @return a Folder object that represents the folder in which a copy of the
	 *         e-mail message will be saved after being sent.
	 */
	public Folder getSaveSentMessageFolder() {
		
		return new Folder(getAutomationProperty("SaveSentMessageFolder"));
	}
	
	/**
	 * Returns or sets a Folder object that represents the folder in which a
	 * copy of the e-mail message will be saved after being sent. Read/write.
	 * 
	 * @param folder
	 *            a Folder object that represents the folder in which a copy of
	 *            the e-mail message will be saved after being sent.
	 */
	public void setSaveSentMessageFolder(Folder folder) {
		
		setProperty("SaveSentMessageFolder", folder.getIDispatch());
	}
	
	/**
	 * Sends the e-mail message.
	 */
	public void send() {
		
		invokeNoReply("Send");
	}
	
	/**
	 * Returns a String that represents the e-mail address of the sender of the
	 * Outlook item. Read-only.
	 * <p>
	 * This property corresponds to the MAPI property PidTagSenderEmailAddress.
	 * </p>
	 * 
	 * @return a String that represents the e-mail address of the sender of the
	 *         Outlook item.
	 */
	public String getSenderEmailAddress() {
		
		return getStringProperty("SenderEmailAddress");
	}
	
	/**
	 * Returns a String that represents the type of entry for the e-mail address
	 * of the sender of the Outlook item, such as 'SMTP' for Internet address,
	 * 'EX' for a Microsoft Exchange server address, etc. Read-only.
	 * 
	 * @return a String that represents the type of entry for the e-mail address
	 *         of the sender of the Outlook item, such as 'SMTP' for Internet
	 *         address, 'EX' for a Microsoft Exchange server address, etc.
	 */
	public String getSenderEmailType() {
		
		return getStringProperty("SenderEmailType");
	}
	
	/**
	 * Returns a String indicating the display name of the sender for the
	 * Outlook item. Read-only.
	 * <p>
	 * This property corresponds to the MAPI property PidTagSenderName.
	 * </p>
	 * <p>
	 * If you wish to retrieve the fully qualified e-mail address of the sender,
	 * use the SenderEmailAddress property.
	 * </p>
	 * 
	 * @return a String indicating the display name of the sender for the
	 *         Outlook item.
	 */
	public String getSenderName() {
		
		return getStringProperty("SenderName");
	}
	
	/**
	 * Returns or sets an Account object that represents the account under which
	 * the MailItem is to be sent. Read/write.
	 * <p>
	 * The SendUsingAccount property can be used to specify the account that
	 * should be used to send the MailItem when the Send method is called. This
	 * property returns Null (Nothing in Visual Basic) if the account specified
	 * for the MailItem no longer exists.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @return an Account object that represents the account under which the
	 *         MailItem is to be sent.
	 */
	public Account getSendUsingAccount() {
		
		return new Account(getAutomationProperty("SendUsingAccount"));
	}
	
	/**
	 * Returns or sets an Account object that represents the account under which
	 * the MailItem is to be sent. Read/write.
	 * <p>
	 * The SendUsingAccount property can be used to specify the account that
	 * should be used to send the MailItem when the Send method is called. This
	 * property returns Null (Nothing in Visual Basic) if the account specified
	 * for the MailItem no longer exists.
	 * </p>
	 * <p>
	 * Added in Outlook 2007.
	 * </p>
	 * 
	 * @param acct
	 *            an Account object that represents the account under which the
	 *            MailItem is to be sent.
	 */
	public void setSendUsingAccount(Account acct) {
		
		setProperty("SendUsingAccount", acct.getIDispatch());
	}
	
	/**
	 * Returns a boolean value that indicates if a message has been sent.
	 * Read-only.
	 * <p>
	 * In general, there are three different kinds of messages: sent, posted,
	 * and saved. Sent messages are items sent to a recipient or public folder.
	 * Posted messages are created in a public folder. Saved messages are
	 * created and saved without either sending or posting.
	 * </p>
	 * 
	 * @return a boolean value that indicates if a message has been sent.
	 */
	public boolean isSent() {
		
		return getBooleanProperty("Sent");
	}
	
	/**
	 * Returns a Date indicating the date and time on which the Outlook item was
	 * sent. Read-only.
	 * <p>
	 * This property corresponds to the MAPI property PidTagClientSubmitTime.
	 * When you send an item using the object's Send method, the transport
	 * provider sets the ReceivedTime and SentOn properties for you.
	 * </p>
	 * 
	 * @return a Date indicating the date and time on which the Outlook item was
	 *         sent.
	 */
	public Date getSentOn() {
		
		return getDateProperty("SentOn");
	}
	
	/**
	 * Returns a String indicating the display name for the intended sender of
	 * the mail message. Read/write.
	 * <p>
	 * This property corresponds to the MAPI property
	 * PidTagSentRepresentingName.
	 * </p>
	 * 
	 * @return a String indicating the display name for the intended sender of
	 *         the mail message.
	 */
	public String getSentOnBehalfOfName() {
		
		return getStringProperty("SentOnBehalfOfName");
	}
	
	/**
	 * Returns a String indicating the display name for the intended sender of
	 * the mail message. Read/write.
	 * <p>
	 * This property corresponds to the MAPI property
	 * PidTagSentRepresentingName.
	 * </p>
	 * 
	 * @param name
	 *            a String indicating the display name for the intended sender
	 *            of the mail message.
	 */
	public void setSentOnBehalfOfName(String name) {
		
		setProperty("SentOnBehalfOfName", name);
	}
	
	/**
	 * Returns a boolean value that is true if the item has been submitted
	 * (submitted: When a message is submitted, the store provider places the
	 * message in its outgoing queue, where it gets picked up by the spooler and
	 * handed to one or more transport providers for delivery.). Read-only.
	 * <p>
	 * A message is always created and submitted in a folder, usually the
	 * Outbox.
	 * </p>
	 * 
	 * @return a boolean value that is true if the item has been submitted.
	 */
	public boolean isSubmitted() {
		
		return getBooleanProperty("Submitted");
	}
	
	/**
	 * Returns or sets a semicolon-delimited String list of display names for
	 * the To recipients for the Outlook item. Read/write.
	 * <p>
	 * This property contains the display names only. The To property
	 * corresponds to the MAPI property PidTagDisplayTo. The Recipients
	 * collection should be used to modify this property.
	 * </p>
	 * 
	 * @return a semicolon-delimited String list of display names for the To
	 *         recipients for the Outlook item.
	 */
	public String getTo() {
		
		return getStringProperty("To");
	}
	
	/**
	 * Returns or sets a semicolon-delimited String list of display names for
	 * the To recipients for the Outlook item. Read/write.
	 * <p>
	 * This property contains the display names only. The To property
	 * corresponds to the MAPI property PidTagDisplayTo. The Recipients
	 * collection should be used to modify this property.
	 * </p>
	 * 
	 * @param to
	 *            a semicolon-delimited String list of display names for the To
	 *            recipients for the Outlook item.
	 */
	public void setTo(String to) {
		
		setProperty("To", to);
	}
	
}
