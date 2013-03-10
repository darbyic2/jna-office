package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class Account extends BaseOutlookObject {

	Account(IDispatch iDisp) {
		super(iDisp);
	}
	
	public AccountType getAccountType() {
		
		return AccountType.parse(getShortProperty("AccountType"));
	}
	
	public AutoDiscoverConnectionMode getAutoDiscoverConnectionMode() {
		
		return AutoDiscoverConnectionMode.parse(getShortProperty("AutoDiscoverConnectionMode"));
	}
	
	public String getAutoDiscoverXML() {
		
		return getStringProperty("AutoDiscoverXML");
	}
	
	public Recipient getCurrentUser() {
		
		return new Recipient(getAutomationProperty("CurrentUser"));
	}
	
	public Store getDeliveryStore() {
		
		return new Store(getAutomationProperty("DeliveryStore"));
	}
	
	public String getDisplayName() {
		
		return getStringProperty("DisplayName");
	}
	
	public ExchangeConnectionMode getExchangeConnectionMode() {
		
		return ExchangeConnectionMode.parse(getShortProperty("ExchangeConnectionMode"));
	}
	
	public String getExchangeMailboxServerName() {
		
		return getStringProperty("ExchangeMailboxServerName");
	}
	
	public String getExchangeMailboxServerVersion() {
		
		return getStringProperty("ExchangeMailboxServerVersion");
	}
	
	public AddressEntry GetAddressEntryFromID(String id) {
		
		return new AddressEntry((IDispatch) invoke("GetAddressEntryFromID", newVariant(id)).getValue());
	}
	
	public Recipient GetRecipientFromID(String id) {
		
		return new Recipient((IDispatch) invoke("GetRecipientFromID", newVariant(id)).getValue());
	}
	
	public String getSmtpAddress() {
		
		return getStringProperty("SmtpAddress");
	}
	
	public String getUserName() {
		
		return getStringProperty("UserName");
	}
	
}
