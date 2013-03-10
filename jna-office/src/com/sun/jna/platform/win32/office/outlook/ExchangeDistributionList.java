package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class ExchangeDistributionList extends AddressEntry {

	ExchangeDistributionList(IDispatch iDisp) {
		super(iDisp);
	}
	
	public String getAlias() {
		
		return getStringProperty("Alias");
	}
	
	public String getComments() {
		
		return getStringProperty("Comments");
	}
	
	public void setComments(String comment) {
		
		setProperty("Comments", comment);
	}
	
	public AddressEntries getExchangeDistributionListMembers() {
		
		return new AddressEntries((IDispatch) invoke("GetExchangeDistributionListMembers").getValue());
	}
	
	public AddressEntries getMemberOfList() {
		
		return new AddressEntries((IDispatch) invoke("GetMemberOfList").getValue());
	}
	
	public AddressEntries getOwners() {
		
		return new AddressEntries((IDispatch) invoke("GetOwners").getValue());
	}
	
	public String getPrimarySmtpAddress() {
		
		return getStringProperty("PrimarySmtpAddress");
	}
	
}
