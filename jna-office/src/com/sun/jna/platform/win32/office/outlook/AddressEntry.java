package com.sun.jna.platform.win32.office.outlook;

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;

public class AddressEntry extends BaseOutlookObject {

	AddressEntry(IDispatch iDisp) {
		super(iDisp);
	}
	
	public String getAddress() {
		
		return getStringProperty("Address");
	}
	
	public void setAddress(String address) {
		
		setProperty("Address", address);
	}
	
	public AddressEntryUserType getAddressEntryUserType() {
		
		return AddressEntryUserType.parse(getShortProperty("AddressEntryUserType"));
	}
	
	public void delete() {
		
		invokeNoReply("Delete");
	}
	
	public void showDetails() {
		
		invokeNoReply("Details");
	}
	
	public void showDetails(int hWnd) {
		
		invokeNoReply("Details", newVariant(hWnd));
	}
	
	public DisplayType getDisplayType() {
		
		return DisplayType.parse(getShortProperty("DisplayType"));
	}
	
	public ContactItem getContact() {
		
		return new ContactItem((IDispatch) invoke("GetContact").getValue());
	}
	
	public ExchangeDistributionList getExchangeDistributionList() {
		
		return new ExchangeDistributionList((IDispatch) invoke("GetExchangeDistributionList").getValue());
	}
	
	public ExchangeUser getExchangeUser() {
		
		return new ExchangeUser((IDispatch) invoke("GetExchangeUser").getValue());
	}
	
	public String getFreeBusy(Date start, int minsPerChar) {
		
		return getFreeBusy(start, minsPerChar, false);
	}
	
	public String getFreeBusy(Date start, int minsPerChar, boolean useCompleteFormat) {
		
		return invoke("GetFreeBusy", newVariant(start), newVariant(minsPerChar), newVariant(useCompleteFormat)).getValue().toString();
	}
	
	public String getID() {
		
		return getStringProperty("ID");
	}
	
	public String getName() {
		
		return getStringProperty("Name");
	}
	
	public void setName(String name) {
		
		setProperty("Name", name);
	}
	
	public PropertyAccessor getPropertyAccessor() {
		
		return new PropertyAccessor(getAutomationProperty("PropertyAccessor"));
	}
	
	public String getType() {
		
		return getStringProperty("Type");
	}
	
	public void setType(String typ) {
		
		setProperty("Type", typ);
	}
	
	public void update() {
		
		update(true, false);
	}
	
	public void update(boolean makePermanent) {
		
		update(makePermanent, false);
	}
	
	public void update(boolean makePermanent, boolean refresh) {
		
		invokeNoReply("Update", newVariant(makePermanent), newVariant(refresh));
	}
	
}
