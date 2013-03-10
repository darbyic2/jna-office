package com.sun.jna.platform.win32.office.outlook;

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;

public class ExchangeUser extends BaseOutlookObject {

	ExchangeUser(IDispatch iDisp) {
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
	
	public String getAlias() {
		
		return getStringProperty("Alias");
	}
	
	public String getAssistantName() {
		
		return getStringProperty("AssistantName");
	}
	
	public void setAssistantName(String name) {
		
		setProperty("AssistantName", name);
	}
	
	public String getBusinessTelephoneNumber() {
		
		return getStringProperty("BusinessTelephoneNumber");
	}
	
	public void setBusinessTelephoneNumber(String tel) {
		
		setProperty("BusinessTelephoneNumber", tel);
	}
	
	public String getCity() {
		
		return getStringProperty("City");
	}
	
	public void setCity(String city) {
		
		setProperty("City", city);
	}
	
	public String getComments() {
		
		return getStringProperty("Comments");
	}
	
	public void setComments(String comment) {
		
		setProperty("Comments", comment);
	}
	
	public String getCompanyName() {
		
		return getStringProperty("CompanyName");
	}
	
	public void setCompanyName(String name) {
		
		setProperty("CompanyName", name);
	}
	
	public void delete() {
		
		invokeNoReply("Delete");
	}
	
	public String getDepartment() {
		
		return getStringProperty("Department");
	}
	
	public void setDepartment(String name) {
		
		setProperty("Department", name);
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
	
	public String getFirstName() {
		
		return getStringProperty("FirstName");
	}
	
	public void setFirstName(String name) {
		
		setProperty("FirstName", name);
	}
	
	public ContactItem getContact() {
		
		return new ContactItem((IDispatch) invoke("GetContact").getValue());
	}
	
	public AddressEntries getDirectReports() {
		
		return new AddressEntries((IDispatch) invoke("GetDirectReports").getValue());
	}
	
	public ExchangeDistributionList getExchangeDistributionList() {
		
		return new ExchangeDistributionList((IDispatch) invoke("GetExchangeDistributionList").getValue());
	}
	
	public ExchangeUser getParentExchangeUser() {
		
		return new ExchangeUser((IDispatch) invoke("GetExchangeUser").getValue());
	}
	
	public ExchangeUser getExchangeUserManager() {
		
		return new ExchangeUser((IDispatch) invoke("GetExchangeUserManager").getValue());
	}
	
	public String getFreeBusy(Date start, int minsPerChar) {
		
		return getFreeBusy(start, minsPerChar, false);
	}
	
	public String getFreeBusy(Date start, int minsPerChar, boolean useCompleteFormat) {
		
		return invoke("GetFreeBusy", newVariant(start), newVariant(minsPerChar), newVariant(useCompleteFormat)).getValue().toString();
	}
	
	public AddressEntries getMemberOfList() {
		
		return new AddressEntries((IDispatch) invoke("GetMemberOfList").getValue());
	}
	
	public String getID() {
		
		return getStringProperty("ID");
	}
	
	public String getJobTitle() {
		
		return getStringProperty("JobTitle");
	}
	
	public void setJobTitle(String title) {
		
		setProperty("JobTitle", title);
	}
	
	public String getLastName() {
		
		return getStringProperty("LastName");
	}
	
	public void setLastName(String name) {
		
		setProperty("LastName", name);
	}
	
	public String getMobileTelephoneNumber() {
		
		return getStringProperty("MobileTelephoneNumber");
	}
	
	public void setMobileTelephoneNumber(String tel) {
		
		setProperty("MobileTelephoneNumber", tel);
	}
	
	public String getName() {
		
		return getStringProperty("Name");
	}
	
	public void setName(String name) {
		
		setProperty("Name", name);
	}
	
	public String getOfficeLocation() {
		
		return getStringProperty("OfficeLocation");
	}
	
	public void setOfficeLocation(String location) {
		
		setProperty("OfficeLocation", location);
	}
	
	public String getPostalCode() {
		
		return getStringProperty("PostalCode");
	}
	
	public void setPostalCode(String postCode) {
		
		setProperty("PostalCode", postCode);
	}
	
	public String getPrimarySmtpAddress() {
		
		return getStringProperty("PrimarySmtpAddress");
	}
	
	public PropertyAccessor getPropertyAccessor() {
		
		return new PropertyAccessor(getAutomationProperty("PropertyAccessor"));
	}
	
	public String getStateOrProvince() {
		
		return getStringProperty("StateOrProvince");
	}
	
	public void setStateOrProvince(String stateOrProvince) {
		
		setProperty("StateOrProvince", stateOrProvince);
	}
	
	public String getStreetAddress() {
		
		return getStringProperty("StreetAddress");
	}
	
	public void setStreetAddress(String street) {
		
		setProperty("StreetAddress", street);
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
	
	public String getYomiCompanyName() {
		
		return getStringProperty("YomiCompanyName");
	}
	
	public void setYomiCompanyName(String val) {
		
		setProperty("YomiCompanyName", val);
	}
	
	public String getYomiDepartment() {
		
		return getStringProperty("YomiDepartment");
	}
	
	public void setYomiDepartment(String val) {
		
		setProperty("YomiDepartment", val);
	}
	
	public String getYomiDisplayName() {
		
		return getStringProperty("YomiDisplayName");
	}
	
	public void setYomiDisplayName(String val) {
		
		setProperty("YomiDisplayName", val);
	}
	
	public String getYomiFirstName() {
		
		return getStringProperty("YomiFirstName");
	}
	
	public void setYomiFirstName(String val) {
		
		setProperty("YomiFirstName", val);
	}
	
	public String getYomiLastName() {
		
		return getStringProperty("YomiLastName");
	}
	
	public void setYomiLastName(String val) {
		
		setProperty("YomiLastName", val);
	}
	
}
