package com.sun.jna.platform.win32.office.outlook;

import java.util.Date;

import com.sun.jna.platform.win32.COM.IDispatch;

public class ContactItem extends BaseItemLevel3 {

	ContactItem(IDispatch iDisp) {
		super(iDisp);
	}
	
	public String getAccount() {
		
		return getStringProperty("Account");
	}
	
	public void setAccount(String ac) {
		
		setProperty("Account", ac);
	}
	
	public void addBusinessCardLogoPicture(String path) {
		
		invokeNoReply("AddBusinessCardLogoPicture", newVariant(path));
	}
	
	public void addPicture(String path) {
		
		invokeNoReply("AddPicture", newVariant(path));
	}
	
	public Date getAnniversary() {
		
		return getDateProperty("Anniversary");
	}
	
	public void setAnniversary(Date dat) {
		
		setProperty("Anniversary", dat);
	}
	
	public String getAssistantName() {
		
		return getStringProperty("AssistantName");
	}
	
	public void setAssistantName(String name) {
		
		setProperty("AssistantName", name);
	}
	
	public String getAssistantTelephoneNumber() {
		
		return getStringProperty("AssistantTelephoneNumber");
	}
	
	public void setAssistantTelephoneNumber(String tel) {
		
		setProperty("AssistantTelephoneNumber", tel);
	}
	
	public Date getBirthday() {
		
		return getDateProperty("Birthday");
	}
	
	public void setBirthday(Date dat) {
		
		setProperty("Birthday", dat);
	}
	
	public String getBusiness2TelephoneNumber() {
		
		return getStringProperty("Business2TelephoneNumber");
	}
	
	public void setBusiness2TelephoneNumber(String tel) {
		
		setProperty("Business2TelephoneNumber", tel);
	}
	
	public String getBusinessAddress() {
		
		return getStringProperty("BusinessAddress");
	}
	
	public void setBusinessAddress(String addr) {
		
		setProperty("BusinessAddress", addr);
	}
	
	public String getBusinessAddressCity() {
		
		return getStringProperty("BusinessAddressCity");
	}
	
	public void setBusinessAddressCity(String addr) {
		
		setProperty("BusinessAddressCity", addr);
	}
	
	public String getBusinessAddressCountry() {
		
		return getStringProperty("BusinessAddressCountry");
	}
	
	public void setBusinessAddressCountry(String addr) {
		
		setProperty("BusinessAddressCountry", addr);
	}
	
	public String getBusinessAddressPostalCode() {
		
		return getStringProperty("BusinessAddressPostalCode");
	}
	
	public void setBusinessAddressPostalCode(String code) {
		
		setProperty("BusinessAddressPostalCode", code);
	}
	
	public String getBusinessAddressPostOfficeBox() {
		
		return getStringProperty("BusinessAddressPostOfficeBox");
	}
	
	public void setBusinessAddressPostOfficeBox(String boxNum) {
		
		setProperty("BusinessAddressPostOfficeBox", boxNum);
	}
	
	public String getBusinessAddressState() {
		
		return getStringProperty("BusinessAddressState");
	}
	
	public void setBusinessAddressState(String state) {
		
		setProperty("BusinessAddressState", state);
	}
	
	public String getBusinessAddressStreet() {
		
		return getStringProperty("BusinessAddressStreet");
	}
	
	public void setBusinessAddressStreet(String street) {
		
		setProperty("BusinessAddressStreet", street);
	}
	
	public String getBusinessCardLayoutXml() {
		
		return getStringProperty("BusinessCardLayoutXml");
	}
	
	public void setBusinessCardLayoutXml(String xml) {
		
		setProperty("BusinessCardLayoutXml", xml);
	}
	
	public BusinessCardType getBusinessCardType() {
		
		return BusinessCardType.parse(getShortProperty("BusinessCardType"));
	}
	
	public String getBusinessFaxNumber() {
		
		return getStringProperty("BusinessFaxNumber");
	}
	
	public void setBusinessFaxNumber(String fax) {
		
		setProperty("BusinessFaxNumber", fax);
	}
	
	public String getBusinessHomePage() {
		
		return getStringProperty("BusinessHomePage");
	}
	
	public void setBusinessHomePage(String fax) {
		
		setProperty("BusinessHomePage", fax);
	}
	
	public String getBusinessTelephoneNumber() {
		
		return getStringProperty("BusinessTelephoneNumber");
	}
	
	public void setBusinessTelephoneNumber(String tel) {
		
		setProperty("BusinessTelephoneNumber", tel);
	}
	
	public String getCallbackTelephoneNumber() {
		
		return getStringProperty("CallbackTelephoneNumber");
	}
	
	public void setCallbackTelephoneNumber(String tel) {
		
		setProperty("CallbackTelephoneNumber", tel);
	}
	
	public String getCarTelephoneNumber() {
		
		return getStringProperty("CarTelephoneNumber");
	}
	
	public void setCarTelephoneNumber(String tel) {
		
		setProperty("CarTelephoneNumber", tel);
	}
	
	public String getChildren() {
		
		return getStringProperty("Children");
	}
	
	public void setChildren(String kids) {
		
		setProperty("Children", kids);
	}
	
	public String getCompanyAndFullName() {
		
		return getStringProperty("CompanyAndFullName");
	}
	
	public String getCompanyLastFirstNoSpace() {
		
		return getStringProperty("CompanyLastFirstNoSpace");
	}
	
	public String getCompanyLastFirstSpaceOnly() {
		
		return getStringProperty("CompanyLastFirstSpaceOnly");
	}
	
	public String getCompanyMainTelephoneNumber() {
		
		return getStringProperty("CompanyMainTelephoneNumber");
	}
	
	public void setCompanyMainTelephoneNumber(String tel) {
		
		setProperty("CompanyMainTelephoneNumber", tel);
	}
	
	public String getCompanyName() {
		
		return getStringProperty("CompanyName");
	}
	
	public void setCompanyName(String name) {
		
		setProperty("CompanyName", name);
	}
	
	public String getComputerNetworkName() {
		
		return getStringProperty("ComputerNetworkName");
	}
	
	public void setComputerNetworkName(String name) {
		
		setProperty("ComputerNetworkName", name);
	}
	
	public String getCustomerID() {
		
		return getStringProperty("CustomerID");
	}
	
	public void setCustomerID(String id) {
		
		setProperty("CustomerID", id);
	}
	
	public String getDepartment() {
		
		return getStringProperty("Department");
	}
	
	public void setDepartment(String name) {
		
		setProperty("Department", name);
	}
	
	public String getEmail1Address() {
		
		return getStringProperty("Email1Address");
	}
	
	public void setEmail1Address(String val) {
		
		setProperty("Email1Address", val);
	}
	
	public String getEmail1AddressType() {
		
		return getStringProperty("Email1AddressType");
	}
	
	public void setEmail1AddressType(String val) {
		
		setProperty("Email1AddressType", val);
	}
	
	public String getEmail1DisplayName() {
		
		return getStringProperty("Email1DisplayName");
	}
	
	public void setEmail1DisplayName(String val) {
		
		setProperty("Email1DisplayName", val);
	}
	
	public String getEmail1EntryID() {
		
		return getStringProperty("Email1EntryID");
	}
	
	public String getEmail2Address() {
		
		return getStringProperty("Email2Address");
	}
	
	public void setEmail2Address(String val) {
		
		setProperty("Email2Address", val);
	}
	
	public String getEmail2AddressType() {
		
		return getStringProperty("Email2AddressType");
	}
	
	public void setEmail2AddressType(String val) {
		
		setProperty("Email2AddressType", val);
	}
	
	public String getEmail2DisplayName() {
		
		return getStringProperty("Email2DisplayName");
	}
	
	public void setEmail2DisplayName(String val) {
		
		setProperty("Email2DisplayName", val);
	}
	
	public String getEmail2EntryID() {
		
		return getStringProperty("Email1EntryID");
	}
	
	public String getEmail3Address() {
		
		return getStringProperty("Email3Address");
	}
	
	public void setEmail3Address(String val) {
		
		setProperty("Email3Address", val);
	}
	
	public String getEmail3AddressType() {
		
		return getStringProperty("Email3AddressType");
	}
	
	public void setEmail3AddressType(String val) {
		
		setProperty("Email3AddressType", val);
	}
	
	public String getEmail3DisplayName() {
		
		return getStringProperty("Email3DisplayName");
	}
	
	public void setEmail3DisplayName(String val) {
		
		setProperty("Email3DisplayName", val);
	}
	
	public String getEmail3EntryID() {
		
		return getStringProperty("Email1EntryID");
	}
	
	public String getFileAs() {
		
		return getStringProperty("FileAs");
	}
	
	public void setFileAs(String val) {
		
		setProperty("FileAs", val);
	}
	
	public String getFirstName() {
		
		return getStringProperty("FirstName");
	}
	
	public void setFirstName(String name) {
		
		setProperty("FirstName", name);
	}
	
	public MailItem forwardAsBusinessCard() {
		
		return new MailItem((IDispatch) invoke("ForwardAsBusinessCard").getValue());
	}
	
	public MailItem forwardAsVcard() {
		
		return new MailItem((IDispatch) invoke("ForwardAsVcard").getValue());
	}
	
	public String getFTPSite() {
		
		return getStringProperty("FTPSite");
	}
	
	public void setFTPSite(String site) {
		
		setProperty("FTPSite", site);
	}
	
	public String getFullName() {
		
		return getStringProperty("FullName");
	}
	
	public void setFullName(String name) {
		
		setProperty("FullName", name);
	}
	
	public String getFullNameAndCompany() {
		
		return getStringProperty("FullNameAndCompany");
	}
	
	public Gender getGender() {
		
		return Gender.parse(getShortProperty("Gender"));
	}
	
	public void setGender(Gender sex) {
		
		setProperty("Gender", sex.value());
	}
	
	public String getGovernmentIDNumber() {
		
		return getStringProperty("GovernmentIDNumber");
	}
	
	public void setGovernmentIDNumber(String id) {
		
		setProperty("GovernmentIDNumber", id);
	}
	
	public boolean hasPicture() {
		
		return getBooleanProperty("HasPicture");
	}
	
	public String getHobby() {
		
		return getStringProperty("Hobby");
	}
	
	public void setHobby(String hobby) {
		
		setProperty("Hobby", hobby);
	}
	
	public String getHomeAddress() {
		
		return getStringProperty("HomeAddress");
	}
	
	public void setHomeAddress(String addr) {
		
		setProperty("HomeAddress", addr);
	}
	
	public String getHomeAddressCity() {
		
		return getStringProperty("HomeAddressCity");
	}
	
	public void setHomeAddressCity(String city) {
		
		setProperty("HomeAddressCity", city);
	}
	
	public String getHomeAddressCountry() {
		
		return getStringProperty("HomeAddressCountry");
	}
	
	public void setHomeAddressCountry(String country) {
		
		setProperty("HomeAddressCountry", country);
	}
	
	public String getHomeAddressPostalCode() {
		
		return getStringProperty("HomeAddressPostalCode");
	}
	
	public void setHomeAddressPostalCode(String code) {
		
		setProperty("HomeAddressPostalCode", code);
	}
	
	public String getHomeAddressPostOfficeBox() {
		
		return getStringProperty("HomeAddressPostOfficeBox");
	}
	
	public void setHomeAddressPostOfficeBox(String code) {
		
		setProperty("HomeAddressPostOfficeBox", code);
	}
	
	public String getHomeAddressState() {
		
		return getStringProperty("HomeAddressState");
	}
	
	public void setHomeAddressState(String state) {
		
		setProperty("HomeAddressState", state);
	}
	
	public String getHomeAddressStreet() {
		
		return getStringProperty("HomeAddressStreet");
	}
	
	public void setHomeAddressStreet(String street) {
		
		setProperty("HomeAddressStreet", street);
	}
	
	public String getHomeFaxNumber() {
		
		return getStringProperty("HomeFaxNumber");
	}
	
	public void setHomeFaxNumber(String num) {
		
		setProperty("HomeFaxNumber", num);
	}
	
	public String getHomeTelephoneNumber() {
		
		return getStringProperty("HomeTelephoneNumber");
	}
	
	public void setHomeTelephoneNumber(String num) {
		
		setProperty("HomeTelephoneNumber", num);
	}
	
	public String getIMAddress() {
		
		return getStringProperty("IMAddress");
	}
	
	public void setIMAddress(String addr) {
		
		setProperty("IMAddress", addr);
	}
	
	public String getInitials() {
		
		return getStringProperty("Initials");
	}
	
	public void setInitials(String initials) {
		
		setProperty("Initials", initials);
	}
	
	public String getInternetFreeBusyAddress() {
		
		return getStringProperty("InternetFreeBusyAddress");
	}
	
	public void setInternetFreeBusyAddress(String addr) {
		
		setProperty("InternetFreeBusyAddress", addr);
	}
	
	public String getISDNNumber() {
		
		return getStringProperty("ISDNNumber");
	}
	
	public void setISDNNumber(String num) {
		
		setProperty("ISDNNumber", num);
	}
	
	public String getJobTitle() {
		
		return getStringProperty("JobTitle");
	}
	
	public void setJobTitle(String title) {
		
		setProperty("JobTitle", title);
	}
	
	public boolean hasJournal() {
		
		return getBooleanProperty("Journal");
	}
	
	public void setJournal(boolean flag) {
		
		setProperty("Journal", flag);
	}
	
	public String getLanguage() {
		
		return getStringProperty("Language");
	}
	
	public void setLanguage(String language) {
		
		setProperty("Language", language);
	}
	
	public String getLastFirstAndSuffix() {
		
		return getStringProperty("LastFirstAndSuffix");
	}
	
	public String getLastFirstNoSpace() {
		
		return getStringProperty("LastFirstNoSpace");
	}
	
	public String getLastFirstNoSpaceAndSuffix() {
		
		return getStringProperty("LastFirstNoSpaceAndSuffix");
	}
	
	public String getLastFirstNoSpaceCompany() {
		
		return getStringProperty("LastFirstNoSpaceCompany");
	}
	
	public String getLastFirstSpaceOnly() {
		
		return getStringProperty("LastFirstSpaceOnly");
	}
	
	public String getLastFirstSpaceOnlyCompany() {
		
		return getStringProperty("LastFirstSpaceOnlyCompany");
	}
	
	public String getLastName() {
		
		return getStringProperty("LastName");
	}
	
	public void setLastName(String name) {
		
		setProperty("LastName", name);
	}
	
	public String getLastNameAndFirstName() {
		
		return getStringProperty("LastNameAndFirstName");
	}
	
	public String getMailingAddress() {
		
		return getStringProperty("MailingAddress");
	}
	
	public void setMailingAddress(String addr) {
		
		setProperty("MailingAddress", addr);
	}
	
	public String getMailingAddressCity() {
		
		return getStringProperty("MailingAddressCity");
	}
	
	public void setMailingAddressCity(String city) {
		
		setProperty("MailingAddressCity", city);
	}
	
	public String getMailingAddressCountry() {
		
		return getStringProperty("MailingAddressCountry");
	}
	
	public void setMailingAddressCountry(String country) {
		
		setProperty("MailingAddressCountry", country);
	}
	
	public String getMailingAddressPostalCode() {
		
		return getStringProperty("MailingAddressPostalCode");
	}
	
	public void setMailingAddressPostalCode(String code) {
		
		setProperty("MailingAddressPostalCode", code);
	}
	
	public String getMailingAddressPostOfficeBox() {
		
		return getStringProperty("MailingAddressPostOfficeBox");
	}
	
	public void setMailingAddressPostOfficeBox(String box) {
		
		setProperty("MailingAddressPostOfficeBox", box);
	}
	
	public String getMailingAddressState() {
		
		return getStringProperty("MailingAddressState");
	}
	
	public void setMailingAddressState(String state) {
		
		setProperty("MailingAddressState", state);
	}
	
	public String getMailingAddressStreet() {
		
		return getStringProperty("MailingAddressStreet");
	}
	
	public void setMailingAddressStreet(String street) {
		
		setProperty("MailingAddressStreet", street);
	}
	
	public String getManagerName() {
		
		return getStringProperty("ManagerName");
	}
	
	public void setManagerName(String name) {
		
		setProperty("ManagerName", name);
	}
	
	public String getMiddleName() {
		
		return getStringProperty("MiddleName");
	}
	
	public void setMiddleName(String name) {
		
		setProperty("MiddleName", name);
	}
	
	public String getMobileTelephoneNumber() {
		
		return getStringProperty("MobileTelephoneNumber");
	}
	
	public void setMobileTelephoneNumber(String num) {
		
		setProperty("MobileTelephoneNumber", num);
	}
	
	public String getNetMeetingAlias() {
		
		return getStringProperty("NetMeetingAlias");
	}
	
	public void setNetMeetingAlias(String alias) {
		
		setProperty("NetMeetingAlias", alias);
	}
	
	public String getNetMeetingServer() {
		
		return getStringProperty("NetMeetingServer");
	}
	
	public void setNetMeetingServer(String server) {
		
		setProperty("NetMeetingServer", server);
	}
	
	public String getNickName() {
		
		return getStringProperty("NickName");
	}
	
	public void setNickName(String name) {
		
		setProperty("NickName", name);
	}
	
	public String getOfficeLocation() {
		
		return getStringProperty("OfficeLocation");
	}
	
	public void setOfficeLocation(String location) {
		
		setProperty("OfficeLocation", location);
	}
	
	public String getOrganizationalIDNumber() {
		
		return getStringProperty("OrganizationalIDNumber");
	}
	
	public void setOrganizationalIDNumber(String id) {
		
		setProperty("OrganizationalIDNumber", id);
	}
	
	public String getOtherAddress() {
		
		return getStringProperty("OtherAddress");
	}
	
	public void setOtherAddress(String addr) {
		
		setProperty("OtherAddress", addr);
	}
	
	public String getOtherAddressCity() {
		
		return getStringProperty("OtherAddressCity");
	}
	
	public void setOtherAddressCity(String city) {
		
		setProperty("OtherAddressCity", city);
	}
	
	public String getOtherAddressCountry() {
		
		return getStringProperty("OtherAddressCountry");
	}
	
	public void setOtherAddressCountry(String country) {
		
		setProperty("OtherAddressCountry", country);
	}
	
	public String getOtherAddressPostalCode() {
		
		return getStringProperty("OtherAddressPostalCode");
	}
	
	public void setOtherAddressPostalCode(String code) {
		
		setProperty("OtherAddressPostalCode", code);
	}
	
	public String getOtherAddressPostOfficeBox() {
		
		return getStringProperty("OtherAddressPostOfficeBox");
	}
	
	public void setOtherAddressPostOfficeBox(String box) {
		
		setProperty("OtherAddressPostOfficeBox", box);
	}
	
	public String getOtherAddressState() {
		
		return getStringProperty("OtherAddressState");
	}
	
	public void setOtherAddressState(String state) {
		
		setProperty("OtherAddressState", state);
	}
	
	public String getOtherAddressStreet() {
		
		return getStringProperty("OtherAddressStreet");
	}
	
	public void setOtherAddressStreet(String street) {
		
		setProperty("OtherAddressStreet", street);
	}
	
	public String getOtherFaxNumber() {
		
		return getStringProperty("OtherFaxNumber");
	}
	
	public void setOtherFaxNumber(String num) {
		
		setProperty("OtherFaxNumber", num);
	}
	
	public String getOtherTelephoneNumber() {
		
		return getStringProperty("OtherTelephoneNumber");
	}
	
	public void setOtherTelephoneNumber(String num) {
		
		setProperty("OtherTelephoneNumber", num);
	}
	
	public String getPagerNumber() {
		
		return getStringProperty("PagerNumber");
	}
	
	public void setPagerNumber(String num) {
		
		setProperty("PagerNumber", num);
	}
	
	public String getPersonalHomePage() {
		
		return getStringProperty("PersonalHomePage");
	}
	
	public void setPersonalHomePage(String url) {
		
		setProperty("PersonalHomePage", url);
	}
	
	public String getPrimaryTelephoneNumber() {
		
		return getStringProperty("PrimaryTelephoneNumber");
	}
	
	public void setPrimaryTelephoneNumber(String num) {
		
		setProperty("PrimaryTelephoneNumber", num);
	}
	
	public String getProfession() {
		
		return getStringProperty("Profession");
	}
	
	public void setProfession(String prof) {
		
		setProperty("Profession", prof);
	}
	
	public String getRadioTelephoneNumber() {
		
		return getStringProperty("RadioTelephoneNumber");
	}
	
	public void setRadioTelephoneNumber(String num) {
		
		setProperty("RadioTelephoneNumber", num);
	}
	
	public String getReferredBy() {
		
		return getStringProperty("ReferredBy");
	}
	
	public void setReferredBy(String name) {
		
		setProperty("ReferredBy", name);
	}
	
	public void removePicture() {
		
		invokeNoReply("RemovePicture");
	}
	
	public void resetBusinessCard() {
		
		invokeNoReply("ResetBusinessCard");
	}
	
	public void saveBusinessCardImage(String path) {
		
		invokeNoReply("SaveBusinessCardImage", newVariant(path));
	}
	
	public MailingAddressType getSelectedMailingAddress() {
		
		return MailingAddressType.parse(getShortProperty("SelectedMailingAddress"));
	}
	
	public void setSelectedMailingAddress(MailingAddressType addrType) {
		
		setProperty("SelectedMailingAddress", addrType.value());
	}
	
	public void showBusinessCardEditor() {
		
		invokeNoReply("ShowBusinessCardEditor");
	}
	
	public void showCheckPhoneDialog(ContactPhoneNumberType num) {
		
		invokeNoReply("ShowCheckPhoneDialog", newVariant(num.value()));
	}
	
	public String getSpouse() {
		
		return getStringProperty("Spouse");
	}
	
	public void setSpouse(String name) {
		
		setProperty("Spouse", name);
	}
	
	public String getSuffix() {
		
		return getStringProperty("Suffix");
	}
	
	public void setSuffix(String val) {
		
		setProperty("Suffix", val);
	}
	
	public String getTelexNumber() {
		
		return getStringProperty("TelexNumber");
	}
	
	public void setTelexNumber(String num) {
		
		setProperty("TelexNumber", num);
	}
	
	public String getTitle() {
		
		return getStringProperty("Title");
	}
	
	public void setTitle(String ttl) {
		
		setProperty("Title", ttl);
	}
	
	public String getUser1() {
		
		return getStringProperty("User1");
	}
	
	public void setUser1(String usr) {
		
		setProperty("User1", usr);
	}
	
	public String getUser2() {
		
		return getStringProperty("User2");
	}
	
	public void setUser2(String usr) {
		
		setProperty("User2", usr);
	}
	
	public String getUser3() {
		
		return getStringProperty("User3");
	}
	
	public void setUser3(String usr) {
		
		setProperty("User3", usr);
	}
	
	public String getUser4() {
		
		return getStringProperty("User4");
	}
	
	public void setUser4(String usr) {
		
		setProperty("User4", usr);
	}
	
	public String getWebPage() {
		
		return getStringProperty("WebPage");
	}
	
	public void setWebPage(String url) {
		
		setProperty("WebPage", url);
	}
	
	public String getYomiCompanyName() {
		
		return getStringProperty("YomiCompanyName");
	}
	
	public void setYomiCompanyName(String val) {
		
		setProperty("YomiCompanyName", val);
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
