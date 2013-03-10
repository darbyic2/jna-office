package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.Variant;
import com.sun.jna.platform.win32.COM.COMException;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant.VARIANT;

public class Outlook  extends BaseOutlookObject {

	public Outlook() {
		super("Outlook.Application", true);
	}
	
	Outlook(IDispatch iDispatch) {
		super(iDispatch);
	}
	
	public Outlook getApplication() {
		
		return new Outlook(getAutomationProperty("Application"));
	}
	
	public AppointmentItem createAppointmentItem() {
		
		VARIANT result = invoke("CreateItem", newVariant(ItemType.APPOINTMENT_ITEM.value()));
		
		if (result == null || result.getVarType().intValue() == Variant.VT_EMPTY || result.getVarType().intValue() != Variant.VT_DISPATCH)
			return null;
		else
			return new AppointmentItem((IDispatch) result.getValue());
	}
	
	public ContactItem createContactItem() {
		
		VARIANT result = invoke("CreateItem", newVariant(ItemType.CONTACT_ITEM.value()));
		
		if (result == null || result.getVarType().intValue() == Variant.VT_EMPTY || result.getVarType().intValue() != Variant.VT_DISPATCH)
			return null;
		else
			return new ContactItem((IDispatch) result.getValue());
	}
	
	public DistributionListItem createDistributionListItem() {
		
		VARIANT result = invoke("CreateItem", newVariant(ItemType.DISTRIBUTION_LIST_ITEM.value()));
		
		if (result == null || result.getVarType().intValue() == Variant.VT_EMPTY || result.getVarType().intValue() != Variant.VT_DISPATCH)
			return null;
		else
			return new DistributionListItem((IDispatch) result.getValue());
	}
	
	public JournalItem createJournalItem() {
		
		VARIANT result = invoke("CreateItem", newVariant(ItemType.JOURNAL_ITEM.value()));
		
		if (result == null || result.getVarType().intValue() == Variant.VT_EMPTY || result.getVarType().intValue() != Variant.VT_DISPATCH)
			return null;
		else
			return new JournalItem((IDispatch) result.getValue());
	}
	
	public MailItem createMailItem() {
		
		VARIANT result = invoke("CreateItem", newVariant(ItemType.MAIL_ITEM.value()));
		
		if (result == null || result.getVarType().intValue() == Variant.VT_EMPTY || result.getVarType().intValue() != Variant.VT_DISPATCH)
			return null;
		else
			return new MailItem((IDispatch) result.getValue());
	}
	
	public MobileItem createMobileMMSItem() {
		
		VARIANT result = invoke("CreateItem", newVariant(ItemType.MOBILE_ITEM_MMS.value()));
		
		if (result == null || result.getVarType().intValue() == Variant.VT_EMPTY || result.getVarType().intValue() != Variant.VT_DISPATCH)
			return null;
		else
			return new MobileItem((IDispatch) result.getValue());
	}
	
	public MobileItem createMobileSMSItem() {
		
		VARIANT result = invoke("CreateItem", newVariant(ItemType.MOBILE_ITEM_SMS.value()));
		
		if (result == null || result.getVarType().intValue() == Variant.VT_EMPTY || result.getVarType().intValue() != Variant.VT_DISPATCH)
			return null;
		else
			return new MobileItem((IDispatch) result.getValue());
	}
	
	public NoteItem createNoteItem() {
		
		VARIANT result = invoke("CreateItem", newVariant(ItemType.NOTE_ITEM.value()));
		
		if (result == null || result.getVarType().intValue() == Variant.VT_EMPTY || result.getVarType().intValue() != Variant.VT_DISPATCH)
			return null;
		else
			return new NoteItem((IDispatch) result.getValue());
	}
	
	public PostItem createPostItem() {
		
		VARIANT result = invoke("CreateItem", newVariant(ItemType.POST_ITEM.value()));
		
		if (result == null || result.getVarType().intValue() == Variant.VT_EMPTY || result.getVarType().intValue() != Variant.VT_DISPATCH)
			return null;
		else
			return new PostItem((IDispatch) result.getValue());
	}
	
	public TaskItem createTaskItem() {
		
		VARIANT result = invoke("CreateItem", newVariant(ItemType.TASK_ITEM.value()));
		
		if (result == null || result.getVarType().intValue() == Variant.VT_EMPTY || result.getVarType().intValue() != Variant.VT_DISPATCH)
			return null;
		else
			return new TaskItem((IDispatch) result.getValue());
	}
	
	public String getDefaultProfileName() {
		
		return getStringProperty("DefaultProfileName");
	}
	
	public String getName() {
		
		return getStringProperty("Name");
	}
	
	public Namespace getNamespace(String name) {
		
		return new Namespace((IDispatch) invoke("GetNameSpace", newVariant(name)).getValue());
	}
	
	public String getProductCode() {
		
		return getStringProperty("ProductCode");
	}
	
	public void quit() {
		invokeNoReply("Quit");
	}
	
	public boolean isTrusted() {
		
		return getBooleanProperty("IsTrusted");
	}
	
	public String getVersion() throws COMException {
		
		return getStringProperty("Version");
	}

}
