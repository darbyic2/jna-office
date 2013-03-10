package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.Variant;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant.VARIANT;

public class Items extends BaseOutlookObject {

	public Items(IDispatch iDisp) {
		super(iDisp);
	}
	
	private Object wrapItem(VARIANT var) {
		IDispatch iDisp;
		
		if (var == null || var.getVarType().intValue() == Variant.VT_EMPTY || var.getVarType().intValue() == Variant.VT_NULL || var.getVarType().intValue() != Variant.VT_DISPATCH) {
			return null;
			
		} else {
			iDisp = (IDispatch) var.getValue();
			
			switch((new BaseOutlookObject(iDisp)).getClassEnumValue()) {
			
//				case ClassEnum.olAppointment:
//					return new AppointmentItem(auto);
//				
//				case ClassEnum.olContact:
//					return new ContactItem(auto);
//					
//				case ClassEnum.olJournal:
//					return new JournalItem(auto);
					
				case ClassEnum.olMail:
					return new MailItem(iDisp);
					
//				case ClassEnum.olNote:
//					return new NoteItem(auto);
//					
//				case ClassEnum.olPost:
//					return new PostItem(auto);
//					
//				case ClassEnum.olTask:
//					return new TaskItem(auto);
				
				default:
					return null;
			}
		}
	}
	
	public BaseItemLevel1 add(String name) {
		
		return (BaseItemLevel1) wrapItem(invoke("Add", newVariant(name)));
	}
	
	public BaseItemLevel1 add(String name, FolderType typ) {
		
		return (BaseItemLevel1) wrapItem(invoke("Add", newVariant(name), newVariant(typ.value())));
	}
	
	public int count() {
		
		return getIntProperty("Count");
	}
	
	public BaseItemLevel1 find(String filter) {
		
		return (BaseItemLevel1) wrapItem(invoke("Find", newVariant(filter)));
	}
	
	private BaseItemLevel1 getHelper(String command) {
		
		return (BaseItemLevel1) wrapItem(invoke(command));
	}
	
	public BaseItemLevel1 getFirst() {
		
		return getHelper("GetFirst");
	}
	
	public BaseItemLevel1 getLast() {
		
		return getHelper("GetLast");
	}
	
	public BaseItemLevel1 getNext() {
		
		return getHelper("GetNext");
	}
	
	public BaseItemLevel1 getPrevious() {
		
		return getHelper("GetPrevious");
	}
	
	public boolean getIncludeRecurrences() {
		
		return getBooleanProperty("IncludeRecurrences");
	}
	
	public void setIncludeRecurrences(boolean flag) {
		
		setProperty("IncludeRecurrences", flag);
	}
	
	public BaseItemLevel1 get(int index) {
		
		return (BaseItemLevel1) wrapItem(invoke("Item", newVariant(index)));
	}
	
	public void remove(int index) {
		
		invokeNoReply("Display", newVariant(index));
	}
	
	public void resetColumns() {
		
		invokeNoReply("ResetColumns");
	}
	
	public Items restrict(String filter) {
		
		return new Items((IDispatch) invoke("Restrict", newVariant(filter)).getValue());
	}
	
	public void setColumns(String columnNames) {
		
		invokeNoReply("SetColumns", newVariant(columnNames));
	}
	
	public void sort(String propertyName, boolean useDescendingOrder) {
		
		invokeNoReply("Sort", newVariant(propertyName), newVariant(useDescendingOrder));
	}
	
	public void sort(String propertyName) {
		
		sort(propertyName, false);
	}
	
}
