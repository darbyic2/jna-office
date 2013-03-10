package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class DistributionListItem extends BaseItemLevel3 {
	
	public DistributionListItem(IDispatch iDisp) {
		super(iDisp);
	}

	public void addMember(Recipient rx) {
		
		invokeNoReply("Delete", newVariant(rx.getIDispatch()));
	}
	
	public void addMembers(Recipients list) {
		
		invokeNoReply("Delete", newVariant(list.getIDispatch()));
	}
	
	public String getDLName() {
		
		return getStringProperty("DLName");
	}
	
	public void setDLName(String name) {
		
		setProperty("DLName", name);
	}
	
	public Recipient getMember(int index) {
		
		return new Recipient((IDispatch) invoke("GetMember", newVariant(index)).getValue());
	}
	
	public int getMemberCount() {
		
		return getIntProperty("MemberCount");
	}
	
	public void removeMember(Recipient rx) {
		
		invokeNoReply("RemoveMember", newVariant(rx.getIDispatch()));
	}
	
	public void removeMembers(Recipients list) {
		
		invokeNoReply("RemoveMembers", newVariant(list.getIDispatch()));
	}
	
}
