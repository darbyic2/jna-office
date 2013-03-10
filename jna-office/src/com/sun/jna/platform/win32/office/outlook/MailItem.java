package com.sun.jna.platform.win32.office.outlook;

import com.sun.jna.platform.win32.COM.IDispatch;

public class MailItem extends BaseItemLevel4 {

	MailItem(IDispatch iDispatch) {
		super(iDispatch);
	}
	
	public MailItem forward() {
		
		return new MailItem((IDispatch) invoke("Forward").getValue());
	}
	
	public MailItem copy() {
		
		return new MailItem((IDispatch) invoke("Copy").getValue());
	}
	
	public AddressEntry getSender() {
		
		return new AddressEntry(getAutomationProperty("Sender"));
	}
	
	public String getVotingOptions() {
		
		return getStringProperty("VotingOptions");
	}
	
	public String getVotingResponse() {
		
		return getStringProperty("VotingResponse");
	}
	
	public MailItem reply() {
		
		return new MailItem((IDispatch) invoke("Reply").getValue());
	}
	
	public MailItem replyAll() {
		
		return new MailItem((IDispatch) invoke("ReplyAll").getValue());
	}
	
	public void setSender(AddressEntry addr) {
		
		setProperty("Sender", addr.getIDispatch());
	}
	
	public void setVotingOptions(String optionsList) {
		
		setProperty("VotingOptions", optionsList);
	}
	
	public void setVotingResponse(String response) {
		
		setProperty("VotingResponse", response);
	}
	
}
