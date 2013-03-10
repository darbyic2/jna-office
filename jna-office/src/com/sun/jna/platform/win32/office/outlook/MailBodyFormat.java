package com.sun.jna.platform.win32.office.outlook;

public class MailBodyFormat {

	public final static MailBodyFormat UNSPECIFIED = new MailBodyFormat(0, "UNSPECIFIED");
	public final static MailBodyFormat PLAIN_TEXT = new MailBodyFormat(1, "PLAIN_TEXT");
	public final static MailBodyFormat HTML_TEXT = new MailBodyFormat(2, "HTML_TEXT");
	public final static MailBodyFormat RICH_TEXT = new MailBodyFormat(3, "RICH_TEXT");
	
	private int value;
	private String name;
	
	private MailBodyFormat(int fmt, String name) {
		super();
		value = fmt;
		this.name = name;
	}
	
	public int value() {
		return value;
	}
	
	public static MailBodyFormat parse(int value) {
		switch(value) {
		case 1:
			return PLAIN_TEXT;
			
		case 2:
			return HTML_TEXT;
			
		case 3:
			return RICH_TEXT;
			
		default:
			return UNSPECIFIED;
		}
	}

	@Override
	public String toString() {
		return name;
	}
	
	
}
