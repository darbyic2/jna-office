package com.sun.jna.platform.win32.office.outlook;

public class AutoDiscoverConnectionMode extends AbstractEnum {
	
	public final static AutoDiscoverConnectionMode olAutoDiscoverConnectionUnknown		  = new AutoDiscoverConnectionMode(0, "olAutoDiscoverConnectionUnknown"); //	Other or unknown connection, or no connection.
	public final static AutoDiscoverConnectionMode olAutoDiscoverConnectionExternal		  = new AutoDiscoverConnectionMode(1, "olAutoDiscoverConnectionExternal"); //	Connection is over the Internet.
	public final static AutoDiscoverConnectionMode olAutoDiscoverConnectionInternal		  = new AutoDiscoverConnectionMode(2, "olAutoDiscoverConnectionInternal"); //	Connection is over the Intranet.
	public final static AutoDiscoverConnectionMode olAutoDiscoverConnectionInternalDomain = new AutoDiscoverConnectionMode(3, "olAutoDiscoverConnectionInternalDomain"); //	Connection is in the same domain over the Intranet.

	private AutoDiscoverConnectionMode(int mode, String name) {
		super((short) mode, name);
	}

	public static AutoDiscoverConnectionMode parse(short mode) {
		
		switch(mode) {
		
		case 0:
			return olAutoDiscoverConnectionUnknown;
			
		case 1:
			return olAutoDiscoverConnectionExternal;
			
		case 2:
			return olAutoDiscoverConnectionInternal;
			
		case 3:
			return olAutoDiscoverConnectionInternalDomain;
			
		default:
			return olAutoDiscoverConnectionUnknown;
		}
	}
}
