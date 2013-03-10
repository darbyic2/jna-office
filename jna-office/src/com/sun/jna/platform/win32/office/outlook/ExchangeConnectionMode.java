package com.sun.jna.platform.win32.office.outlook;

public class ExchangeConnectionMode extends AbstractEnum {
	
	public final static ExchangeConnectionMode	olNoExchange	= new ExchangeConnectionMode(0, "olNoExchange"); //	The account does not use an Exchange server.
	public final static ExchangeConnectionMode	olOffline	= new ExchangeConnectionMode(100, "olOffline"); //	The account is not connected to an Exchange server and is in the classic offline mode. This also occurs when the user selects Work Offline from the File menu.
	public final static ExchangeConnectionMode	olCachedOffline	= new ExchangeConnectionMode(200, "olCachedOffline"); //	The account is using cached Exchange mode and the user has selected Work Offline from the File menu.
	public final static ExchangeConnectionMode	olDisconnected	= new ExchangeConnectionMode(300, "olDisconnected"); //	The account has a disconnected connection to the Exchange server.
	public final static ExchangeConnectionMode	olCachedDisconnected	= new ExchangeConnectionMode(400, "olCachedDisconnected"); //	The account is using cached Exchange mode with a disconnected connection to the Exchange server.
	public final static ExchangeConnectionMode	olCachedConnectedHeaders	= new ExchangeConnectionMode(500, "olCachedConnectedHeaders"); //	The account is using cached Exchange mode on a dial-up or slow connection with the Exchange server, such that only headers are downloaded. Full item bodies and attachments remain on the server. The user can also select this state manually regardless of connection speed.
	public final static ExchangeConnectionMode	olCachedConnectedDrizzle	= new ExchangeConnectionMode(600, "olCachedConnectedDrizzle"); //	The account is using cached Exchange mode such that headers are downloaded first, followed by the bodies and attachments of full items.
	public final static ExchangeConnectionMode	olCachedConnectedFull	= new ExchangeConnectionMode(700, "olCachedConnectedFull"); //	The account is using cached Exchange mode on a Local Area Network or a fast connection with the Exchange server. The user can also select this state manually, disabling auto-detect logic and always downloading full items regardless of connection speed.
	public final static ExchangeConnectionMode	olOnline	= new ExchangeConnectionMode(800, "olOnline"); //	The account is connected to an Exchange server and is in the classic online mode.

	private ExchangeConnectionMode(int mode, String name) {
		super((short) mode, name);
	}

	public static ExchangeConnectionMode parse(short mode) {
		
		switch(mode) {
		
		case 0:
			return olNoExchange;
			
		case 100:
			return olOffline;
			
		case 200:
			return olCachedOffline;
			
		case 300:
			return olDisconnected;
			
		case 400:
			return olCachedDisconnected;
			
		case 500:
			return olCachedConnectedHeaders;
			
		case 600:
			return olCachedConnectedDrizzle;
			
		case 700:
			return olCachedConnectedFull;
			
		case 800:
			return olOnline;
		
		default:
			return olNoExchange;
		}
	}
}
