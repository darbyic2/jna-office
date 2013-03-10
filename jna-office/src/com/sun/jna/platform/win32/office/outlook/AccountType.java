package com.sun.jna.platform.win32.office.outlook;

public class AccountType extends AbstractEnum {
	
	public final static AccountType	olExchange		= new AccountType(0, "olExchange");		//An Exchange account.
	public final static AccountType	olImap			= new AccountType(1, "olImap");			//An IMAP account.
	public final static AccountType	olPop3			= new AccountType(2, "olPop3");			//A POP3 account.
	public final static AccountType	olHttp			= new AccountType(3, "olHttp");			//An HTTP account.
	public final static AccountType	olOtherAccount	= new AccountType(5, "olOtherAccount");	//Other or unknown account.

	private AccountType(int typ, String name) {
		super((short) typ, name);
	}

	public static AccountType parse(short typ) {
		
		switch(typ) {
		
		case 0:
			return olExchange;
			
		case 1:
			return olImap;
			
		case 2:
			return olPop3;
			
		case 3:
			return olHttp;
			
		case 5:
			return olOtherAccount;
			
		default:
			throw new RuntimeException("AccountType Enum: " + typ + " not recognised.");
		}
	}
}
