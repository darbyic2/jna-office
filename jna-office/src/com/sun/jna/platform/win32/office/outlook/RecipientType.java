package com.sun.jna.platform.win32.office.outlook;

/**
 * Convenience class to tie hierarchy of Mail/Meeting/Journal/Task Recipient types.
 * has no implementation beyond a constructor.
 * 
 * @author 801767553
 *
 */
public abstract class RecipientType extends AbstractEnum {

	protected RecipientType(int typ, String name) {
		super((short) typ, name);
	}
	
}
