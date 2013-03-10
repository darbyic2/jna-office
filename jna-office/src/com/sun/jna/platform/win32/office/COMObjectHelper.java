package com.sun.jna.platform.win32.office;

import java.util.Date;

import com.sun.jna.platform.win32.OleAuto;
import com.sun.jna.platform.win32.Variant;
import com.sun.jna.platform.win32.WinDef;
import com.sun.jna.platform.win32.COM.COMException;
import com.sun.jna.platform.win32.COM.COMObject;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.OaIdl.DATE;
import com.sun.jna.platform.win32.OaIdl.VARIANT_BOOL;
import com.sun.jna.platform.win32.Variant.VARIANT;
import com.sun.jna.platform.win32.WinDef.LONG;
import com.sun.jna.platform.win32.WinDef.SHORT;

public class COMObjectHelper extends COMObject {
	
	private final static long COM_DAYS_ADJUSTMENT = 25569L; 	//((1969 - 1899) * 365) +1 + Leap years = Days
	private final static long MS_PER_DAY = 86400000L;   		//24L * 60L * 60L * 1000L;

	protected static Date convertFromCOMDate(VARIANT comDate) {
		
		double doubleDate = ((DATE) comDate.getValue()).date;
		long longDate = (long) doubleDate;
		
		double doubleTime = doubleDate - ((double) longDate);
		long longTime = (long) (doubleTime * ((double) MS_PER_DAY));
		
		return new Date(((longDate  - COM_DAYS_ADJUSTMENT) * MS_PER_DAY) + longTime);
	}
	
	protected static VARIANT convertToCOMDate(Date javaDate) {
		long longTime = javaDate.getTime() % MS_PER_DAY;
		long longDate = ((javaDate.getTime() - longTime) / MS_PER_DAY) + COM_DAYS_ADJUSTMENT;
		
		float floatTime = ((float) longTime) / ((float) MS_PER_DAY);
		float floatDateTime = floatTime + ((float) longDate);
		return new VARIANT(new DATE(floatDateTime));
	}

	
	protected COMObjectHelper(IDispatch iDispatch) {
		super(iDispatch);
	}

	public COMObjectHelper(String progId, boolean useActiveInstance)
			throws COMException {
		super(progId, useActiveInstance);
	}
	
	protected IDispatch getAutomationProperty(String propertyName) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_PROPERTYGET, result, this.iDispatch,
				propertyName);

		return ((IDispatch) result.getValue());
	}

	protected boolean getBooleanProperty(String propertyName) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_PROPERTYGET, result, this.iDispatch,
				propertyName);

		return (((VARIANT_BOOL) result.getValue()).intValue() != 0);
	}
	
	protected Date getDateProperty(String propertyName) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_PROPERTYGET, result, this.iDispatch,
				propertyName);
		
		return convertFromCOMDate(result);
	}
	
	protected int getIntProperty(String propertyName) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_PROPERTYGET, result, this.iDispatch,
				propertyName);

		return ((LONG) result.getValue()).intValue();
	}
	
	protected short getShortProperty(String propertyName) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_PROPERTYGET, result, this.iDispatch,
				propertyName);

		return ((SHORT) result.getValue()).shortValue();
	}
	
	protected String getStringProperty(String propertyName) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_PROPERTYGET, result, this.iDispatch,
				propertyName);

		return result.getValue().toString();
	}
	
	protected VARIANT invoke(String methodName) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_METHOD, result, this.iDispatch,
				methodName);

		return result;
	}
	
	protected VARIANT invoke(String methodName, VARIANT arg) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_METHOD, result, this.iDispatch,
				methodName, arg);

		return result;
	}
	
	protected VARIANT invoke(String methodName, VARIANT[] args) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_METHOD, result, this.iDispatch,
				methodName, args);

		return result;
	}
	
	protected VARIANT invoke(String methodName, VARIANT arg1, VARIANT arg2) {
		
		return invoke(methodName, new VARIANT[] {arg1, arg2});
	}
	
	protected VARIANT invoke(String methodName, VARIANT arg1, VARIANT arg2, VARIANT arg3) {
		
		return invoke(methodName, new VARIANT[] {arg1, arg2, arg3});
	}
	
	protected VARIANT invoke(String methodName, VARIANT arg1, VARIANT arg2, VARIANT arg3, VARIANT arg4) {
		
		return invoke(methodName, new VARIANT[] {arg1, arg2, arg3, arg4});
	}
	
	protected void invokeNoReply(String methodName) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_METHOD, result, this.iDispatch, methodName);
	}
	
	protected void invokeNoReply(String methodName, VARIANT arg) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_METHOD, result, this.iDispatch, methodName, arg);
	}
	
	protected void invokeNoReply(String methodName, VARIANT[] args) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_METHOD, result, this.iDispatch, methodName, args);
	}
	
	protected void invokeNoReply(String methodName, VARIANT arg1, VARIANT arg2) {
		
		invokeNoReply(methodName, new VARIANT[] {arg1, arg2});
	}
	
	protected void invokeNoReply(String methodName, VARIANT arg1, VARIANT arg2, VARIANT arg3) {
		
		invokeNoReply(methodName, new VARIANT[] {arg1, arg2, arg3});
	}
	
	protected void invokeNoReply(String methodName, VARIANT arg1, VARIANT arg2, VARIANT arg3, VARIANT arg4) {
		
		invokeNoReply(methodName, new VARIANT[] {arg1, arg2, arg3, arg4});
	}
	
	protected static VARIANT newVariant(boolean value) {
		
		return new VARIANT(new VARIANT_BOOL((value? 1: 0)));
	}
	
	protected static VARIANT newVariant(IDispatch value) {
		
		VARIANT v = new VARIANT();
		v.setValue(Variant.VT_DISPATCH, value);
		return v;
	}
	
	protected static VARIANT newVariant(DATE value) {
		
		return new VARIANT(value);
	}
	
	protected static VARIANT newVariant(Date value) {
		
		return convertToCOMDate(value);
	}
	
	protected static VARIANT newVariant(int value) {
		
		return new VARIANT(new WinDef.LONG(value));
	}
	
	protected static VARIANT newVariant(short value) {
		
		return new VARIANT(new WinDef.SHORT(value));
	}
	
	protected static VARIANT newVariant(String value) {
		
		return new VARIANT(OleAuto.INSTANCE.SysAllocString(value));
	}
	
	protected void setProperty(String propertyName, boolean value) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_PROPERTYPUT, result, this.iDispatch,
				propertyName, newVariant(value));
	}
	
	protected void setProperty(String propertyName, Date value) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_PROPERTYPUT, result, this.iDispatch,
				propertyName, newVariant(value));
	}
	
	protected void setProperty(String propertyName, IDispatch value) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_PROPERTYPUT, result, this.iDispatch,
				propertyName, newVariant(value));
	}
	
	protected void setProperty(String propertyName, int value) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_PROPERTYPUT, result, this.iDispatch,
				propertyName, newVariant(value));
	}
	
	protected void setProperty(String propertyName, short value) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_PROPERTYPUT, result, this.iDispatch,
				propertyName, newVariant(value));
	}
	
	protected void setProperty(String propertyName, String value) {
		
		VARIANT.ByReference result = new VARIANT.ByReference();
		this.oleMethod(OleAuto.DISPATCH_PROPERTYPUT, result, this.iDispatch,
				propertyName, newVariant(value));
	}
	
	public VARIANT toVariant() {
		
		return newVariant(this.iDispatch);
	}
	
}
