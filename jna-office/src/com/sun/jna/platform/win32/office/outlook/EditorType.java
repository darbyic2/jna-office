package com.sun.jna.platform.win32.office.outlook;

public class EditorType extends AbstractEnum {

	public final static EditorType	olEditorText	=new EditorType(1, "olEditorText"); //Text editor
	public final static EditorType	olEditorHTML	=new EditorType(2, "olEditorHTML"); //HTML editor
	public final static EditorType	olEditorRTF	=new EditorType(3, "olEditorRTF"); //Real Text Format (RTF) editor
	public final static EditorType	olEditorWord	=new EditorType(4, "olEditorWord"); //Microsoft Office Word editor

	private EditorType(int val, String name) {
		super((short) val, name);
	}
	
	public static EditorType parse(short typ) {
		
		switch(typ) {
		
		case 1:
			return olEditorText;
			
		case 2:
			return olEditorHTML;
			
		case 3:
			return olEditorRTF;
			
		case 4:
			return olEditorWord;
			
		default:
			return olEditorText;
		}
	}
}
