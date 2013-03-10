package com.sun.jna.platform.win32.office.outlook;

public class SaveAsType extends AbstractEnum {
	
	public final static SaveAsType	olTXT			= new SaveAsType(0, "olTXT");			//Text format (.txt)
	public final static SaveAsType	olRTF			= new SaveAsType(1, "olRTF");			//Rich Text format (.rtf)
	public final static SaveAsType	olTemplate		= new SaveAsType(2, "olTemplate");		//Microsoft Outlook template (.oft)
	public final static SaveAsType	olMSG			= new SaveAsType(3, "olMSG");			//Outlook message format (.msg)
	public final static SaveAsType	olDoc			= new SaveAsType(4, "olDoc");			//Microsoft Office Word format (.doc)
	public final static SaveAsType	olHTML			= new SaveAsType(5, "olHTML");			//HTML format (.html)
	public final static SaveAsType	olVCard			= new SaveAsType(6, "olVCard");			//VCard format (.vcf)
	public final static SaveAsType	olVCal			= new SaveAsType(7, "olVCal");			//VCal format (.vcs)
	public final static SaveAsType	olICal			= new SaveAsType(8, "olICal");			//iCal format (.ics)
	public final static SaveAsType	olMSGUnicode	= new SaveAsType(9, "olMSGUnicode");	//Outlook Unicode message format (.msg)
	public final static SaveAsType	olMHTML			= new SaveAsType(10, "olMHTML");		//MIME HTML format (.mht)

	private SaveAsType(int typ, String name) {
		super((short) typ, name);
	}
	
	public static SaveAsType parse(short typ) {
		
		switch(typ) {
		case 0:
			return olTXT;
			
		case 1:
			return olRTF;
			
		case 2:
			return olTemplate;
			
		case 3:
			return olMSG;
			
		case 4:
			return olDoc;
			
		case 5:
			return olHTML;
			
		case 6:
			return olVCard;
			
		case 7:
			return olVCal;
			
		case 8:
			return olICal;
			
		case 9:
			return olMSGUnicode;
			
		case 10:
			return olMHTML;
			
		default:
			return olTXT;
		}
	}
}
