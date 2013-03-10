package com.sun.jna.platform.win32.office.outlook;

public class ActionReplyStyle extends AbstractEnum {
	
	public final static ActionReplyStyle	olOmitOriginalText		= new ActionReplyStyle(	0	, "olOmitOriginalText");  //	The reply will not include any references to the original item or its text.
	public final static ActionReplyStyle	olEmbedOriginalItem		= new ActionReplyStyle(	1	, "olEmbedOriginalItem");  //	The reply will include the original item embedded in it.
	public final static ActionReplyStyle	olIncludeOriginalText	= new ActionReplyStyle(	2	, "olIncludeOriginalText");  //	The reply will include the text of the original item.
	public final static ActionReplyStyle	olIndentOriginalText	= new ActionReplyStyle(	3	, "olIndentOriginalText");  //	The reply will include the indented text of the original item.
	public final static ActionReplyStyle	olLinkOriginalItem		= new ActionReplyStyle(	4	, "olLinkOriginalItem");  //	The reply will include a link to the original item.
	public final static ActionReplyStyle	olUserPreference		= new ActionReplyStyle(	5	, "olUserPreference");  //	The reply style will be set based on the user's preference.
	public final static ActionReplyStyle	olReplyTickOriginalText	= new ActionReplyStyle(	1000, "olReplyTickOriginalText");  //	The reply will include the original text with each line preceded by a symbol such as ">".
	
	private ActionReplyStyle(int val, String name) {
		super((short) val, name);
	}

	public static ActionReplyStyle parse(short style) {
		switch(style) {
		
		case 0:
			return olOmitOriginalText;
			
		case 1:
			return olEmbedOriginalItem;
			
		case 2:
			return olIncludeOriginalText;
			
		case 3:
			return olIndentOriginalText;
			
		case 4:
			return olLinkOriginalItem;
			
		case 5:
			return olUserPreference;
			
		case 1000:
			return olReplyTickOriginalText;
			
		default:
			throw new RuntimeException("ActionReplyStyle Enum: " + style + " not recognised.");
		}
	}
}
