package com.sun.jna.platform.win32.wsh;

import com.sun.jna.platform.win32.OleAuto;
import com.sun.jna.platform.win32.COM.COMObject;
import com.sun.jna.platform.win32.COM.IDispatch;
import com.sun.jna.platform.win32.Variant.VARIANT;

public class WshShell extends COMObject {

    public WshShell() {
        super("WScript.Shell", false);
    }
    
    private static VARIANT newVariant(String value) {
        
        return new VARIANT(OleAuto.INSTANCE.SysAllocString(value));
    }
    
    public WshShortcut createShortcut(String shortcutPathName) {
        
        VARIANT.ByReference result = new VARIANT.ByReference();
        this.oleMethod(OleAuto.DISPATCH_METHOD, result, this.iDispatch,
                    "CreateShortcut", newVariant(shortcutPathName));
        
        return new WshShortcut((IDispatch) result.getValue());
    }
    
    public String getSpecialFolder(String specialFolderName) {
        
        VARIANT.ByReference result = new VARIANT.ByReference();
        this.oleMethod(OleAuto.DISPATCH_METHOD, result, this.iDispatch,
                    "SpecialFolders", newVariant(specialFolderName));
        
        return result.getValue().toString();
    }
    
    public class WshShortcut extends COMObject {
        
        WshShortcut(IDispatch iDisp) {
            super(iDisp);
        }
        
        public void save() {
            
            VARIANT.ByReference result = new VARIANT.ByReference();
            this.oleMethod(OleAuto.DISPATCH_METHOD, result, this.iDispatch,
                         "Save");
        }
        
        public void setTargetPath(String targetPath) {
            
            VARIANT.ByReference result = new VARIANT.ByReference();
            this.oleMethod(OleAuto.DISPATCH_PROPERTYPUT, result, this.iDispatch,
                        "TargetPath", newVariant(targetPath));
        }
    }
    
    public static void main(String[] args) {
        
        WshShell shell = new WshShell();
        WshShortcut link = shell.createShortcut(shell.getSpecialFolder("Desktop") + "/LinkToAutoexec.lnk");
        link.setTargetPath("c:/autoexec.bat");
        link.save();
    }
}
