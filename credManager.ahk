#Requires AutoHotkey v2.0

/*
===============================================================================
  credManager.ahk - Windows Credential Manager GUI for AutoHotkey v2
===============================================================================

Script Name: AHK CredManager
Author: Dixon Hill
Date: 2025-05-14
Version: 1.0
License: MIT
URL: https://github.com
Description: This script provides a graphical user interface (GUI) for securely managing credentials-such as usernames and passwords-using the Windows Credential Manager, all through AutoHotkey v2. It allows users to easily create, edit, delete, and view credentials via a convenient window that can be opened with a customizable hotkey. Credentials are stored with a specific prefix in Windows Credential Manager, ensuring they are kept secure and separate from other credentials. The script also exposes simple methods for other AutoHotkey scripts to programmatically retrieve or update credentials, making it a practical tool for both end users and developers who need to manage sensitive login information within their automation workflows.

USAGE INSTRUCTIONS:

1. Run this script with AutoHotkey v2.
2. Press SHIFT + CTRL + ALT + C to open the Credentials Manager window (you can change the hotkey in the script).

   - Click "Create" to add a new credential (name, username, password).
   - Right-click a credential in the list to Edit, Delete, or Copy its password.

3. To use credentials in your own scripts:
     myCred := CredManager["myCredName"]
     MsgBox(myCred.user)  ; Username
     MsgBox(myCred.pass)  ; Password

   To store/update credentials in code:
     CredManager["myCredName"] := {user: "MyUser", pass: "MyPass"}

NOTES:
- Credentials are stored in Windows Credential Manager with the prefix "Ahk_".
- The script creates only one main GUI window; it is refreshed after changes.
- The tray/taskbar icon can be customized by placing an .ico file in the script folder and using:
      Menu, Tray, Icon, myIcon.ico

===============================================================================
*/

global MainGui := false ; Stores the main GUI reference for refreshing

; Hotkey: Ctrl+Alt+C opens the main credentials GUI
^!+c:: CredManager.ShowGUI()

; Usage notes:
; To get a credential: myCred := CredManager["myCred"]
;   Then use myCred.user / myCred.pass or just CredManager["myCred"].pass
; To set a credential: CredManager["myCred"] := {user: "myUser", pass: "myPass"}

class CredManager {
    static prefix := "Ahk_"          ; Prefix for all credentials stored in Windows Credential Manager
    static filter := this.prefix "*" ; Filter for enumerating credentials

    ; Allows array-like access for getting/setting credentials
    static __Item[name] {
        get => this.Read(name)
        set {
            if Type(value) = "Object" && value.HasOwnProp("user") & value.HasOwnProp("pass")
                this.Write(name, value.user, value.pass)
        }
    }

    ; Adds prefix if missing
    static AddPrefixName(name) => (SubStr(name, 1, StrLen(this.prefix)) = this.prefix ? "" : this.prefix) . name

    ; Removes prefix if present
    static DelPrefixName(name) => SubStr(name, 1, StrLen(this.prefix)) = this.prefix ? SubStr(name, StrLen(this.prefix)+1) : name

    ; Writes a credential to Windows Credential Manager
    static Write(name, username, password) {
        name := this.AddPrefixName(name)
        cred := Buffer(24 + A_PtrSize * 7, 0)
        cbPassword := StrLen(password) * 2
        NumPut("UInt", 1, cred, 4 + A_PtrSize * 0)                ; Type = CRED_TYPE_GENERIC
        NumPut("Ptr", StrPtr(name), cred, 8 + A_PtrSize * 0)      ; TargetName
        NumPut("UInt", cbPassword, cred, 16 + A_PtrSize * 2)      ; CredentialBlobSize
        NumPut("Ptr", StrPtr(password), cred, 16 + A_PtrSize * 3) ; CredentialBlob
        NumPut("UInt", 3, cred, 16 + A_PtrSize * 4)               ; Persist = CRED_PERSIST_ENTERPRISE
        NumPut("Ptr", StrPtr(username), cred, 24 + A_PtrSize * 6) ; UserName
        return DllCall("Advapi32.dll\CredWriteW", "Ptr", cred, "UInt", 0, "UInt")
    }

    ; Deletes a credential from Windows Credential Manager
    static Delete(name) {
        return DllCall("Advapi32.dll\CredDeleteW",
            "WStr", this.AddPrefixName(name),
            "UInt", 1,
            "UInt", 0,
            "UInt")
    }

    ; Reads a credential from Windows Credential Manager
    static Read(name) {
        pCred := 0
        DllCall("Advapi32.dll\CredReadW",
            "Str", this.AddPrefixName(name),
            "UInt", 1,
            "UInt", 0,
            "Ptr*", &pCred,
            "UInt")
        if !pCred
            return { name: false, user: "", pass: "" }
        name := StrGet(NumGet(pCred, 8 + A_PtrSize * 0, "UPtr"), 256, "UTF-16")
        usrExist := NumGet(pCred, 24 + A_PtrSize * 6, "UPtr")
        username := usrExist = 0 ? "" : StrGet(usrExist, 256, "UTF-16")
        len := NumGet(pCred, 16 + A_PtrSize * 2, "UInt")
        password := StrGet(NumGet(pCred, 16 + A_PtrSize * 3, "UPtr"), len / 2, "UTF-16")
        DllCall("Advapi32.dll\CredFree", "Ptr", pCred)
        return { name: name, user: username, pass: password }
    }

    ; Returns the count of credentials
    static Count() => this.Enumerate().count

    ; Returns all credential objects
    static GetAllCreds() => this.Enumerate().creds

    ; Enumerates credentials using the filter
    static Enumerate(filter := this.filter) {
        flags := count := creds := 0
        DllCall("Advapi32.dll\CredEnumerateW", "Str", filter, "UInt", flags, "UInt*", &count, "Ptr*", &creds)
        SetTimer () => DllCall("Advapi32.dll\CredFree", "Ptr", creds), -2000
        return { count: count, creds: creds }
    }

    ; Returns a list of all credentials as objects
    static ListAll(filter := this.filter) {
        List := []
        creds := this.Enumerate(filter)
        loop creds.count {
            PtrCred := NumGet(creds.creds, (A_Index - 1) * A_PtrSize, "Int")
            if PtrCred != 0 {
                name := StrGet(NumGet(PtrCred, 8, "UPtr"), 256, "UTF-16")
                List.Push(this.Read(name))
            }
        }
        return List
    }

    ; Shows the main GUI window for managing credentials
    static ShowGUI() {
        global MainGui
        if IsObject(MainGui) {
            MainGui.Show()
            CredManager.LoadCreds(MainGui)
            return
        }
        MainGui := Gui("", "Ahk_CredManager")
        MainGui.SetFont("s12")
        LV1 := MainGui.Add("ListView", "x10 y10 r10 w510 vLV1", ["Name", "User"])
        LV1.OnEvent("ContextMenu", ShowContextMenu)
        MainGui.Add("Button", "x430 y+5 w80 h25", "Create").OnEvent("Click", EditOrNewCred.Bind(""))
        MainGui.SetFont("s10")
        MainGui.Add("StatusBar", "vSB", "") ; Add a status bar named SB
        CredManager.LoadCreds(MainGui)
        MainGui.Show()
        SetIcon(MainGui)
    }

    ; Loads/refreshes the ListView in the main GUI
    static LoadCreds(guiObj) {
        LV1 := guiObj["LV1"]
        LV1.Delete()
        pList := CredManager.ListAll()
        SB := guiObj["SB"] ; Access the status bar by its variable name
        SB.SetText("--- " pList.Length " credentials ---")
        col1 := 1
        loop pList.Length {
            nameCred := CredManager.DelPrefixName(pList[A_Index].name)
            n := LV1.add(, nameCred, pList[A_Index].user)
            LV1.Modify(n, "Vis")
            col1 := 9 * StrLen(nameCred) > col1 ? 9 * StrLen(nameCred) : col1
        }
        LV1.ModifyCol(1, (col1 > 200 ? 200 : col1) " Sort")
        LV1.ModifyCol(2, 490 - (col1 > 200 ? 200 : col1))
        LV1.Opt("+Redraw")
    }
}

; Opens the credential editor window for creating or editing a credential
EditOrNewCred(cred := "", *) {
    eGui := Gui("", "Ahk_CredManager")
    eGui.SetFont("s12")
    eGui.AddText("w300 CGray", "Credential Name")
    eGui.AddEdit("vName xp y+2 w200", cred)
    eGui.AddText("xp y+5 w330 CGray", "Username")
    eGui.AddEdit("vUser xp y+2 w265", cred ? CredManager[cred].user : "")
    eGui.AddText("xp y+2 w330 CGray", "Password")
    eGui.AddEdit("vPass xp y+2 w265 +Password", cred ? CredManager[cred].pass : "")
    ; Checkbox to show/hide password
    eGui.AddCheckbox("x+10 yp h23 -Tabstop", "").OnEvent("Click", (gObj, *) => (
        eGui["Pass"].Opt((gObj.Value ? "-" : "+") . "Password")
    ))
    ; Save button: validates and saves the credential, then refreshes main GUI
    eGui.Add("Button", "x225 y35 w80 h25", "Save").OnEvent("Click", (*) => (
        (eGui["Name"].Value != "" && eGui["Pass"].Value != "") ? (
            (eGui["Name"].Value != cred && cred != "") ? CredManager.Delete(cred) : "",
            CredManager.Write(eGui["Name"].Value, eGui["User"].Value, eGui["Pass"].Value),
            eGui.Destroy(),
            (IsObject(MainGui) ? CredManager.LoadCreds(MainGui) : "")
        ) : MsgBox("Credential fields cannot be empty", "Error", 0x10)
    ))
    eGui.Show("w320")
    SetIcon(eGui)
}

; Shows a context menu for a ListView item (right-click)
ShowContextMenu(LV, Item, IsRightClick, X, Y) {
    credName := LV.GetText(Item, 1)
    if credName != "" {
        CMenu := Menu()
        CMenu.Add("Edit", ObjBindMethod(EditOrNewCred,,credName))
        CMenu.Add("Copy", (*) => A_Clipboard := CredManager[credName].pass)
        CMenu.Add("")
        ; Delete with confirmation dialog
        CMenu.Add("Delete", (*) => (
            (MsgBox("Delete credential '" credName "'?", "Confirm", "YesNo Icon?") = "Yes")
                ? (CredManager.Delete(credName), IsObject(MainGui) ? CredManager.LoadCreds(MainGui) : "")
                : ""
        ))
        CMenu.Show(X, Y)
    }
}

; Sets a custom icon for the GUI window
SetIcon(MyGui) {
    hIcon := Buffer(4)
    DllCall("PrivateExtractIcons", "str", "imageres.dll", "int", 225, "int", 256, "int", 256,
        "Ptr", hIcon, "Ptr", 0, "uint", 1, "uint", 0)
    SendMessage(0x80, 0, hIcon.ptr, , "ahk_id " MyGui.Hwnd)
    SendMessage(0x80, 1, hIcon.ptr, , "ahk_id " MyGui.Hwnd)
}
