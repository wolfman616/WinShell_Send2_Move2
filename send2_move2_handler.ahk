#NoEnv ; (MW:2023) (MW:2023)
#NoTrayicon
SetBatchLines,-1
ListLines,Off
#Persistent
#Singleinstance,Force
a_scriptStartTime:= a_tickCount
; Setworkingdir,% (splitpath(A_AhkPath)).dir
DetectHiddenWindows,On
DetectHiddenText,	On
SetTitleMatchMode,2
SetTitleMatchMode,Slow

if(!A_Args[2])
	exitapp,

targetDir:= A_Args[1]

for,i in A_Args
	max_i:= i

fileList:= ""
loop,% max_i
	fileList.= A_Args[a_index+1] "`n"

try,FileToClipboard(fileList,"Cut")
	invokeVerb(targetDir,"paste")
exitapp,

FileToClipboard(PathToCopy,Method="Cut") {
	FileCount:=0, PathLength:=0, Offset:=0
	Loop,Parse,PathToCopy,`n,`r
	{
		FileCount++
		PathLength+=StrLen(A_LoopField)
	}
	pid:=DllCall("GetCurrentProcessId","uint")
	hwnd:=WinExist("ahk_pid " . pid) 	; 0x42 = GMEM_MOVEABLE(0x2) | GMEM_ZEROINIT(0x40)
	hPath := DllCall("GlobalAlloc","uint",0x42,"uint",20 + (PathLength + FileCount + 1) * 2,"UPtr")
	pPath := DllCall("GlobalLock","UPtr",hPath)
	NumPut(20,pPath+0),pPath += 16 ; DROPFILES.pFiles = offset of file list
	NumPut(1,pPath+0),pPath += 4 ; fWide = 0 -->ANSI,fWide = 1 -->Unicode
	Loop,Parse,PathToCopy,`n,`r ; Rows are delimited by linefeeds (`r`n).
		offset += StrPut(A_LoopField,pPath+offset,StrLen(A_LoopField)+1,"UTF-16") * 2
	DllCall("GlobalUnlock","UPtr",hPath)
	DllCall("OpenClipboard","UPtr",hwnd)
	DllCall("EmptyClipboard")
	DllCall("SetClipboardData","uint",0xF,"UPtr",hPath) ; 0xF = CF_HDROP
		; Write Preferred DropEffect structure to clipboard to switch between copy/cut operations
		; 0x42 = GMEM_MOVEABLE(0x2) | GMEM_ZEROINIT(0x40)
	mem := DllCall("GlobalAlloc","uint",0x42,"uint",4,"UPtr")
	str := DllCall("GlobalLock","UPtr",mem)
	if (Method="copy")
		DllCall("RtlFillMemory","UPtr",str,"uint",1,"UChar",0x05)
	Else if (Method="cut")
		DllCall("RtlFillMemory","UPtr",str,"uint",1,"UChar",0x02)
	Else {
		DllCall("CloseClipboard")
		Return
	}
	DllCall("GlobalUnlock","UPtr",mem)
	cfFormat := DllCall("RegisterClipboardFormat","Str","Preferred DropEffect")
	DllCall("SetClipboardData","uint",cfFormat,"UPtr",mem)
	DllCall("CloseClipboard")
	Return,!errorlevel
}

InvokeVerb(Path,Menu,validate=true) {
	objShell:= ComObjCreate("Shell.Application")
	if(InStr(FileExist(Path),"D") || InStr(Path,"::{")) {
		objFolder:= objShell.NameSpace(Path)
		, objFolderItem:= objFolder.Self
	} else {
		SplitPath,Path,Name,Dir
		objFolder:= objShell.NameSpace(dir)
		, objFolderItem:= objFolder.ParseName(Name)
	} if(validate) {
		colVerbs:= objFolderItem.Verbs
		loop,% colVerbs.Count {
			Verb:= colVerbs.Item(A_Index- 1)
			, RetMenu:= Verb.Name
			StringReplace,RetMenu,RetMenu,&
			if(RetMenu=Menu) {
				Verb.DoIt
				return,True
			}
		} return,False
	} else,objFolderItem.InvokeVerbEx(Menu)
}