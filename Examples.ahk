#Requires AutoHotkey v2.0
#SingleInstance

#Include <v2\GDIp\Gdip_All>
#Include <v2\IconExtract\IconExtract>

;==============================================

if (!pToken := Gdip_Startup())
    MsgBox('GDI+ failed to start. Please ensure you have GDI+ on your system.',, 'Iconx'), ExitApp()

OnExit((*) => Gdip_Shutdown(pToken))

;==============================================

; Save icon group index 2 from shell32.dll as a multi-resolution ICO file (Desktop, auto-generated filename)
try {
    savedPath := IconExtract.SaveIconToIco(A_WinDir '\System32\shell32.dll', 2, A_Desktop)
    MsgBox('ICO saved to:`n' savedPath,, 'iconi')
} catch as err {
    MsgBox('SaveIconToIco failed:`n' err.Message,, 'iconx')
}

; Save the highest quality variant from icon group index 4 in imageres.dll as a PNG file (explicit output file path)
try {
    savedPath := IconExtract.SaveIconToPng(A_WinDir '\System32\imageres.dll', 4, A_Desktop '\imageres-dll-4.png')
    MsgBox('PNG saved to:`n' savedPath,, 'iconi')
} catch as err {
    MsgBox('SaveIconToPng failed:`n' err.Message,, 'iconx')
}

; Save all icon groups from the AutoHotkey executable as separate multi-resolution ICO files (Desktop\AHK-Icons)
try {
    savedFiles := IconExtract.SaveAllIconsToIco(A_AhkPath, A_Desktop '\AHK-Icons')
    MsgBox('Extracted ' savedFiles.Length ' ICO file(s).`nFirst file:`n' savedFiles[1],, 'iconi')
} catch as err {
    MsgBox('SaveAllIconsToIco failed:`n' err.Message,, 'iconx')
}

; Save the highest quality variant from each icon group in notepad.exe as PNG files (auto-generated output directory)
try {
    savedFiles := IconExtract.SaveAllIconsToPng(A_WinDir '\System32\notepad.exe')
    MsgBox('Extracted ' savedFiles.Length ' PNG file(s).`nFirst file:`n' savedFiles[1],, 'iconi')
} catch as err {
    MsgBox('SaveAllIconsToPng failed:`n' err.Message,, 'iconx')
}