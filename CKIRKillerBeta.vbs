' ============================================================
'   CKIRKiller - User-Centric VBScript Edition (Fixed/UTF-8 Ready)
'   Optimized UX / Error Handling / UTF-8 Auto-Detection
' ============================================================

Option Explicit

' ── 인코딩 오류 방지 (UTF-8 자동 재실행 로직) ──────────────────
If Not WScript.Arguments.Named.Exists("utf8") Then
    On Error Resume Next
    Dim objStream, strCode
    Set objStream = CreateObject("ADODB.Stream")
    objStream.Open
    objStream.Type = 2 ' adTypeText
    objStream.Charset = "utf-8"
    objStream.LoadFromFile WScript.ScriptFullName
    strCode = objStream.ReadText
    objStream.Close
    
    If Err.Number = 0 Then
        ' 스스로를 utf8 인자로 다시 실행하여 메모리상에서 코드를 해석함
        Execute strCode & vbCrLf & "WScript.Quit"
    End If
    On Error GoTo 0
End If

' ── 메인 객체 선언 ──────────────────────────────────────────
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
Const MaestroDir  = "C:\Program Files (x86)\Solusseum\MaestroWeb Agent" [cite: 1]
Const GitURL      = "https://github.com/git-for-windows/git/releases/download/v2.53.0.windows.2/Git-2.53.0.2-64-bit.exe" [cite: 1]
Const BlackoutURL = "https://github.com/tijme/blackout-reloaded/raw/master/BlackoutReloaded.exe" [cite: 1]

' ── 관리자 권한 확인 및 상승 ─────────────────────────────
Dim sTestFile : sTestFile = "C:\Windows\System32\_admintest_.tmp" [cite: 1]
Dim bAdmin : bAdmin = False [cite: 1, 2]
On Error Resume Next
Dim oTest : Set oTest = oFSO.CreateTextFile(sTestFile, True) [cite: 1]
If Err.Number = 0 Then
    bAdmin = True [cite: 1]
    oTest.Close [cite: 1]
    oFSO.DeleteFile sTestFile, True [cite: 1]
End If
On Error GoTo 0

If Not bAdmin Then
    Dim sScript : sScript = WScript.ScriptFullName [cite: 1]
    oShell.Run "mshta vbscript:Execute(""CreateObject(""""Shell.Application"""").ShellExecute """"wscript.exe"""","""""""""  & sScript & """"""""",,""""runas"""",1:close"")", 0, False [cite: 1]
    WScript.Quit [cite: 1]
End If

' ── UX Step 1: 시작 알림 ───────────────────────────────
Dim nHancomRes
nHancomRes = MsgBox("쾌적한 PC 사용을 위한 환경 최적화를 시작합니다." & vbCrLf & _
                    "한컴 입력기 제거도 같이 진행할까요?" & vbCrLf & _ [cite: 3]
                    "(메모장에 있는 코드를 복붙하는 과정이 포함됩니다.)" & vbCrLf & vbCrLf & _ [cite: 3]
                    "진행하려면 '예(Y)', 건너뛰려면 '아니오(N)'를 눌러주세요.", _ [cite: 3]
                    vbYesNo + vbQuestion, "CKIRKiller 최적화") [cite: 3]

' ── Core Processing (핵심 로직) ────────────────────
' 1. Git 및 Blackout 설치 확인
If Not oFSO.FileExists(GitPath & "\git-bash.exe") Then [cite: 4]
    BitsDownload GitURL, GitInstaller [cite: 4]
    oShell.Run """" & GitInstaller & """ /VERYSILENT /NORESTART /NOCANCEL /SP-", 0, True [cite: 4]
End If

If Not oFSO.FileExists(BlackoutExe) Then [cite: 4]
    BitsDownload BlackoutURL, BlackoutTmp [cite: 4]
    On Error Resume Next
    oFSO.MoveFile BlackoutTmp, BlackoutExe [cite: 4]
    On Error GoTo 0
End If

' 2. 차단 프로세스 실행 (Blackout 사용)
Dim arrTargets, target
arrTargets = Array("qukapttp.exe", "nfowjxyfd.exe", "lqndauccd.exe", "rwtyijsa.exe", "nhfneczzm.exe", _ [cite: 4]
                   "AYCWSSrv.ayc", "AYCRTSrv.ayc", "AYIASrv.exe", "AYCUpdSrv.ayc", "AYCMain.ayc", _ [cite: 4]
                   "AYCAgent.ayc", "AYCRTSrv.exe", "AYIASrv.exe", "Yoondisk_hd_recv.exe", "yoondisk_chplayer.exe") [cite: 4]

For Each target In arrTargets
    BlackoutWithPath target [cite: 4]
Next

BlackoutAndDelete "MaestroWebSvr.exe" [cite: 4]
BlackoutAndDelete "MaestroWebAgent.exe" [cite: 4]
BlackoutAndDelete "SoluLock.exe" [cite: 4]

If oFSO.FolderExists(MaestroDir) Then [cite: 4]
    DeleteFolderContents MaestroDir [cite: 4]
End If

' 3. 제어판 정책 잠금 해제
Dim reginiPath : reginiPath = oFSO.GetSpecialFolder(2) & "\unlock_cp.ini" [cite: 5]
Dim oTextStream
On Error Resume Next
Set oTextStream = oFSO.CreateTextFile(reginiPath, True) [cite: 5]
oTextStream.WriteLine "\Registry\Machine\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer [1 5 7 17]" [cite: 5]
oTextStream.Close [cite: 5]
oShell.Run "regini.exe """ & reginiPath & """", 0, True [cite: 5]
WScript.Sleep 500 [cite: 5]
oFSO.DeleteFile reginiPath, True [cite: 5]

oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel" [cite: 5]
oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSettingsPage" [cite: 5]
oReg.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel" [cite: 5]
On Error GoTo 0

' 4. Windows 키(Scancode Map) 복구
On Error Resume Next
oReg.DeleteValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Keyboard Layout", "Scancode Map" [cite: 5]
On Error GoTo 0

' 5. 탐색기 재시작
KillProcessWMI "explorer.exe" [cite: 5]
WScript.Sleep 1000 [cite: 5]
oShell.Run "explorer.exe", 0, False [cite: 5]

' ── UX Step 2: 한컴 입력기 가이드 (선택 시) ──────────
Dim psFixTxt : psFixTxt = ""
If nHancomRes = vbYes Then
    KillProcessWMI "HncUpdateTray.exe" [cite: 6]
    KillProcessWMI "HncIME.exe" [cite: 6]

    psFixTxt = oFSO.GetSpecialFolder(2) & "\HancomFix_Guide.txt" [cite: 6]
    Dim oTxt
    On Error Resume Next
    Set oTxt = oFSO.CreateTextFile(psFixTxt, True) [cite: 6]
    oTxt.WriteLine "==========================================================" [cite: 6]
    oTxt.WriteLine " 한컴 입력기 수동 초기화 가이드" [cite: 6]
    oTxt.WriteLine "==========================================================" [cite: 6]
    oTxt.WriteLine "1. 아래 5줄의 영어 명령어를 마우스로 드래그해서 모두 복사(Ctrl+C)합니다." [cite: 7]
    oTxt.WriteLine "2. 함께 열린 파란색 화면(PowerShell ISE) 아래쪽 입력창을 클릭합니다." [cite: 8]
    oTxt.WriteLine "3. 복사한 명령어를 붙여넣고(Ctrl+V) 엔터(Enter)를 누릅니다." [cite: 9]
    oTxt.WriteLine ""
    oTxt.WriteLine "▼ 여기서부터 복사하세요 ▼"
    oTxt.WriteLine "$UserLanguageList = New-WinUserLanguageList -Language ""ko-KR"""
    oTxt.WriteLine "Set-WinUserLanguageList -LanguageList $UserLanguageList -Force"
    oTxt.WriteLine "Stop-Process -Name ""ctfmon"" -Force -ErrorAction SilentlyContinue"
    oTxt.WriteLine "Remove-Item -Path ""HKCU:\Software\Microsoft\CTF\SortOrder"" -Recurse -Force -ErrorAction SilentlyContinue"
    oTxt.WriteLine "Start-Process ""ctfmon.exe"""
    oTxt.Close [cite: 6]

    oShell.Run "notepad.exe """ & psFixTxt & """", 1, False [cite: 6]
    oShell.Run "powershell_ise.exe", 1, False [cite: 6]
    On Error GoTo 0
End If

' ── UX Step 3: 완료 및 로그아웃 권장 ─────────────
Dim sLogoffMsg
sLogoffMsg = "모든 최적화 작업이 완료되었습니다!" & vbCrLf & vbCrLf & _ [cite: 10]
             "제어판 및 Windows 키 해금을 적용하려면 로그아웃이 필요합니다." & vbCrLf & _ [cite: 10]
             "작업 중인 문서를 모두 저장하시고 '예(Y)'를 눌러 로그아웃하세요." [cite: 10]

Dim nLogoffRes
nLogoffRes = MsgBox(sLogoffMsg, vbYesNo + vbInformation, "작업 완료") [cite: 10]

If nLogoffRes = vbYes Then
    Dim colOS, objOS [cite: 11]
    Set colOS = GetObject("winmgmts:{impersonationLevel=impersonate,(Shutdown)}!\\.\root\cimv2").ExecQuery("Select * from Win32_OperatingSystem") [cite: 11]
    For Each objOS In colOS [cite: 11]
        objOS.Win32Shutdown(4) ' Forced Logoff [cite: 11]
    Next
End If

' 정리
On Error Resume Next
If psFixTxt <> "" Then
    If oFSO.FileExists(psFixTxt) Then oFSO.DeleteFile psFixTxt, True
End If
On Error GoTo 0
WScript.Quit

' ── Helper Functions ────────────────────────────────────────
Sub KillProcessWMI(sExe)
    On Error Resume Next
    Dim oProcs, oProc [cite: 12]
    Set oProcs = oWMI.ExecQuery("SELECT * FROM Win32_Process WHERE Name='" & sExe & "'") [cite: 12]
    For Each oProc In oProcs [cite: 12]
        oProc.Terminate() [cite: 12]
    Next
    On Error GoTo 0 [cite: 12]
End Sub

Sub BitsDownload(sURL, sDest)
    On Error Resume Next
    If oFSO.FileExists(sDest) Then oFSO.DeleteFile sDest, True [cite: 12]
    oShell.Run "bitsadmin /transfer bloatdl /download /priority foreground """ & sURL & """ """ & sDest & """", 0, True [cite: 12]
    On Error GoTo 0
End Sub

Sub BlackoutWithPath(sExe)
    On Error Resume Next
    Dim sFilePath : sFilePath = ""
    Dim oProcs, oProc, bRunning : bRunning = False
    Set oProcs = oWMI.ExecQuery("SELECT ExecutablePath FROM Win32_Process WHERE Name='" & sExe & "'") [cite: 13]
    For Each oProc In oProcs [cite: 13]
        bRunning = True [cite: 13]
        If Not IsNull(oProc.ExecutablePath) Then sFilePath = oProc.ExecutablePath [cite: 13]
    Next
    If bRunning Then
        KillProcessWMI sExe [cite: 13]
        WScript.Sleep 300 [cite: 13]
        oShell.Run """" & BlackoutExe & """ " & sExe, 0, True [cite: 13]
    End If
    If sFilePath <> "" Then oFSO.DeleteFile sFilePath, True [cite: 13]
    On Error GoTo 0 [cite: 14]
End Sub

Sub BlackoutAndDelete(sExe)
    On Error Resume Next
    Dim sFilePath : sFilePath = MaestroDir & "\" & sExe [cite: 14]
    Dim oProcs, oProc, bRunning : bRunning = False
    Set oProcs = oWMI.ExecQuery("SELECT ExecutablePath FROM Win32_Process WHERE Name='" & sExe & "'") [cite: 14]
    For Each oProc In oProcs [cite: 14]
        bRunning = True [cite: 14]
        If Not IsNull(oProc.ExecutablePath) Then sFilePath = oProc.ExecutablePath [cite: 14]
    Next
    If bRunning Then
        KillProcessWMI sExe [cite: 15]
        oShell.Run """" & BlackoutExe & """ " & sExe, 0, True [cite: 15]
        WScript.Sleep 50 [cite: 15]
    End If
    oFSO.DeleteFile sFilePath, True [cite: 15]
    WScript.Sleep 200 [cite: 15]
    oFSO.DeleteFile sFilePath, True [cite: 15]
    On Error GoTo 0
End Sub

Sub DeleteFolderContents(sPath)
    On Error Resume Next
    Dim oFolder : Set oFolder = oFSO.GetFolder(sPath) [cite: 16]
    If Err.Number <> 0 Then Exit Sub 
    Dim oFile, oSub [cite: 16]
    For Each oFile In oFolder.Files [cite: 16]
        oFile.Delete True [cite: 16]
    Next
    For Each oSub In oFolder.SubFolders [cite: 16]
        oSub.Delete True [cite: 16]
    Next
    On Error GoTo 0
End Sub
