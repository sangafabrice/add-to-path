Option Explicit
Dim Arguments : Set Arguments = WScript.Arguments
Dim Named : Set Named = Arguments.Named
Dim UnNamed : Set UnNamed = Arguments.UnNamed
Dim WsShell : Set WsShell = CreateObject("WScript.Shell")
Dim FsoShell : Set FsoShell = CreateObject("Scripting.FileSystemObject")
Dim ScriptPath : ScriptPath = WScript.ScriptFullName
Dim ScriptDir : ScriptDir = FsoShell.GetParentFolderName(ScriptPath)
Dim TempScript : TempScript = ScriptDir & "\set_path_temp.vbs"
Dim IconPath : IconPath = ScriptDir & "\check.ico"
Const ShortcutID = "AddToPath"
Const USERENVPATH_VALUENAME = "HKCU\Environment\Path"
Const SYSTEMENVPATH_VALUENAME = "HKLM\SYSTEM\CurrentControlSet\Control\Session Manager\Environment\Path"
Dim DirBackgroundShell : DirBackgroundShell = "HKCU\SOFTWARE\Classes\Directory\Background\Shell\"
Dim AddToPathKey : AddToPathKey = DirBackgroundShell & ShortcutID
Dim AddToPathShellKey : AddToPathShellKey = AddToPathKey & "\Path\Shell\"
Const ResetListValue = "System.Kind:-""Folder"""
Const ElevatedTaskRoot = "\CustomUI"
Const USERPATH_VALUENAME = "USERPATH"
Const SYSTEMPATH_VALUENAME = "SYSTEMPATH"
Dim IconValueName, PathID, Action

'---------------------------------------------------------------------------------------------------
If Named.Exists("Install") Then
    Select Case UnNamed.Length
        Case 0                                                                                      ' Install -Add to Path- shortcut menus
            IsAllAllowed Array("Install"),_
            "Usage: set-path /Install"
            PriviledgesRequired
            InstallMenu
        Case 1                                                                                      ' Install shortcut-menu paths identifiers
            InstallationRequired
            IsAllAllowed Array("Install", "User", "System"),_
            "Usage: set-path PathID /Install [/User:Path[;...]] /System:Path[;...]" & vbCrLf &_
            "Usage: set-path PathID /Install /User:Path[;...] [/System:Path[;...]]"
            UserOrSystemMustBeSpecifiedAndNotEmpty
            If Named("System") <> "" Then PriviledgesRequired                
            PathIDFirstCharIsNotPlusOrMinus
            PathID = CommandLineArgument
            Action = "+"
            InstallPathID
    End Select
    CleanUpAndQuit(0)
End If

'---------------------------------------------------------------------------------------------------
InstallationRequired
If IsResetArgSet Then
    IsAllAllowed Array("Reset", "Elevate"),_
    "Usage: set-path /Reset [/Elevate]"
    ElevateCommand("Reset")
    PathID = "___PATHID___"
    SetPath GetRef("ResetPath")
    CleanUpAndQuit(0)
End If

'---------------------------------------------------------------------------------------------------
PathIDMustBeProvided
IsAllAllowed Array("Path", "Elevate"),_
"Usage: set-path [-|+]PathID [/Elevate]" & vbCrLf &_
"Usage: set-path -[FolderPath[;...]] [/Path] [/Elevate]" & vbCrLf &_
"Usage: set-path [+][FolderPath[;...]] [/Path[:{System|User}]] [/Elevate]"
ElevateCommand("Path")

'---------------------------------------------------------------------------------------------------
PathID = Mid(CommandLineArgument, 2)
Select Case ArgFirstChar
    Case "-"                                                                                        ' Remove directories from environment variable PATH
        QuitIfPathEmptyAfterParsing
        Action = "+"
        SetPath GetRef("RemoveFromPath")
    Case Else                                                                                       ' Add directories to environment variable PATH
        If ArgFirstChar = "+" Then
            QuitIfPathEmptyAfterParsing
        Else
            PathID = CommandLineArgument
        End If
        Action = "-"
        SetPath GetRef("AddToPath")
End Select
CleanUpAndQuit(0)


'***************************************************************************************************

Public Sub InstallMenu
    ' Implement shortcut menu verbs:
    ' On directory background - "Add to PATH"
    ' On directory object - "Add this to PATH"

    Dim TempReg : TempReg = ScriptDir & "\set_path_setup_temp.reg"
    Dim ReadHandle : Set ReadHandle = FsoShell.OpenTextFile(ScriptDir & "\setup.reg", 1)
    Dim WriteHandle : Set WriteHandle = FsoShell.OpenTextFile(TempReg, 2, True)
    WriteHandle.Write(Replace(Replace(Replace(Replace(Replace(ReadHandle.ReadAll(),_
    "___ICON_PATH___", EscapeSlashChar(ScriptDir & "\main.ico")),_
    "___SHORTCUT_ID___", ShortcutID),_
    "___SCRIPT_PATH___", EscapeSlashChar(ScriptPath)),_
    "___USER_ENV_PATH___", USERPATH_VALUENAME),_
    "___SYSTEM_ENV_PATH___", SYSTEMPATH_VALUENAME))
    ReadHandle.Close()
    WriteHandle.Close()
    Dim ObjExec : Set ObjExec = WsShell.Exec("Reg Import " & TempReg & " /Reg:64")
    Do : Loop Until objExec.Status = 1
    FsoShell.DeleteFile TempReg
    WsShell.Exec("SchTasks /Create /SC ONCE /TN " & ElevatedTaskRoot & "\" & ShortcutID & _
    " /TR " & TempScript & " /ST 00:00 /SD 01/01/2022 /RL HIGHEST /F")
    UpdateDirectoryShellObject
    Set ObjExec = Nothing
    Set ReadHandle = Nothing
    Set WriteHandle = Nothing
End Sub

Public Sub InstallPathID
    ' Implement a Path ID shortcut submenu verb
    ' on directory background - "Add to PATH"

    On Error Resume Next
    Dim PathList : Set PathList = CreateObject("Scripting.Dictionary")
    Dim KeyList : Set KeyList = CreateObject("Scripting.Dictionary")
    Dim paramName
    For Each paramName In Named
        If LCase(paramName) <> "install" Then
            Dim paramPattern
            Dim paramPath : For Each paramPath In Split(Named(paramName),";")
                Dim ExpandedPathArg
                PathMustExist paramPath, ExpandedPathArg
                Dim AbsolutePathArg : AbsolutePathArg = FsoShell.GetFolder(ExpandedPathArg).Path
                If FsoShell.GetDriveName(ExpandedPathArg) = "" Then paramPath = AbsolutePathArg
                paramPattern = paramName & "*"
                PathList.Add AbsolutePathArg, paramPattern & paramPath
            Next
            KeyList.Add StoredPathKey(paramName),_
            Replace(Join(Filter(PathList.Items, paramPattern, True, vbTextCompare), ";"),_
            paramPattern, "")
        End If
    Next
    For Each paramName In Array("User", "System") : WsShell.RegDelete StoredPathKey(paramName) : Next
    Dim pKey : For Each pKey In KeyList.Keys : WsShell.RegWrite pKey, KeyList.Item(pKey) : Next
    RegWriteCommand
    Set PathList = Nothing
    Set KeyList = Nothing
End Sub

Public Sub SetPath(ModificationFuntionHandle)
    ' Modify Path Environment Variable
    ' Change UI accordingly:
    ' Check/Uncheck PathID in directory background
    ' Switch verbs between "Remove/Add to Path"

    IconValueName = PathIDKey & "\Icon"
    Dim pathKey : For Each pathKey In GetStoredPath.Keys : ModificationFuntionHandle pathKey : Next
    UpdateDirectoryShellObject
    If IsPathArgSet Or IsResetArgSet Then Exit Sub
    RegWriteCommand
End Sub

Public Sub RemoveFromPath(PathKey)
    On Error Resume Next
    Dim PathEnvValueName : For Each PathEnvValueName In Array(_
        USERENVPATH_VALUENAME, SYSTEMENVPATH_VALUENAME)
        Dim pEnv : pEnv = GetPathEnv(PathEnvValueName)
        Dim pathList : pathList = Split(pEnv, ";")
        Dim index : For Each index In InPath(pEnv, PathKey) : pathList(index) = "" : Next
        RegWritePath PathEnvValueName, CleanPath(Join(pathList, ";"))
    Next
    If IsPathArgSet Then Exit Sub
    WsShell.RegDelete IconValueName
End Sub

Public Sub AddToPath(PathKey)
    Dim PathKeyDict : Set PathKeyDict = GetStoredPath.Item(PathKey)
    Dim PathEnvType : PathEnvType = Join(PathKeyDict.Keys)
    Dim PathEnvValueName : PathEnvValueName = GetEnvironmentKey(PathEnvType)
    Dim pEnv : pEnv = GetPathEnv(PathEnvValueName)
    Dim InPathCopy : InPathCopy = InPath(pEnv, PathKey)
    Dim InPathCopyUBound : InPathCopyUBound = UBound(InPathCopy)
    Dim pathList : pathList = Split(pEnv, ";")
    Dim index : For index = 1 To InPathCopyUBound : pathList(InPathCopy(index)) = "" : Next
    If InPathCopyUBound < 0 Then
        Dim pathListUBound : pathListUBound = UBound(pathList) + 1
        Redim Preserve pathList(pathListUBound)
        pathList(pathListUBound) = PathKeyDict.Item(PathEnvType)
    End If
    RegWritePath PathEnvValueName, CleanPath(Join(pathList, ";"))
    Set PathKeyDict = Nothing
    If IsPathArgSet Then Exit Sub
    WsShell.RegWrite IconValueName, IconPath
End Sub

Public Sub ResetPath(PathKey)
    On Error Resume Next
    Const HKCU = &H80000001
    Dim BgVerbs
    Action = "+"
    GetObject("winmgmts:\\.\root\default:StdRegProv").EnumKey HKCU, Mid(AddToPathShellKey, 6), BgVerbs
    Dim dotPathID : For Each dotPathID In BgVerbs
        If Left(dotPathID, 1) = "." Then
            PathID = Mid(dotPathID, 2)
            Err.Clear
            WsShell.RegDelete Replace(IconValueName, "___PATHID___", PathID)
            If Err.Number = 0 Then RegWriteCommand
        End If
    Next
    RegWritePath GetEnvironmentKey(PathKey), CleanPath(GetStoredPath.Item(PathKey))
End Sub 

Private Function GetStoredPath
    ' Parse /Path argument
    ' Tokenize USERPATH/SYSTEMPATH value names when /Path is not set
    ' Store directory path tokens in dictionary:
    ' Key: Directory full path
    ' Item: A single item dictionary 
    ' (Key: USERPATH/SYSTEMPATH, Item: Unexpanded Path)

    On Error Resume Next
    Set GetStoredPath = CreateObject("Scripting.Dictionary")
    Dim pathType
    If IsResetArgSet Then
        Dim arrPathType : arrPathType = Array(Named.Item("Reset"))
        If UBound(arrPathType) = 0 And (arrPathType(0) = "" Or arrPathType(0) = "*") Then
            arrPathType = Array("User", "System")
        End If
        For Each pathType In arrPathType
            pathType = GetPathType(pathType)
            GetStoredPath.Add pathType, WsShell.RegRead(AddToPathKey & "\" & pathType)
        Next
        Exit Function
    End If
    If IsPathArgSet Then
        pathType = Named.Item("Path")
        If pathType = "" Then pathType = "User"
        SetStoredPathDictionary GetStoredPath, GetPathType(pathType), PathID
        Exit Function
    End If
    For Each pathType In Array(SYSTEMPATH_VALUENAME, USERPATH_VALUENAME)
        Err.Clear
        Dim fullPath : fullPath = WsShell.RegRead(PathIDKey & "\" & pathType)
        If Err.Number = 0 Then SetStoredPathDictionary GetStoredPath, pathType, fullPath
    Next
End Function

Private Sub SetStoredPathDictionary(ByRef StoredPathDico, pathType, fullPath)
    Dim indivPath : For Each indivPath In Split(fullPath, ";")
        Dim ExpandedPathArg : ExpandedPathArg = WsShell.ExpandEnvironmentStrings(indivPath)
        If FsoShell.FolderExists(ExpandedPathArg) Then
            Dim UnExpandedPath : Set UnExpandedPath = CreateObject("Scripting.Dictionary")
            UnExpandedPath.Add pathType, indivPath
            StoredPathDico.Add FsoShell.GetFolder(ExpandedPathArg).Path, UnExpandedPath
            Set UnExpandedPath = Nothing
        End If
    Next
End Sub

Private Function InPath(PathEnvVarString, PathItem)
    ' Find the positions of a folder path
    ' in PATH environment variable

    InPath = Array()
    Dim InPathCopy()
    Dim InPathSize : InPathSize = 0
    Dim Counter : Counter = 0
    Dim path : For Each path In Split(WsShell.ExpandEnvironmentStrings(PathEnvVarString), ";")
        If FsoShell.FolderExists(path) And FsoShell.GetFolder(path).Path = PathItem Then
            Redim Preserve InPathCopy(InPathSize)
            InPathCopy(InPathSize) = Counter
            InPathSize = InPathSize + 1
        End If
        Counter = Counter + 1
    Next
    If InPathSize > 0 Then InPath = InPathCopy
End Function

Private Sub ElevateCommand(ArgumentName)
    ' Elevate the [set-path.vbs ...] command
    ' when priviledges are required
    
    If Not Named.Exists("Elevate") Then Exit Sub
    If TestPriviledges Then Exit Sub
    Dim NamedArgument
    Dim FileHandle : Set FileHandle = FsoShell.OpenTextFile(TempScript, 2, True)
    If Named.Exists(ArgumentName) Then
        Dim i : For i = 0 To Arguments.Length - 1
            If UCase(Left(Arguments(i), Len(ArgumentName) + 1)) = "/" & UCase(ArgumentName) Then
                NamedArgument = " " & Arguments(i)
                Exit For
            End If
        Next
    End If
    PathID = ""
    If Not IsEmpty(CommandLineArgument) Then PathID = """""" & CommandLineArgument & """"""
    FileHandle.Write("CreateObject(""WScript.Shell"").Run """ &_
    GetCommandLine(PathID & NamedArgument) & """, 0, True")
    FileHandle.Close()
    Dim SchTasks : Set SchTasks = CreateObject("Schedule.Service")
    SchTasks.Connect()
    Dim TaskInstance : Set TaskInstance = SchTasks.GetFolder(ElevatedTaskRoot).GetTask(ShortcutID)
    TaskInstance.Run Null
    Do : Loop Until TaskInstance.State <> 4
    FsoShell.DeleteFile TempScript
    Set TaskInstance = Nothing
    Set SchTasks = Nothing
    Set FileHandle = Nothing
    CleanUpAndQuit(0)
End Sub

Private Sub UpdateDirectoryShellObject
    ' Update AppliesTo values with Parsed Paths
    ' that are the set of folders from which menus can display
    ' AND-AQS helps exclude folders : menu option hidden
    ' OR-AQS helps include folders: menu option visible

    Dim UPath : UPath = UCase(GetPathEnv(USERENVPATH_VALUENAME))
    Dim SPath : SPath = UCase(GetPathEnv(SYSTEMENVPATH_VALUENAME))
    Dim UPathAndAQS : UPathAndAQS = ParseCanonicalAQS(UPath, "AND")
    Dim SPathAndAQS : SPathAndAQS = ParseCanonicalAQS(SPath, "AND")
    Dim DirVerbKey : DirVerbKey = Replace(DirBackgroundShell, "\Background", "") & ShortcutID
    UpdateAppliesToKey DirVerbKey, UPathAndAQS, SPathAndAQS
    UpdateAppliesToKey DirVerbKey & ".Reverse", ParseCanonicalAQS(UPath, "OR"), SPathAndAQS
    UpdateAppliesToKey DirVerbKey & ".Reverse.Admin", UPathAndAQS, ParseCanonicalAQS(SPath, "OR")
End Sub


'***************************************************************************************************

Private Sub IsAllAllowed(AllowedParameterList, UsageMessage)
    Dim AllowedParameters : Set AllowedParameters = CreateObject("Scripting.Dictionary")
    Dim paramName
    For Each paramName In AllowedParameterList : AllowedParameters.Add UCase(paramName), "" : Next
    For Each paramName In WScript.Arguments.Named
        If Not AllowedParameters.Exists(UCase(paramName)) Then
            WScript.Echo """/" & paramName & """ unexpected."
            WScript.Echo UsageMessage
            Set AllowedParameters = Nothing
            CleanUpAndQuit(2)
        End If
    Next
    Set AllowedParameters = Nothing
End Sub

Private Sub PathIDFirstCharIsNotPlusOrMinus
    If InStr("+-", ArgFirstChar) > 0 Then
        WScript.Echo ArgFirstChar & " unexpected."
        CleanUpAndQuit(1)
    End If
End Sub

Private Sub PriviledgesRequired
    If Not TestPriviledges Then
        WScript.Echo "Elevated priviledges required."
        CleanUpAndQuit(1)
    End If
End Sub

Private Sub UserOrSystemMustBeSpecifiedAndNotEmpty
    Dim Counter : Counter = 0
    Dim pName : For Each pName In Named
        If LCase(pName) <> "install" And Named(pName) = "" Then
            Counter = Counter + 1
            If Counter = 2 Then
                WScript.Echo "User or System must be specified and not both empty."
                CleanUpAndQuit(1)
            End If
        End If
    Next
End Sub

Private Sub InstallationRequired
    On Error Resume Next
    Err.Clear
    WsShell.RegRead(AddToPathKey & "\")
    If Err.Number <> 0 Then
        Dim ObjExec : Set ObjExec = WsShell.Exec("Cscript " & ScriptPath & " /Install //NoLogo")
        Do : WScript.Echo objExec.StdOut.ReadLine() : Loop Until objExec.StdOut.AtEndOfStream
        If objExec.ExitCode <> 0 Then
            Set ObjExec = Nothing
            CleanUpAndQuit(1)
        End If
        Set ObjExec = Nothing
    End If
End Sub

Private Sub PathMustExist(PathArg, ByRef ExpandedPathArg)
    ExpandedPathArg = WsShell.ExpandEnvironmentStrings(PathArg)
    If Not FsoShell.FolderExists(ExpandedPathArg) Then
        WScript.Echo PathArg & " cannot be found."
        CleanUpAndQuit(1)
    End If
End Sub

Private Sub PathIDMustBeProvided
    If UnNamed.Length <> 1 Then CleanUpAndQuit(1)
End Sub

Private Sub QuitIfPathEmptyAfterParsing
    If PathID = "" Then CleanUpAndQuit(1)
End Sub

Private Sub QuitIfPathArgUnknown(PathArg)
    PathArg = UCase(PathArg)
    If PathArg <> "USER" And PathArg <> "SYSTEM" Then
        WScript.Echo PathArg & " unknown value."
        CleanUpAndQuit(1)
    End If
End Sub


'---------------------------------------------------------------------------------------------------

Private Function TestPriviledges
    On Error Resume Next
    Err.Clear
    WsShell.RegRead("HKEY_USERS\S-1-5-19\Environment\TEMP")
    TestPriviledges = Err.Number = 0
End Function

Private Function GetPathType(pathType)
    QuitIfPathArgUnknown(pathType)
    GetPathType = GetValueName(pathType)
End Function

Private Function CommandLineArgument
    CommandLineArgument = ""
    If UnNamed.Length <> 0 Then CommandLineArgument = UnNamed.Item(0)
End Function

Private Function ArgFirstChar : ArgFirstChar = Left(CommandLineArgument, 1) : End Function

Private Function GetCommandLine(LineArgument)
    GetCommandLine = ScriptPath & " " & LineArgument
End Function

Private Function IsPathArgSet : IsPathArgSet = Named.Exists("Path") : End Function

Private Function IsResetArgSet : IsResetArgSet = Named.Exists("Reset") : End Function

Private Function StoredPathKey(PathType)
    StoredPathKey =  PathIDKey & "\" & GetValueName(PathType)
End Function

Private Function EscapeSlashChar(PathString)
    EscapeSlashChar = Replace(PathString, "\", "\\")
End Function

Private Function CleanPath(PathValue)
    strCleanPath PathValue, ";"
    strCleanPath PathValue, "\"
    CleanPath = PathValue
End Function

Private Sub strCleanPath(PathValue, delimiter)
    PathValue = Replace(PathValue, delimiter & ";", ";")
    If Right(PathValue, 1) = delimiter Then PathValue = Left(PathValue, Len(PathValue) - 1)
End Sub

Private Function GetPathEnv(ENVPATH_VALUENAME)
    On Error Resume Next
    GetPathEnv = CleanPath(WsShell.RegRead(ENVPATH_VALUENAME))
End Function

Private Function GetEnvironmentKey(PathValueName)
    Select Case PathValueName
        Case USERPATH_VALUENAME : GetEnvironmentKey = USERENVPATH_VALUENAME
        Case SYSTEMPATH_VALUENAME : GetEnvironmentKey = SYSTEMENVPATH_VALUENAME
    End Select
End Function

Private Sub UpdateAppliesToKey(DirVerbKey, UserPathsAQS, SystemPathsAQS)
    WsShell.RegWrite DirVerbKey & "\AppliesTo", UserPathsAQS & " AND " & SystemPathsAQS
End Sub

Private Function ParseCanonicalAQS(PathString, LogicalOperator)
    Dim AqsSign : AqsSign = "System.ItemPathDisplay:" & "-" & """"
    If LogicalOperator = "OR" Then AqsSign = Replace(AqsSign, "-", "")
    ParseCanonicalAQS = AqsSign & Replace(WsShell.ExpandEnvironmentStrings(PathString),_
    ";", """ " & LogicalOperator & " " & AqsSign) & """"
End Function

Private Sub RegWritePath(EnvKey, PathString)
    WsShell.RegWrite EnvKey, PathString, "REG_EXPAND_SZ"
End Sub

Private Sub RegWriteCommand
    WsShell.RegWrite PathIDKey & "\", PathID
    WsShell.RegWrite PathIDKey & "\Command\", "Wscript.exe " &_
    GetCommandLine(Action & PathID) & " /Elevate"
End Sub

Private Function PathIDKey
    PathIDKey = AddToPathShellKey & "." & PathID
End Function

Private Function GetValueName(PathType) : GetValueName = UCase(PathType) & "PATH" : End Function

Private Sub CleanUp
    On Error Resume Next
    Set WsShell = Nothing
    Set FsoShell = Nothing
    Set Named = Nothing
    Set UnNamed = Nothing
    Set Arguments = Nothing
End Sub

Private Sub CleanUpAndQuit(ExitCode)
    CleanUp
    WScript.Quit(ExitCode)
End Sub