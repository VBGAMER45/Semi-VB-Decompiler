Attribute VB_Name = "modCmdLine"
'*********************************************
'modCmdLine
'Headless / command-line driver for Semi VB Decompiler.
'
'  SemiVBDecompiler.exe <inputfile> [/out <dir>] [/vbp] [/dism] [/solution]
'
'  <inputfile>   VB4/5/6 or .NET exe/dll/ocx to decompile (first plain arg)
'  /out <dir>    where the generated project / .NET solution is written
'  /vbp          generate a VB project (.vbp + sources)
'  /dism         generate a VB project using raw disassembly bodies
'  /solution     build a .NET solution (C# + VB.NET)
'
'When the requested action does not match the file type the correct one is
'chosen automatically (VB -> .vbp, .NET -> solution).  No message boxes are
'shown; a decompile.log is written and the process exits with 0 (ok) or
'non-zero (error).
'*********************************************
Option Explicit

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

'True when the program was launched with command-line arguments.
Public Function HasCommandLine() As Boolean
    HasCommandLine = (Len(Trim$(Command$)) > 0)
End Function

'Split a command line into tokens, honouring "quoted paths".
Private Function Tokenize(ByVal s As String) As Collection
    Dim col As Collection
    Set col = New Collection
    Dim i As Long, cur As String, inQ As Boolean
    cur = ""
    inQ = False
    For i = 1 To Len(s)
        Dim ch As String
        ch = Mid$(s, i, 1)
        If ch = """" Then
            inQ = Not inQ
        ElseIf (ch = " " Or ch = vbTab) And Not inQ Then
            If Len(cur) > 0 Then col.Add cur: cur = ""
        Else
            cur = cur & ch
        End If
    Next
    If Len(cur) > 0 Then col.Add cur
    Set Tokenize = col
End Function

Public Sub RunCli()
    On Error GoTo fatal
    Dim toks As Collection
    Set toks = Tokenize(Command$)

    Dim inputPath As String, outDir As String
    Dim wantVbp As Boolean, wantDism As Boolean, wantSln As Boolean
    Dim i As Long
    i = 1
    Do While i <= toks.Count
        Dim t As String
        t = toks(i)
        If Left$(t, 1) = "/" Then
            Select Case LCase$(t)
                Case "/out"
                    If i < toks.Count Then i = i + 1: outDir = toks(i)
                Case "/vbp": wantVbp = True
                Case "/dism": wantDism = True
                Case "/solution", "/sln": wantSln = True
                Case "/?", "/help", "/h"
                    Call FinishCli(0, "Usage: SemiVBDecompiler.exe <input> [/out <dir>] [/vbp] [/dism] [/solution]", outDir)
            End Select
        ElseIf Len(inputPath) = 0 Then
            inputPath = t
        End If
        i = i + 1
    Loop

    If Len(inputPath) = 0 Then
        Call FinishCli(2, "No input file specified.", outDir)
    End If
    If Len(Dir$(inputPath)) = 0 Then
        Call FinishCli(2, "Input file not found: " & inputPath, outDir)
    End If

    'Everything from here runs silently.
    gQuietMode = True

    'File title = name with extension.
    Dim title As String, p As Long
    title = inputPath
    p = InStrRev(title, "\")
    If p > 0 Then title = Mid$(title, p + 1)

    'Decompile / analyze.  The form loads hidden; this populates dump\<name>\.
    Call frmMain.OpenVBExe(inputPath, title)

    Dim msg As String
    msg = "Processed " & inputPath
    msg = msg & IIf(bISVBNET, " (.NET)", " (VB" & VBVersion & ")")

    Dim wantGen As Boolean
    wantGen = wantVbp Or wantDism Or wantSln

    If wantGen Then
        Dim target As String
        target = outDir
        If Len(target) = 0 Then target = App.Path & "\dump\" & SFile
        Call EnsureCliDir(target)
        If bISVBNET = True Then
            Call modVBNET.BuildDotNetSolution(target)
            msg = msg & "; built .NET solution in " & target
        Else
            If wantDism Then gExportDisassembly = True
            Call frmMain.GenerateProject(target)
            gExportDisassembly = False
            msg = msg & "; generated VB project in " & target
            If wantDism Then msg = msg & " (disassembly)"
        End If
    Else
        msg = msg & "; decompiled output in " & App.Path & "\dump\" & SFile
    End If

    Call FinishCli(0, msg, outDir)
    Exit Sub
fatal:
    gExportDisassembly = False
    Call FinishCli(1, "Error " & err.Number & ": " & err.Description, outDir)
End Sub

Private Sub EnsureCliDir(ByVal path As String)
    On Error Resume Next
    If Len(Dir$(path, vbDirectory)) = 0 Then MkDir path
End Sub

'Write the result to decompile.log and terminate with the given exit code.
Private Sub FinishCli(ByVal code As Long, ByVal message As String, ByVal outDir As String)
    On Error Resume Next
    Dim logDir As String
    If Len(outDir) > 0 Then
        logDir = outDir
    ElseIf Len(SFile) > 0 Then
        logDir = App.Path & "\dump\" & SFile
    Else
        logDir = App.Path
    End If
    EnsureCliDir logDir
    Dim ff As Integer
    ff = FreeFile
    Open logDir & "\decompile.log" For Append As #ff
    Print #ff, Now & "  [exit " & code & "] " & message
    Close #ff
    ExitProcess code
End Sub
