Attribute VB_Name = "MainModule"
Option Explicit

Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)

'param0: "w" (writeprop) / "m" (runmacro) / "d" (deleteProp)
'param1: filename
'param2: propertyname or macroname
'param3: "y" (yes) / "n" (no) (iscustomproperty)
'param4: propertyvalue
'
'errorlevels:
'9901 ... wrong method code (must be "w" or "m")
'9902 ... no filename specified
'9903 ... no property or macro name specified
'9904 ... parameter iscustomproperty is invalid (must be "y" or "n")
'9905 ... no property value specified (not used any more)
'9906 ... wrong no of parameters
Public Sub Main()

    Dim oVAIAutomation As New MAINTECSAP2Word.SAP2Word
    Dim iResult As Long
    Dim args() As String
    Dim argsTemp() As String
    Dim sFileName, sPropertyName, sPropertyValue, sMacroName As String
    Dim bIsCustomProperty As Boolean
    Dim sTemp As String
    Dim sChar As String

    On Error GoTo Error_Main

    Dim bOpenQuotes, bOpenSpace As Boolean
    Dim iCurrentParam As Integer
    Dim sCurrentParam As String
    Dim i As Integer

    sTemp = Trim(Command$)
    bOpenQuotes = False
    bOpenSpace = False
    iCurrentParam = 0
    For i = 1 To Len(sTemp)
        sChar = Mid(sTemp, i, 1)

        If sChar = """" Then
            If Not bOpenQuotes Then
                'start new param with quote
                bOpenQuotes = True
            Else
                'end param with quotes -> write param
                '4th parameter (property value) -> take the rest of all
                If iCurrentParam < 4 Or i >= Len(sTemp) Then
                    ReDim Preserve args(iCurrentParam)
                    bOpenQuotes = False
                    args(iCurrentParam) = sCurrentParam
                    iCurrentParam = iCurrentParam + 1
                    sCurrentParam = ""
                    'skip space between params
                    If i < Len(sTemp) Then
                        i = i + 1
                    End If
                Else
                    sCurrentParam = sCurrentParam & sChar
                End If
            End If
        Else
            If sChar = " " Then
                If bOpenQuotes Then
                    'start new param with space (except beeing within quotes)
                    sCurrentParam = sCurrentParam & sChar
                Else
                    'end param with space -> write param
                    ReDim Preserve args(iCurrentParam)
                    args(iCurrentParam) = sCurrentParam
                    iCurrentParam = iCurrentParam + 1
                    sCurrentParam = ""
                End If
            Else
                sCurrentParam = sCurrentParam & sChar
                If i >= Len(sTemp) Then
                    ReDim Preserve args(iCurrentParam)
                    args(iCurrentParam) = sCurrentParam
                End If
            End If
        End If

    Next

    'sTemp = Replace(Command$, """", "##")
    'sTemp = Replace(sTemp, "#### ", "##<NULL>##")

    'sTemp = Replace(sTemp, "##", """")
    'sTemp = Replace(sTemp, """", "+++++")
    'argsTemp = Split(sTemp, "+++")

    'For i = 0 To UBound(argsTemp)
    '    If InStr(argsTemp(i), "++ ") > 0 Then
    '        argsTemp(i) = Replace(argsTemp(i), "++ ", "")
    '    End If
    '    If InStr(argsTemp(i), "++") > 0 Then
    '        argsTemp(i) = Replace(argsTemp(i), " ", "###")
    '        argsTemp(i) = Replace(argsTemp(i), "++", "")
    '    End If
    '    argsTemp(i) = Trim(argsTemp(i))
    'Next
    'sTemp = Trim(Join(argsTemp))
    'sTemp = Replace(sTemp, "  ", " ")

    ''join last parameters
    'argsTemp = Split(sTemp, " ")

    'If UBound(argsTemp) > 4 Then
    '    ReDim args(4)
    '    For i = 0 To 3
    '        args(i) = argsTemp(i)
    '    Next
    '    For i = 4 To UBound(argsTemp)
    '        args(4) = args(4) & argsTemp(i)
    '    Next
    'Else
    '    args = Split(sTemp, " ")
    'End If


    'For i = 0 To UBound(args)
    '    args(i) = Trim(Replace(args(i), "###", " "))
    '    args(i) = Trim(Replace(args(i), "<NULL>", ""))
    'Next

    'check parameters
    iResult = 0
    If Len(Command$) = 0 Then iResult = 9906
    If iResult = 0 Then
        If Not (LCase(args(0)) = "w" Or LCase(args(0)) = "m" Or LCase(args(0)) = "d") Then iResult = 9901
        If LCase(args(0)) = "w" And UBound(args) <> 4 Then iResult = 9906
        If LCase(args(0)) = "m" And UBound(args) <> 2 Then iResult = 9906
        If LCase(args(0)) = "d" And UBound(args) <> 2 Then iResult = 9906
    End If
    If iResult = 0 Then
        If Len(args(1)) = 0 Then iResult = 9902
        If Len(args(2)) = 0 Then iResult = 9903
        If LCase(args(0)) = "w" Then
            If Not (LCase(args(3)) = "y" Or LCase(args(3)) = "n") Then iResult = 9904
        End If
    End If

    'if parameters ok -> call dll
    If iResult = 0 Then
        Select Case LCase(args(0))

            Case "w"
                sFileName = args(1)
                sPropertyName = CStr(args(2))
                bIsCustomProperty = (LCase(args(3)) = "y")
                sPropertyValue = CStr(args(4))
                iResult = oVAIAutomation.WriteProp(sFileName, sPropertyName, bIsCustomProperty, sPropertyValue)

            Case "m"
                sFileName = args(1)
                sMacroName = CStr(args(2))
                iResult = oVAIAutomation.RunMacro(sFileName, sMacroName)
            Case "d"
                sFileName = args(1)
                sMacroName = CStr(args(2))
                iResult = oVAIAutomation.deleteProp(sFileName, sMacroName)
        End Select
    End If

    oVAIAutomation = Nothing

    'exit and return result
    If Compiled Then
        Call ExitProcess(iResult)
    End If


    Exit Sub

Error_Main:

    'exit and return result
    iResult = Err.Number
    oVAIAutomation = Nothing
    If Compiled Then
        Call ExitProcess(iResult)
    End If

End Sub


Private Function Compiled() As Boolean
    ' Determine if running from EXE/IDE.
    On Error Resume Next
    Debug.Print(1 / 0)
    Compiled = (Err.Number = 0)
End Function


