const module_name   = "Edit"
const module_ver    = "1.0"
const module_title  = "Edit"

Sub Init
    'Here you can adjust your keys, but first remap your original keymap

    addMenuItem "Selection to &Right" , module_name , "SelectToRight" , "Ctrl+Alt+Right"
    addMenuItem "Selection to &Left"  , module_name , "SelectToLeft"  , "Ctrl+Alt+Left"
    addMenuItem "Select Line &Down"   , module_name , "SelectLine"    , "Ctrl+L"
    addMenuItem "Select Line &Up"     , module_name , "SelectLineUp"  , "Shift+Ctrl+L"

    addMenuItem "New Line &After"          , module_name , "CtrlEnter"                  , "Ctrl+Enter"
    addMenuItem "New Line Between Smth"    , module_name , "InsertLineBetween"          , "Shift+Enter"
    addMenuItem "New Line &Before Current" , module_name , "InsertNewLineBeforeCurrent" , "Shift+Ctrl+Enter"

    addMenuItem "Tab &Next"     , module_name , "NextTab" , "Ctrl+PgDn"
    addMenuItem "Tab &Previous" , module_name , "PrivTab" , "Ctrl+PgUp"

    addMenuItem "&1. Add '' Single Quotes To Selection"   , module_name , "AddSingleQuotesToSelection" , "Ctrl+'"
    addMenuItem "&2. Add """" Double Quotes To Selection" , module_name , "AddSlashesToSelectionn"     , "Shift+Ctrl+'"
    addMenuItem "&3. Add [] Brackets To Selection"        , module_name , "AddBracketsToSelection"     , "Ctrl+["
    addMenuItem "&4. Add {} Braces To Selection"          , module_name , "AddBracesToSelection"     , "Shift+Ctrl+["

    addMenuItem "&5. Add ( ) Round Brackets To Selection" , module_name , "AddBracsToSpSelection" , "Ctrl+9"
    addMenuItem "&6. Add () Round Brackets To Selection"  , module_name , "AddBracsToSelection"   , "Shift+Ctrl+9"

    addMenuItem "&7. Add `` Apostrophes To Selection"     , module_name , "AddApostrophesToSelection" , ""
    addMenuItem "&8. Add %% Procents To Selection"        , module_name , "AddProcentsToSelection"    , ""

    addMenuItem "Open &TODO.txt"         , module_name, "OpenFileBlank", "Shift+Ctrl+Alt+Space"
    addMenuItem "&Copy Current Full Path", module_name, "CopyPath", "Alt+C"

    addMenuItem "List Selected Items"   , module_name, "ListSelectedItems"   , "Ctrl+0"
    addMenuItem "List Selected Strings" , module_name, "ListSelectedStrings" , "Shift+Ctrl+0"

    addMenuItem "List Selected Items"   , module_name, "ListSelectedItemsToArr"    , "Ctrl+]"
    addMenuItem "List Selected Strings" , module_name, "ListSelectedStringsToSmth" , "Shift+Ctrl+]"

End Sub



' List of items
Sub ListSelectedItems
    Dim item, selTxt, objSelTxt, s
    Set obj = NewEditor()
    obj.assignActiveEditor()
    s = ""
    selTxt = obj.selText()
    arrLines = Split(selTxt, vbCrLf)
'     MsgBox TypeName(arrLines)
    If selTxt <> "" Then
        For Each item In arrLines
            If Trim(Item) <> "" Then
                s = s & Trim(item) & ", "
            End If
        Next
        s = "(" & Left(s, len(s)-2) & ")"
    
    End If
    obj.selText(s)
    setClipboardText(s)
End Sub


' List selected items wrap with strings
Sub ListSelectedStrings
	Dim item, s
    Set obj = NewEditor()
    obj.assignActiveEditor()
    s = ""
    selTxt = obj.selText()
    arrLines = Split(selTxt, vbCrLf)
    If selTxt <> "" Then
        For Each item In arrLines
            If Trim(Item) <> "" Then
                s = s & "'" & Trim(item) & "', "
            End If
        Next
        s = "(" & Left(s, len(s)-2) & ")"
    End If
    obj.selText(s)
    setClipboardText(s)
End Sub


' List of items
Sub ListSelectedItemsToArr
    Dim item, selTxt, objSelTxt, s
    Set obj = NewEditor()
    obj.assignActiveEditor()
    s = ""
    selTxt = obj.selText()
    arrLines = Split(selTxt, vbCrLf)
'     MsgBox TypeName(arrLines)
    If selTxt <> "" Then
        For Each item In arrLines
            If Trim(Item) <> "" Then
                s = s & """" & Trim(item) & """, "
            End If
        Next
        s = "[" & Left(s, len(s)-2) & "]"
    End If
    obj.selText(s)
    setClipboardText(s)
End Sub


' List selected items wrap with strings
Sub ListSelectedStringsToSmth
	Dim item, s
    Set obj = NewEditor()
    obj.assignActiveEditor()
    s = ""
    selTxt = obj.selText()
    arrLines = Split(selTxt, vbCrLf)
    If selTxt <> "" Then
        For Each item In arrLines
            If Trim(Item) <> "" Then
                s = s & """" & Trim(item) & """, "
            End If
        Next
        s = "{" & Left(s, len(s)-2) & "}"
    End If
    obj.selText(s)
    setClipboardText(s)
End Sub




' Copy current full path
Sub CopyPath
    Set activeEditor = newEditor()
    activeEditor.assignActiveEditor()
    curPath = activeEditor.fileName()
    setClipboardText(curPath)
    Set activeEditor = Nothing
End Sub

Sub OpenFileBlank
    Set wshShell = CreateObject( "WScript.Shell" )
    wshShell.Run "PSPad.exe TODO.txt" , False
    Set wshShell = Nothing
    Set activeEditor = Nothing
End Sub



Function AddSingleQuotesTo(ByVal strInput)
    Dim strOutput, blnBinary, numCount, strChr
    strOutput = ""
    blnBinary = False
    For numCount = 1 To Len(strInput)
        strChr = Mid(strInput, numCount, 1)
        strOutput = strOutput & strChr
    Next
    If blnBinary Then
       strOutput = ToHex(strInput)
    Else
       strOutput = "'" & strOutput & "'"
    End If
    AddSingleQuotesTo = strOutput
End Function

Sub AddSingleQuotesToSelection()
    Dim strInput
    With newEditor()
        .assignActiveEditor()
        strInput = AddSingleQuotesTo(.selText())
        .selText(strInput)
    End With
End Sub


Function AddBracketsTo(ByVal strInput)
    Dim strOutput, blnBinary, numCount, strChr
    strOutput = ""
    blnBinary = False
    For numCount = 1 To Len(strInput)
        strChr = Mid(strInput, numCount, 1)
        strOutput = strOutput & strChr
    Next
    If blnBinary Then
       strOutput = ToHex(strInput)
    Else
       strOutput = "[" & strOutput & "]"
    End If
    AddBracketsTo = strOutput
End Function

Sub AddBracketsToSelection()
    Dim strInput
    With newEditor()
        .assignActiveEditor()
        strInput = AddBracketsTo(.selText())
        .selText(strInput)
    End With
End Sub


Function AddBracesTo(ByVal strInput)
    Dim strOutput, blnBinary, numCount, strChr
    strOutput = ""
    blnBinary = False
    For numCount = 1 To Len(strInput)
        strChr = Mid(strInput, numCount, 1)
        strOutput = strOutput & strChr
    Next
    If blnBinary Then
       strOutput = ToHex(strInput)
    Else
       strOutput = "{" & strOutput & "}"
    End If
    AddBracesTo = strOutput
End Function

Sub AddBracesToSelection()
    Dim strInput
    With newEditor()
        .assignActiveEditor()
        strInput = AddBracesTo(.selText())
        .selText(strInput)
    End With
End Sub



Function AddBracsTo(ByVal strInput)
    Dim strOutput, blnBinary, numCount, strChr
    strOutput = ""
    blnBinary = False
    For numCount = 1 To Len(strInput)
        strChr = Mid(strInput, numCount, 1)
        strOutput = strOutput & strChr
    Next
    If blnBinary Then
       strOutput = ToHex(strInput)
    Else
       strOutput = "(" & strOutput & ")"
    End If
    AddBracsTo = strOutput
End Function

Sub AddBracsToSelection()
    Dim strInput
    With newEditor()
         .assignActiveEditor()
         strInput = AddBracsTo(.selText())
         .selText(strInput)
    End With
End Sub


Function AddBracsToSp(ByVal strInput)
    Dim strOutput, blnBinary, numCount, strChr
    strOutput = ""
    blnBinary = False
    For numCount = 1 To Len(strInput)
        strChr = Mid(strInput, numCount, 1)
        strOutput = strOutput & strChr
    Next
    If blnBinary Then
       strOutput = ToHex(strInput)
    Else
       strOutput = "( " & strOutput & " )"
    End If
    AddBracsToSp = strOutput
End Function

Sub AddBracsToSpSelection()
    Dim strInput
    With newEditor()
         .assignActiveEditor()
         strInput = AddBracsToSp(.selText())
         .selText(strInput)
    End With
End Sub


Function AddApostrophesTo(ByVal strInput)
    Dim strOutput, blnBinary, numCount, strChr
    strOutput = ""
    blnBinary = False
    For numCount = 1 To Len(strInput)
        strChr = Mid(strInput, numCount, 1)
        strOutput = strOutput & strChr
    Next
    If blnBinary Then
       strOutput = ToHex(strInput)
    Else
       strOutput = "`" & strOutput & "`"
    End If
    AddApostrophesTo = strOutput
End Function

Sub AddApostrophesToSelection()
    Dim strInput
    With newEditor()
         .assignActiveEditor()
         strInput = AddApostrophesTo(.selText())
         .selText(strInput)
    End With
End Sub

Function AddProcentsTo(ByVal strInput)
    Dim strOutput, blnBinary, numCount, strChr
    strOutput = ""
    blnBinary = False
    For numCount = 1 To Len(strInput)
        strChr = Mid(strInput, numCount, 1)
        strOutput = strOutput & strChr
    Next
    If blnBinary Then
       strOutput = ToHex(strInput)
    Else
       strOutput = "%" & strOutput & "%"
    End If
    AddProcentsTo = strOutput
End Function

Sub AddProcentsToSelection()
    Dim strInput
    With newEditor()
         .assignActiveEditor()
         strInput = AddProcentsTo(.selText())
         .selText(strInput)
    End With
End Sub


Function AddSlashesToo(ByVal strInput)
    Dim strOutput, blnBinary, numCount, strChr
    strOutput = ""
    blnBinary = False
    For numCount = 1 To Len(strInput)
        strChr = Mid(strInput, numCount, 1)
        strOutput = strOutput & strChr
    Next
    If blnBinary Then
       strOutput = ToHex(strInput)
    Else
       strOutput = """" & strOutput & """"
    End If
    AddSlashesToo = strOutput
End Function

Sub AddSlashesToSelectionn()
    Dim strInput
    With newEditor()
        .assignActiveEditor()
        strInput = AddSlashesToo(.selText())
        .selText(strInput)
    End With
End Sub


Sub NextTab()
    Dim strInput
    With newEditor()
        .assignActiveEditor()
        runPSPadAction "aSelectNext"
    End With
End Sub

Sub PrivTab()
    Dim strInput
    With newEditor()
        .assignActiveEditor()
        runPSPadAction "aSelectPrew"
    End With
End Sub

Sub SelectToRight()
    Dim strInput
    With newEditor()
        .assignActiveEditor()
        .command("ecSelLineEnd")
    End With
End Sub

Sub SelectToLeft()
    Dim strInput
    With newEditor()
        .assignActiveEditor()
        .command("ecSelLineStart")
    End With
End Sub

Sub SelectLine()
    Dim strInput
    Set obj = NewEditor()
    obj.assignActiveEditor()
    If obj.selText() <> "" Then
        obj.command("ecSelDown")
    Else
        obj.command("ecNormalSelect")
        obj.command("ecPageLeft")
        obj.command("ecSelDown")
    End If
End Sub

Sub SelectLineUp()
    Dim strInput
    Set obj = NewEditor()
    obj.assignActiveEditor()
    If obj.selText() <> "" Then
        obj.command("ecSelUp")
    Else
        obj.command("ecNormalSelect")
        obj.command("ecPageRight")
        obj.command("ecRight")
        obj.command("ecSelUp")
    End If
End Sub

Sub CtrlEnter()
    Dim strInput
    With newEditor()
        .assignActiveEditor()
        .command("ecPageRight")
        .command("ecLineBreak")
    End With
End Sub

Sub InsertNewLineBeforeCurrent()
    Dim strInput
    With newEditor()
        .assignActiveEditor()
        .command("ecPageLeft")
        .command("ecUp")
        .command("ecPageRight")
        .command("ecLineBreak")
    End With
End Sub

Sub InsertLineBetween()
    Dim strInput
    With newEditor()
        .assignActiveEditor()
        .command("ecLineBreak")
        .command("ecUp")
        .command("ecPageRight")
        .command("ecLineBreak")
        .command("ecTab")
    End With
End Sub

