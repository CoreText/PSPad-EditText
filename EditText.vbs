const module_name   = "Edit"
const module_ver    = "1.0"
const module_title  = "Edit"

'Here you can adjust your keys, but first remap your original keymap
Sub Init

    addMenuItem "Selection to &Right" , module_name , "SelectToRight" , "Ctrl+Alt+Right"
    addMenuItem "Selection to &Left"  , module_name , "SelectToLeft"  , "Ctrl+Alt+Left"
    addMenuItem "Select Line &Down"   , module_name , "SelectLine"    , "Ctrl+L"
    addMenuItem "Select Line &Up"     , module_name , "SelectLineUp"  , "Shift+Ctrl+L"

    addMenuItem "New Line &After"          , module_name , "CtrlEnter"                  , "Ctrl+Enter"
    addMenuItem "New Line Between Smth"    , module_name , "InsertLineBetween"          , "Shift+Enter"
    addMenuItem "New Line &Before Current" , module_name , "InsertNewLineBeforeCurrent" , "Shift+Ctrl+Enter"

    addMenuItem "Tab &Next"     , module_name , "NextTab" , "Alt+Right"
    addMenuItem "Tab &Previous" , module_name , "PrivTab" , "Alt+Left"

    addMenuItem "&1. Add '' Single Quotes To Selection"   , module_name , "AddSingleQuotesToSelection" , "Ctrl+'"
    addMenuItem "&2. Add """" Double Quotes To Selection" , module_name , "AddSlashesToSelectionn"     , "Shift+Ctrl+'"
    addMenuItem "&3. Add [] Brackets To Selection"        , module_name , "AddBracketsToSelection"     , "Ctrl+["
    addMenuItem "&4. Add {} Braces To Selection"          , module_name , "AddBracesToSelection"       , "Shift+Ctrl+["

    addMenuItem "&5. Add ( ) Round Brackets To Selection" , module_name , "AddBracsToSelection" , "Ctrl+9"
    addMenuItem "&6. Add () Round Brackets To Selection"  , module_name , "AddBracsToSelection" , "Shift+Ctrl+9"

    addMenuItem "&7. Add `` Apostrophes To Selection"     , module_name , "AddApostrophesToSelection" , ""
    addMenuItem "&8. Add %% Procents To Selection"        , module_name , "AddProcentsToSelection"    , ""

    addMenuItem "Open &TODO.txt"          , module_name , "OpenFileBlank", "Shift+Ctrl+Alt+Space"
    addMenuItem "&Copy Current Full Path" , module_name , "CopyPath", "Alt+C"

    addMenuItem "List Selected Items"   , module_name , "ListSelectedItems"   , "Ctrl+0"
    addMenuItem "List Selected Strings" , module_name , "ListSelectedStrings" , "Shift+Ctrl+0"

    addMenuItem "List Selected Items"   , module_name , "ListSelectedItemsToArr"    , "Ctrl+]"
    addMenuItem "List Selected Strings" , module_name , "ListSelectedStringsToSmth" , "Shift+Ctrl+]"

    addMenuItem "Focus Move" , module_name, "FocusMove" , "Alt+D"
'     addMenuItem "SelectWord" , module_name, "SelectWord" , "Ctrl+W"
End Sub

' under construction
Sub SelectWord
    Dim line, leng, curx, cury, cursmb, i, begPos, endPos
    Set ed = NewEditor()
    ed.assignActiveEditor()

    line = ed.lineText()
    curx = ed.caretX()
    cury = ed.caretY()
    leng = Len(line)

    i = curx - 1
'     While Mid( line, i, leng ) And i < leng
'     	i = i + 1
'     Wend

'     endPos = i - 1
'      i = curx - 1

    MsgBox TypeName(ed.caretY)
    MsgBox ed.caretY

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
        s = "( " & Left(s, len(s)-2) & " )"
    Else
        runPSPadAction "aFindWord"
        s = obj.selText()
        s = "( " & s & " )"
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
    Else
        runPSPadAction "aFindWord"
        s = obj.selText()
        s = "(" & s & ")"
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
    Else
        runPSPadAction "aFindWord"
        s = obj.selText()
        s = "[" & s & "]"
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
    Else
        runPSPadAction "aFindWord"
        s = obj.selText()
        s = "{" & s & "}"
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
    With newEditor()
        .assignActiveEditor()
        runPSPadAction "aSelectNext"
    End With
End Sub

Sub PrivTab()
    With newEditor()
        .assignActiveEditor()
        runPSPadAction "aSelectPrew"
    End With
End Sub

Sub SelectToRight()
    With newEditor()
        .assignActiveEditor()
        .command("ecSelLineEnd")
    End With
End Sub

Sub SelectToLeft()
    With newEditor()
        .assignActiveEditor()
        .command("ecSelLineStart")
    End With
End Sub

Sub SelectLine()
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
    With newEditor()
        .assignActiveEditor()
        .command("ecPageRight")
        .command("ecLineBreak")
    End With
End Sub

Sub InsertNewLineBeforeCurrent()
    Set obj = newEditor()
        obj.assignActiveEditor()

    If obj.caretY < 2 Then
        obj.command("ecInsertLine")
    Else
        obj.command("ecPageLeft")
        obj.command("ecUp")
        obj.command("ecPageRight")
        obj.command("ecLineBreak")

    End If
End Sub

Sub InsertLineBetween()
    Set obj = newEditor()
        obj.assignActiveEditor()
   
    If obj.selText() <> "" Then
        obj.command("ecCut")
        obj.command("ecLineBreak")
        obj.command("ecUp")
        obj.command("ecPageRight")
        obj.command("ecLineBreak")
        obj.command("ecTab")
        obj.command("ecPaste")
    Else
        obj.command("ecLineBreak")
        obj.command("ecUp")
        obj.command("ecPageRight")
        obj.command("ecLineBreak")
        obj.command("ecTab")
    End If
End Sub

Sub FocusMove
    With NewEditor()
        .assignActiveEditor()
        runPSPadAction "aSwitchLog"
    End with
End Sub
