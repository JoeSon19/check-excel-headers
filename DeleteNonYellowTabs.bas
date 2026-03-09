Sub DeleteNonYellowTabs()

    Dim wsTabList As Worksheet
    Dim keepTabs As Object
    Dim lastRow As Long
    Dim i As Long
    Dim cellColor As Long
    Dim tabName As String
    Dim ws As Worksheet

    ' Yellow color constant (RGB 255,255,0 = 65535)
    Const YELLOW_COLOR As Long = 65535

    ' Use the active sheet as the tab list
    Set wsTabList = ActiveSheet

    ' Build dictionary of yellow-highlighted tab names to keep
    Set keepTabs = CreateObject("Scripting.Dictionary")
    keepTabs.CompareMode = vbTextCompare  ' case-insensitive

    lastRow = wsTabList.Cells(wsTabList.Rows.Count, "A").End(xlUp).Row

    For i = 1 To lastRow
        cellColor = wsTabList.Cells(i, "A").Interior.Color
        tabName = Trim(CStr(wsTabList.Cells(i, "A").Value))

        If tabName <> "" And cellColor = YELLOW_COLOR Then
            keepTabs(tabName) = True
        End If
    Next i

    ' Always keep the active sheet (tab list) itself
    keepTabs(wsTabList.Name) = True

    If keepTabs.Count <= 1 Then
        MsgBox "No yellow-highlighted tabs found in column A. Nothing to do.", vbExclamation
        Exit Sub
    End If

    ' Confirm before deleting
    Dim deleteCount As Long
    Dim deleteList As String
    deleteCount = 0
    deleteList = ""

    For Each ws In ActiveWorkbook.Sheets
        If Not keepTabs.Exists(ws.Name) Then
            deleteCount = deleteCount + 1
            deleteList = deleteList & vbCrLf & "  - " & ws.Name
        End If
    Next ws

    If deleteCount = 0 Then
        MsgBox "All tabs are yellow-highlighted. Nothing to delete.", vbInformation
        Exit Sub
    End If

    Dim answer As VbMsgBoxResult
    answer = MsgBox("This will DELETE " & deleteCount & " tab(s):" & vbCrLf & _
                     deleteList & vbCrLf & vbCrLf & _
                     "Keep " & (keepTabs.Count) & " yellow tab(s)." & vbCrLf & vbCrLf & _
                     "Continue?", vbYesNo + vbExclamation, "Confirm Delete")

    If answer <> vbYes Then
        MsgBox "Cancelled.", vbInformation
        Exit Sub
    End If

    ' Delete non-yellow tabs (loop backwards to avoid skipping sheets)
    Application.DisplayAlerts = False

    For i = ActiveWorkbook.Sheets.Count To 1 Step -1
        Set ws = ActiveWorkbook.Sheets(i)
        If Not keepTabs.Exists(ws.Name) Then
            ' Ensure at least 1 sheet remains
            If ActiveWorkbook.Sheets.Count > 1 Then
                ws.Delete
            End If
        End If
    Next i

    Application.DisplayAlerts = True

    MsgBox "Done! Deleted " & deleteCount & " tab(s).", vbInformation

End Sub
