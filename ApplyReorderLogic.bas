Attribute VB_Name = "Module1"
Sub MergeInventoryAndFormat()

    Dim folderPath As String
    Dim wb2023 As Workbook, wb2024 As Workbook, masterWB As Workbook
    Dim ws2023 As Worksheet, ws2024 As Worksheet, wsMaster As Worksheet
    Dim lastRow As Long, pasteRow As Long
    Dim wsReorder As Worksheet

    ' Set folder path where both files are located
    folderPath = "C:\Users\DELL\Documents\Excel Projects\"

    ' Open 2023 and 2024 workbooks
    Set wb2023 = Workbooks.Open(folderPath & "2023_Inventory_Stock_Reorder_System.xlsx")
    Set wb2024 = Workbooks.Open(folderPath & "2024_Inventory_Stock_Reorder_System.xlsx")

    ' Assume data is on Sheet1 of both
    Set ws2023 = wb2023.Sheets(1)
    Set ws2024 = wb2024.Sheets(1)

    ' Create a new sheet in the current workbook to combine both
    Set masterWB = ThisWorkbook
    On Error Resume Next
    Application.DisplayAlerts = False
    masterWB.Sheets("MergedData").Delete
    masterWB.Sheets("ReorderOnly").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    Set wsMaster = masterWB.Sheets.Add(After:=masterWB.Sheets(masterWB.Sheets.Count))
    wsMaster.Name = "MergedData"

    ' Copy headers
    ws2023.Rows(1).Copy wsMaster.Rows(1)
    pasteRow = 2

    ' Copy all data from 2023 (ignore filter)
    lastRow = ws2023.Cells(ws2023.Rows.Count, 1).End(xlUp).Row
    ws2023.Range("A2:K" & lastRow).Copy wsMaster.Range("A" & pasteRow)
    pasteRow = pasteRow + lastRow - 1

    ' Copy all data from 2024 (ignore filter)
    lastRow = ws2024.Cells(ws2024.Rows.Count, 1).End(xlUp).Row
    ws2024.Range("A2:K" & lastRow).Copy wsMaster.Range("A" & pasteRow)

    ' Close source workbooks
    wb2023.Close False
    wb2024.Close False

    ' Apply conditional formatting to Status column (assumed column E)
    With wsMaster.Range("E2:E" & wsMaster.Cells(wsMaster.Rows.Count, 5).End(xlUp).Row)
        .FormatConditions.Delete
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""OK"""
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(144, 238, 144)
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Reorder"""
        .FormatConditions(.FormatConditions.Count).Interior.Color = RGB(255, 99, 71)
    End With

    ' Create ReorderOnly sheet
    Set wsReorder = masterWB.Sheets.Add(After:=wsMaster)
    wsReorder.Name = "ReorderOnly"

    ' Filter Reorder status and copy visible rows
    With wsMaster
        .Range("A1").AutoFilter Field:=5, Criteria1:="Reorder"
        .Range("A1:K" & .Cells(.Rows.Count, 1).End(xlUp).Row).SpecialCells(xlCellTypeVisible).Copy Destination:=wsReorder.Range("A1")
        .AutoFilterMode = False
        
        ' Auto-fit columns for clean layout
        wsMaster.Columns.AutoFit


    End With

    MsgBox "Merge and Reorder extraction completed!", vbInformation

End Sub



Sub ApplyReorderLogic(ws As Worksheet)
    ' Clear existing filters
    On Error Resume Next
    ws.ShowAllData
    On Error GoTo 0

    ' Clear old conditional formats in column E
    With ws.Columns("E")
        .FormatConditions.Delete

        ' Green for OK
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""OK"""
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .Color = RGB(144, 238, 144) ' Light green
        End With

        ' Red for Reorder
        .FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""Reorder"""
        .FormatConditions(.FormatConditions.Count).SetFirstPriority
        With .FormatConditions(1).Interior
            .Color = RGB(255, 102, 102) ' Light red
        End With
    End With

    ' Apply filter to show only Reorder
    ws.Range("A1").AutoFilter Field:=5, Criteria1:="Reorder"
    wsReorder.Columns.AutoFit
End Sub



Sub ExportReorderSheetToPDF()
    Dim ws As Worksheet
    Dim filePath As String

    ' Set the worksheet to export
    Set ws = ThisWorkbook.Sheets("ReorderOnly") ' Change to your actual sheet name if different

    ' Define the output path
    filePath = ThisWorkbook.Path & "\Reorder_Report_" & Format(Now(), "yyyymmdd_hhmmss") & ".pdf"

    ' Export as PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, fileName:=filePath, Quality:=xlQualityStandard

    MsgBox "PDF exported successfully to: " & filePath
End Sub


Sub EmailReorderPDF()
    Dim outlookApp As Object
    Dim outlookMail As Object
    Dim folderPath As String
    Dim fileName As String
    Dim fileSystem As Object
    Dim folder As Object
    Dim file As Object
    Dim latestFile As String
    Dim latestDate As Date

    ' Set folder path (same folder as workbook)
    folderPath = ThisWorkbook.Path & "\"

    ' Find the most recent PDF file
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    Set folder = fileSystem.GetFolder(folderPath)
    latestDate = DateSerial(1900, 1, 1)

    For Each file In folder.Files
        If LCase(Right(file.Name, 4)) = ".pdf" Then
            If InStr(file.Name, "Reorder_Report_") > 0 Then
                If file.DateLastModified > latestDate Then
                    latestDate = file.DateLastModified
                    latestFile = file.Path
                End If
            End If
        End If
    Next file

    If latestFile = "" Then
        MsgBox "No Reorder PDF found!", vbExclamation
        Exit Sub
    End If

    ' Create and send Outlook email
    Set outlookApp = CreateObject("Outlook.Application")
    Set outlookMail = outlookApp.CreateItem(0)

    With outlookMail
        .To = "fijaytwo@gmail.com" '
        .CC = ""
        .BCC = ""
        .Subject = "Automated Reorder Report"
        .Body = "Hello," & vbCrLf & vbCrLf & _
                "Please find attached the latest Reorder Report generated automatically." & vbCrLf & vbCrLf & _
                "Best regards," & vbCrLf & "Inventory Bot"
        .Attachments.Add latestFile
        .Send
    End With

    MsgBox "Email sent successfully with the latest Reorder PDF!"
End Sub



