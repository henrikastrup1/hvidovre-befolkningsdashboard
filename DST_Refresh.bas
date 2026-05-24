Attribute VB_Name = "DSTRefresh"
Option Explicit

' ═══════════════════════════════════════════════════════
' DST Data Refresh Macro for Hvidovre FOLK1AM Excel
' ═══════════════════════════════════════════════════════
' Sådan importeres:
'   Excel > Udvikler > Visual Basic > File > Import File
'   Vælg denne .bas-fil
'   Kør med Alt+F8 > RefreshDSTData
'
' Eller tilføj en formular-knap:
'   Udvikler > Indsæt > Knap (Formularstyring)
'   Tildel makroen "RefreshDSTData"
' ═══════════════════════════════════════════════════════

' Opdater FOLK1AM data fra Danmarks Statistik API
Sub RefreshDSTData()
    Dim url As String
    Dim csvText As String
    Dim lines As Variant
    Dim i As Long
    Dim ws As Worksheet
    Dim destRow As Long
    Dim destCol As Long
    Dim parts As Variant
    Dim ageCol As Long
    Dim tidCol As Long
    Dim indholdCol As Long
    Dim line As Variant
    Dim currentSheet As String
    
    ' Gem hvilket ark der er aktivt
    currentSheet = ActiveSheet.Name
    
    ' API URL for FOLK1AM - Hvidovre (OMRÅDE=167), KØN=TOT, ALLE aldre, ALLE perioder
    url = "https://api.statbank.dk/v1/data/FOLK1AM/CSV?OMR%C3%85DE=167&K%C3%98N=TOT&ALDER=*&TID=*&lang=da"
    
    ' Status
    Application.StatusBar = "Henter data fra Danmarks Statistik..."
    Application.Cursor = xlWait
    DoEvents
    
    ' Hent CSV data
    On Error GoTo ErrHandler
    csvText = GetCSV(url)
    
    If csvText = "" Then
        MsgBox "Kunne ikke hente data fra DST. Tjek din internetforbindelse.", vbExclamation, "Fejl"
        Gooto Cleanup
    End If
    
    ' Find eller opret "Data" arket
    On Error Resume Next
    Set ws = ThisWorkbook.Sheets("Data_Raw")
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Sheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        ws.Name = "Data_Raw"
    End If
    On Error GoTo ErrHandler
    
    ' Ryd eksisterende data
    ws.Cells.Clear
    
    ' Parse CSV
    lines = Split(csvText, vbCrLf)
    If UBound(lines) < 2 Then
        lines = Split(csvText, vbLf)
    End If
    
    ' Find kolonneindeks
    parts = Split(lines(0), ";")
    ageCol = -1
    tidCol = -1
    indholdCol = -1
    
    For i = 0 To UBound(parts)
        Dim p As String
        p = Trim(Replace(parts(i), Chr(34), ""))
        p = Replace(p, ChrW(197), "Å") ' Å
        p = Replace(p, ChrW(216), "Ø") ' Ø
        
        If p = "ALDER" Then ageCol = i
        If p = "TID" Then tidCol = i
        If p = "INDHOLD" Then indholdCol = i
    Next i
    
    If ageCol = -1 Or tidCol = -1 Or indholdCol = -1 Then
        MsgBox "Kunne ikke finde forventede kolonner i data. Tjek om API'et har ændret format.", vbExclamation, "Fejl"
        GoTo Cleanup
    End If
    
    ' Skriv headers
    ws.Cells(1, 1).Value = "ALDER"
    ws.Cells(1, 2).Value = "PERIODE"
    ws.Cells(1, 3).Value = "INDHOLD"
    ws.Cells(1, 1).Font.Bold = True
    ws.Cells(1, 2).Font.Bold = True
    ws.Cells(1, 3).Font.Bold = True
    
    ' Skriv data (skip header line og "Alder i alt" rækker)
    destRow = 2
    For i = 1 To UBound(lines)
        line = Trim(lines(i))
        If line = "" Then GoTo NextLine
        
        parts = Split(line, ";")
        If UBound(parts) < indholdCol Then GoTo NextLine
        
        ' Skip "Alder i alt" (total række)
        If InStr(1, parts(ageCol), "Alder i alt", vbTextCompare) > 0 Then GoTo NextLine
        ' Skip OMRÅDE og KØN kolonner
        If i = 0 Then GoTo NextLine
        
        ' Rens anførselstegn
        Dim alder As String, periode As String, vaerdi As String
        alder = Trim(Replace(parts(ageCol), Chr(34), ""))
        periode = Trim(Replace(parts(tidCol), Chr(34), ""))
        vaerdi = Trim(Replace(parts(indholdCol), Chr(34), ""))
        
        ' Konverter alder til tal hvis muligt
        Dim alderTal As Integer
        On Error Resume Next
        alderTal = CInt(Split(alder)(0))
        On Error GoTo ErrHandler
        
        ws.Cells(destRow, 1).Value = alder
        ws.Cells(destRow, 2).Value = periode
        ws.Cells(destRow, 3).Value = CLng(vaerdi)
        
        destRow = destRow + 1
NextLine:
    Next i
    
    ' Formatér
    ws.Columns("A:C").AutoFit
    ws.Cells.EntireColumn.AutoFit
    ws.Range("A1").CurrentRegion.Borders.LineStyle = xlContinuous
    
    Application.StatusBar = "Data opdateret: " & Now()
    MsgBox "Data opdateret med " & (destRow - 2) & " rækker fra Danmarks Statistik.", vbInformation, "Færdig"
    
Cleanup:
    Application.Cursor = xlDefault
    Application.StatusBar = False
    On Error GoTo 0
    Exit Sub
    
ErrHandler:
    MsgBox "Fejl: " & Err.Description, vbExclamation, "Fejl"
    Resume Cleanup
End Sub

' Hent CSV via HTTP
Function GetCSV(url As String) As String
    Dim xmlHttp As Object
    Set xmlHttp = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    
    xmlHttp.Open "GET", url, False
    xmlHttp.setRequestHeader "Accept", "text/csv"
    xmlHttp.send
    
    If xmlHttp.Status = 200 Then
        GetCSV = xmlHttp.responseText
    Else
        GetCSV = ""
    End If
    
    Set xmlHttp = Nothing
End Function
