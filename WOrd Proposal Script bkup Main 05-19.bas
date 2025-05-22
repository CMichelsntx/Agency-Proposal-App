Attribute VB_Name = "Module1"

Option Explicit
 
Sub Format_All_Tables_With_Special_Sections()
    Dim doc As Document
    Set doc = ActiveDocument
    Dim tbl As Table
    Dim cellText As String
    Dim foundCE_Table As Table
    Dim i As Long, totalCols As Long
    Dim shp As Shape
       
    '-----------------------------------------------------------
    ' 1) PROCESS ALL TABLES
    '-----------------------------------------------------------
    For Each tbl In doc.Tables
        tbl.AutoFitBehavior wdAutoFitContent
        tbl.Rows.Alignment = wdAlignRowLeft
        tbl.Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
           
        ' Set font size to 10 for all data in the table
        tbl.Range.Font.Size = 9
    Next tbl
       
    '-----------------------------------------------------------
    ' 2) LOCATE "COVERED ENTITY" TABLE
    '-----------------------------------------------------------
    For Each tbl In doc.Tables
        cellText = CleanCellText(tbl.cell(1, 1).Range.Text)
        If UCase(cellText) = "COVERED ENTITY" Then
            Set foundCE_Table = tbl
            Exit For
        End If
    Next tbl
       
    If Not foundCE_Table Is Nothing Then
        foundCE_Table.AutoFitBehavior wdAutoFitWindow
        foundCE_Table.Rows.Alignment = wdAlignRowLeft
        Dim headerMapping As Object
        Set headerMapping = CreateObject("Scripting.Dictionary")
        headerMapping.Add "PROPERTY", "PROPERTY"
        headerMapping.Add "INLAND MARINE", "INLAND" & vbCrLf & "MARINE"
        headerMapping.Add "GENERAL LIABILITY", "GENERAL" & vbCrLf & "LIABILITY"
        headerMapping.Add "COMMERCIAL AUTO", "COMMERCIAL" & vbCrLf & "AUTO"
        headerMapping.Add "WORKERS COMPENSATION", "WORKERS" & vbCrLf & "COMPENSATION"
        headerMapping.Add "UMBRELLA", "UMBRELLA"
        headerMapping.Add "CYBER", "CYBER"
        headerMapping.Add "DIRECTORS & OFFICERS", "DIRECTORS &" & vbCrLf & "OFFICERS"
        headerMapping.Add "EMPLOYMENT PRACTICES", "EMPLOYMENT" & vbCrLf & "PRACTICES"
        headerMapping.Add "CRIME", "CRIME"
        headerMapping.Add "FIDUCIARY LIABILITY", "FIDUCIARY" & vbCrLf & "LIABILITY"
        totalCols = foundCE_Table.Columns.Count
        Dim colIndex As Long
        For colIndex = 2 To totalCols
            cellText = CleanCellText(foundCE_Table.cell(1, colIndex).Range.Text)
            If headerMapping.Exists(UCase(cellText)) Then
                foundCE_Table.cell(1, colIndex).Range.Text = headerMapping(UCase(cellText))
            End If
            foundCE_Table.cell(1, colIndex).Range.Orientation = wdTextOrientationUpward
            foundCE_Table.cell(1, colIndex).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
        Next colIndex
        foundCE_Table.cell(1, 1).Range.Orientation = wdTextOrientationHorizontal
        foundCE_Table.cell(1, 1).Range.ParagraphFormat.Alignment = wdAlignParagraphLeft
        With foundCE_Table.Rows(1)
            .HeightRule = wdRowHeightAtLeast
            .Height = doc.Application.LinesToPoints(7) + 20
        End With
    End If
       
    '-----------------------------------------------------------
    ' 3) REPLACE BOOKMARK MARKERS ON PAGE 1
    '-----------------------------------------------------------
    Dim rTable As Table, rRow As Row
    Dim strNamedInsured As String, strAgentName As String, strDate As String
    For Each rTable In doc.Tables
        For Each rRow In rTable.Rows
            If UCase(Trim(CleanCellText(rRow.Cells(1).Range.Text))) = "NAMED INSURED" Then
                strNamedInsured = Trim(CleanCellText(rRow.Cells(2).Range.Text))
            ElseIf UCase(Trim(CleanCellText(rRow.Cells(1).Range.Text))) = "AGENT NAME" Then
                strAgentName = Trim(CleanCellText(rRow.Cells(2).Range.Text))
            End If
        Next rRow
    Next rTable
    strDate = Format(Date, "mm/dd/yyyy")
    Dim rngFirst As Range
    
    ' --- {Named Insured} ---
    Set rngFirst = doc.Bookmarks("\Page").Range
    With rngFirst.Find
        .ClearFormatting
        .Text = "{Named Insured}"
        .Replacement.ClearFormatting
        If .Execute Then
            Dim markerFontName As String, markerFontSize As Single, markerFontColor As Long
            markerFontName = rngFirst.Font.Name
            markerFontSize = rngFirst.Font.Size
            markerFontColor = rngFirst.Font.Color
            rngFirst.Text = strNamedInsured
            With rngFirst.Font
                .Name = markerFontName
                .Size = markerFontSize
                .Color = markerFontColor
            End With
        End If
    End With
    
    ' --- {Agent Name} ---
    Set rngFirst = doc.Bookmarks("\Page").Range
    With rngFirst.Find
        .ClearFormatting
        .Text = "{Agent Name}"
        .Replacement.ClearFormatting
        If .Execute Then
            Dim markerFontName2 As String, markerFontSize2 As Single, markerFontColor2 As Long
            markerFontName2 = rngFirst.Font.Name
            markerFontSize2 = rngFirst.Font.Size
            markerFontColor2 = rngFirst.Font.Color
            rngFirst.Text = strAgentName
            With rngFirst.Font
                .Name = markerFontName2
                .Size = markerFontSize2
                .Color = markerFontColor2
            End With
            ' Align AGENT NAME to center (changed from left)
            rngFirst.ParagraphFormat.Alignment = wdAlignParagraphCenter
        End If
    End With
    
    ' --- {Date} ---
    Set rngFirst = doc.Bookmarks("\Page").Range
    With rngFirst.Find
        .ClearFormatting
        .Text = "{Date}"
        .Replacement.ClearFormatting
        If .Execute Then
            Dim markerFontName3 As String, markerFontSize3 As Single, markerFontColor3 As Long
            markerFontName3 = rngFirst.Font.Name
            markerFontSize3 = rngFirst.Font.Size
            markerFontColor3 = rngFirst.Font.Color
            rngFirst.Text = strDate
            With rngFirst.Font
                .Name = markerFontName3
                .Size = markerFontSize3
                .Color = markerFontColor3
            End With
        End If
    End With
       
    '-----------------------------------------------------------
    ' 4) DELETE EMPTY TITLE PAGES IF NO TABLES PRESENT
    '-----------------------------------------------------------
    Dim coverageTitles As Variant
    coverageTitles = Array("Policy", "Commercial Property", "General Liability", "Auto", "Inland Marine", "Umbrella", "Workers Compensation", "Workers' Compensation")
    Dim iTitle As Long, pageNum As Long
    Dim para As Paragraph, paraText As String
    For Each para In doc.Paragraphs
        paraText = NormalizeText(para.Range.Text)
        For iTitle = LBound(coverageTitles) To UBound(coverageTitles)
            If Trim(UCase(paraText)) = Trim(UCase(coverageTitles(iTitle))) Then
                pageNum = para.Range.Information(wdActiveEndPageNumber)
                GoSub DeletePageIfNoTable
                Exit For
            End If
        Next iTitle
    Next para
    For Each shp In doc.Shapes
        If shp.Type = msoTextBox And shp.TextFrame.HasText Then
            paraText = NormalizeText(shp.TextFrame.TextRange.Text)
            For iTitle = LBound(coverageTitles) To UBound(coverageTitles)
                If Trim(UCase(paraText)) = Trim(UCase(coverageTitles(iTitle))) Then
                    pageNum = shp.Anchor.Information(wdActiveEndPageNumber)
                    GoSub DeletePageIfNoTable
                    Exit For
                End If
            Next iTitle
        End If
    Next shp
    MsgBox "Tower Street Proposal completed!!", vbInformation
    Exit Sub
 
'-----------------------------------------------------------
' Subroutine: Delete page if no table or linked content
'-----------------------------------------------------------
DeletePageIfNoTable:
    On Error Resume Next
    Dim pgStart As Long, pgEnd As Long
    Dim docRangeStart As Range, docRangeEnd As Range
    Set docRangeStart = ActiveDocument.GoTo(What:=wdGoToPage, Name:=CStr(pageNum))
    pgStart = docRangeStart.Start
    If pageNum < ActiveDocument.ComputeStatistics(wdStatisticPages) Then
        Set docRangeEnd = ActiveDocument.GoTo(What:=wdGoToPage, Name:=CStr(pageNum + 1))
        pgEnd = docRangeEnd.Start - 1
    Else
        pgEnd = ActiveDocument.Content.End
    End If
    Dim pgRange As Range
    Set pgRange = ActiveDocument.Range(Start:=pgStart, End:=pgEnd)
    Dim hasTableOnPage As Boolean: hasTableOnPage = False
    Dim hasTitleTextBox As Boolean: hasTitleTextBox = False
    Dim shapeText As String, matchIndex As Long
    
    For Each tbl In ActiveDocument.Tables
        If tbl.Range.Start >= pgStart And tbl.Range.Start <= pgEnd Then
            hasTableOnPage = True
            Exit For
        End If
    Next tbl
    
    For Each shp In ActiveDocument.Shapes
        If shp.Type = msoTextBox And shp.TextFrame.HasText Then
            shapeText = NormalizeText(shp.TextFrame.TextRange.Text)
            For matchIndex = LBound(coverageTitles) To UBound(coverageTitles)
                If Trim(UCase(shapeText)) = Trim(UCase(coverageTitles(matchIndex))) Then
                    If Not shp.Anchor Is Nothing Then
                        If shp.Anchor.Start >= pgStart And shp.Anchor.Start <= pgEnd Then
                            hasTitleTextBox = True
                            Exit For
                        End If
                    End If
                End If
            Next matchIndex
        End If
        If hasTitleTextBox Then Exit For
    Next shp
    
    ' Check next page for table if title present
    If Not hasTableOnPage And hasTitleTextBox Then
        If pageNum < ActiveDocument.ComputeStatistics(wdStatisticPages) Then
            Dim nextStart As Long, nextEnd As Long
            nextStart = ActiveDocument.GoTo(What:=wdGoToPage, Name:=CStr(pageNum + 1)).Start
            If pageNum + 1 < ActiveDocument.ComputeStatistics(wdStatisticPages) Then
                nextEnd = ActiveDocument.GoTo(What:=wdGoToPage, Name:=CStr(pageNum + 2)).Start - 1
            Else
                nextEnd = ActiveDocument.Content.End
            End If
            Dim nextRange As Range
            Set nextRange = ActiveDocument.Range(Start:=nextStart, End:=nextEnd)
            For Each tbl In ActiveDocument.Tables
                If tbl.Range.Start >= nextRange.Start And tbl.Range.Start <= nextRange.End Then
                    hasTableOnPage = True
                    Exit For
                End If
            Next tbl
        End If
    End If
    
    If Not hasTableOnPage Then
        pgRange.Delete
        Dim breakRange As Range
        Set breakRange = ActiveDocument.Range(pgStart, pgStart + 10)
        With breakRange.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "^m"
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll
        End With
        With breakRange.Find
            .ClearFormatting
            .Replacement.ClearFormatting
            .Text = "^b"
            .Replacement.Text = ""
            .Execute Replace:=wdReplaceAll
        End With
        Dim sec As Section, nextSec As Section, sIndex As Long
        Set sec = pgRange.Sections(1)
        sIndex = sec.Index
        If sIndex > 1 And sIndex < ActiveDocument.Sections.Count Then
            Set nextSec = ActiveDocument.Sections(sIndex + 1)
            If Not nextSec Is Nothing Then
                nextSec.Footers(wdHeaderFooterPrimary).LinkToPrevious = True
                nextSec.Footers(wdHeaderFooterFirstPage).LinkToPrevious = True
                nextSec.Footers(wdHeaderFooterEvenPages).LinkToPrevious = True
            End If
        End If
    End If
Return
End Sub
 
Private Function NormalizeText(rawText As String) As String
    rawText = Replace(rawText, "'", "'")
    rawText = Replace(rawText, "'", "'")
    rawText = Replace(rawText, ChrW(147), """")
    rawText = Replace(rawText, ChrW(148), """")
    rawText = Replace(rawText, vbCr, "")
    rawText = Replace(rawText, vbLf, "")
    rawText = Replace(rawText, Chr(160), " ")
    NormalizeText = Trim(rawText)
End Function
 
Private Function CleanCellText(ByVal rawText As String) As String
    CleanCellText = Trim(Replace(rawText, Chr(13) & Chr(7), ""))
End Function

