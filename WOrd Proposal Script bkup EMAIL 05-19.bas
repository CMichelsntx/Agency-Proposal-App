Attribute VB_Name = "Module2"
Option Explicit

Sub SendProposalEmail()
    Dim doc As Document
    Dim infoTbl As Table
    Dim covTbl As Table
    Dim tbl As Table
    Dim insured As String
    Dim policyPeriod As String
    Dim dedAOP As String
    Dim dedPremOps As String
    Dim dedProdComp As String
    Dim dedAutoColl As String
    Dim dedWHDed As String
    Dim umbrellaAgg As String
    Dim introText As String
    Dim termsText As String
    Dim oOutlook As Object
    Dim oMail As Object
    Dim wdEditor As Object
    Dim insertRange As Range
    Dim para As Paragraph
    Dim r As Long
    Dim tempPdfPath As String

    Set doc = ActiveDocument

    ' 1) Find the Policy Information table (optional) without error if missing
    Set infoTbl = Nothing
    For Each tbl In doc.Tables
        On Error Resume Next
        If LCase(Clean(tbl.cell(1, 1).Range.Text)) = "field" Then
            If LCase(Clean(tbl.cell(1, 2).Range.Text)) = "value" Then
                Set infoTbl = tbl
                On Error GoTo 0
                Exit For
            End If
        End If
        On Error GoTo 0
    Next tbl

    ' 2) Read Named Insured & Proposed Policy Period (if table found)
    If Not infoTbl Is Nothing Then
        For r = 2 To infoTbl.Rows.Count
            Select Case LCase(Clean(infoTbl.cell(r, 1).Range.Text))
                Case "named insured"
                    insured = Clean(infoTbl.cell(r, 2).Range.Text)
                Case "proposed policy period"
                    policyPeriod = Clean(infoTbl.cell(r, 2).Range.Text)
            End Select
        Next r
    End If

    ' 3) Export the open document to PDF named "<Insured> Proposal.pdf"
    tempPdfPath = Environ("TEMP") & "\" & insured & " Proposal.pdf"
    doc.ExportAsFixedFormat _
        OutputFileName:=tempPdfPath, _
        ExportFormat:=wdExportFormatPDF, _
        OpenAfterExport:=False, _
        OptimizeFor:=wdExportOptimizeForPrint, _
        Range:=wdExportAllDocument

    ' 4) Find the Policy Coverages table
    Set covTbl = Nothing
    For Each tbl In doc.Tables
        On Error Resume Next
        If LCase(Clean(tbl.cell(1, 1).Range.Text)) = "coverage" Then
            If LCase(Clean(tbl.cell(1, 2).Range.Text)) = "premium" Then
                Set covTbl = tbl
                On Error GoTo 0
                Exit For
            End If
        End If
        On Error GoTo 0
    Next tbl

    ' 5) Pull Deductibles & Aggregates
    dedAOP = GetFirstDedAfterHeading("Location Coverages", "Ded")
    dedPremOps = GetValueByFuzzyMatch("Prem/Ops")
    dedProdComp = GetExactValue("Prod/Comp Ops")
    dedAutoColl = GetUniqueAfterHeading("Auto Coverage Summary", "Comp Ded")
    dedWHDed = GetUniqueAfterHeading("Location Coverages", "W/H Ded")
    umbrellaAgg = GetFirstDedAfterHeading("Umbrella Limits of Insurance", "Limits")

    ' 6) Build the intro section
    introText = _
        "Thank you for the opportunity to quote one of your preferred accounts. " & _
        "We are pleased to present our proposal, which you will see attached to this email. " & _
        "The terms and conditions, enhancements and estimated premiums are outlined below." & vbCrLf & vbCrLf & _
        "Name: " & insured & vbCrLf & _
        "Effective Date: " & policyPeriod & vbCrLf & vbCrLf & _
        "Binding Subjectivities:" & vbCrLf & vbCrLf & _
        "• If bound, please send insured’s name, number & email for loss control ordering" & vbCrLf & _
        "• Signed Acord Application" & vbCrLf & _
        "• Signed Terrorism Selection or Rejection Form" & vbCrLf & _
        "• Confirm Pay Plan" & vbCrLf & _
        "• Acceptable MVRs – If bound, MVRs will be run prior to issuance." & vbCrLf & _
        "• Updated and Completed Driver’s List" & vbCrLf & _
        "• Acceptable Loss Control Survey – we will order if bound" & vbCrLf & vbCrLf

    ' 7) Build the Terms & Conditions section
    termsText = _
        "Terms and Conditions:" & vbCrLf & vbCrLf & _
        "Please note, the proposal includes underwriting requirements that may differentiate from the original application. Review the policy coverages closely." & vbCrLf & vbCrLf & _
        "Property:" & vbCrLf & _
        "• AOP Deductible = " & dedAOP & vbCrLf & _
        "• Wind/Hail Deductible = " & dedWHDed & vbCrLf & vbCrLf & _
        "General Liability:" & vbCrLf & _
        "• Prem/Ops Deductible = " & dedPremOps & vbCrLf & _
        "• Prod/Comp Ops      = " & dedProdComp & vbCrLf & vbCrLf & _
        "Auto:" & vbCrLf & _
        "• Auto Comp/Coll Deductible = " & dedAutoColl & vbCrLf & vbCrLf & _
        "Umbrella:" & vbCrLf & _
        "• General Aggregate = " & umbrellaAgg & vbCrLf & vbCrLf & _
        "The attached proposal outlined above is valid for 30 days. Coverage cannot be bound until written bind request has been accepted and cannot be backdated." & vbCrLf & vbCrLf & _
        "Please let me know if you have any questions or revisions that would help us secure the account." & vbCrLf

    ' 8) Create and display the Outlook mail (with PDF attached)
    Set oOutlook = CreateObject("Outlook.Application")
    Set oMail = oOutlook.CreateItem(0)
    With oMail
        .Subject = "Proposal for " & insured
        .Attachments.Add tempPdfPath
        .Display
    End With

    ' 9) Insert content into the mail body
    Set wdEditor = oMail.GetInspector.WordEditor
    Set insertRange = wdEditor.Content
    insertRange.Collapse Direction:=wdCollapseStart

    ' 9A) Intro text
    insertRange.InsertAfter introText
    insertRange.Collapse Direction:=wdCollapseEnd

    ' 9B) Coverage table (if found)
    If Not covTbl Is Nothing Then
        covTbl.Range.Copy
        insertRange.Paste
        With wdEditor.Tables(wdEditor.Tables.Count)
            .Rows.LeftIndent = 0
            .Range.ParagraphFormat.LeftIndent = 0
        End With
        insertRange.Collapse Direction:=wdCollapseEnd
    End If

    ' 9C) Blank line between table and Terms
    insertRange.InsertAfter vbCrLf
    insertRange.Collapse Direction:=wdCollapseEnd

    ' 9D) Terms & Conditions
    insertRange.InsertAfter termsText

    ' 10) Apply final formatting
    For Each para In wdEditor.Paragraphs
        With para.Range
            ' Blue + bold for headings
            If .Text Like "Name:*" Or .Text Like "Effective Date:*" Or _
               .Text Like "Binding Subjectivities:*" Or .Text Like "Terms and Conditions:*" Or _
               .Text Like "Property:*" Or .Text Like "General Liability:*" Or _
               .Text Like "Auto:*" Or .Text Like "Umbrella:*" Then
                .Font.Bold = True
                .Font.Color = RGB(21, 93, 139)
            ' Red for intro bullets
            ElseIf Left(.Text, 1) = "•" And InStr(1, introText, .Text, vbTextCompare) > 0 Then
                .Font.Color = RGB(255, 0, 0)
            End If
        End With
    Next para

    ' 11) Clean up the temp PDF file
    On Error Resume Next
    Kill tempPdfPath
    On Error GoTo 0

    ' 12) Release object references
    Set infoTbl = Nothing
    Set covTbl = Nothing
    Set oMail = Nothing
    Set oOutlook = Nothing
End Sub

'————————————————————————————————————
Function GetFirstDedAfterHeading(heading As String, columnHeader As String) As String
    Dim rng As Range, tbl As Table, r As Long, c As Long
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = heading
        .MatchCase = False
        .MatchWholeWord = False
        If .Execute Then
            rng.Collapse Direction:=wdCollapseEnd
            Do While rng.Tables.Count = 0 And rng.End < ActiveDocument.Content.End
                rng.Move Unit:=wdParagraph, Count:=1
            Loop
            If rng.Tables.Count > 0 Then
                Set tbl = rng.Tables(1)
                For c = 1 To tbl.Columns.Count
                    If LCase(Clean(tbl.cell(1, c).Range.Text)) = LCase(columnHeader) Then Exit For
                Next c
                If c <= tbl.Columns.Count Then
                    GetFirstDedAfterHeading = Clean(tbl.cell(2, c).Range.Text)
                    Exit Function
                End If
            End If
        End If
    End With
    GetFirstDedAfterHeading = ""
End Function

'————————————————————————————————————
Function GetUniqueAfterHeading(heading As String, columnHeader As String) As String
    Dim rng As Range, tbl As Table, r As Long, c As Long
    Dim vals As New Collection, v As String
    On Error Resume Next
    Set rng = ActiveDocument.Content
    With rng.Find
        .Text = heading
        .MatchCase = False
        .MatchWholeWord = False
        If .Execute Then
            rng.Collapse Direction:=wdCollapseEnd
            Do While rng.Tables.Count = 0 And rng.End < ActiveDocument.Content.End
                rng.Move Unit:=wdParagraph, Count:=1
            Loop
            If rng.Tables.Count > 0 Then
                Set tbl = rng.Tables(1)
                For c = 1 To tbl.Columns.Count
                    If LCase(Clean(tbl.cell(1, c).Range.Text)) = LCase(columnHeader) Then Exit For
                Next c
                If c <= tbl.Columns.Count Then
                    For r = 2 To tbl.Rows.Count
                        v = Clean(tbl.cell(r, c).Range.Text)
                        If Len(v) > 0 Then vals.Add v, v
                    Next r
                    Dim out As String
                    For r = 1 To vals.Count
                        out = out & vals(r) & IIf(r < vals.Count, " & ", "")
                    Next r
                    GetUniqueAfterHeading = out
                    Exit Function
                End If
            End If
        End If
    End With
    GetUniqueAfterHeading = ""
End Function

'————————————————————————————————————
Function GetValueByFuzzyMatch(lookupTerm As String) As String
    Dim tbl As Table, r As Long, txt As String
    For Each tbl In ActiveDocument.Tables
        For r = 2 To tbl.Rows.Count
            txt = Clean(tbl.cell(r, 1).Range.Text)
            If InStr(1, txt, lookupTerm, vbTextCompare) > 0 Then
                GetValueByFuzzyMatch = Clean(tbl.cell(r, 2).Range.Text)
                Exit Function
            End If
        Next r
    Next tbl
    GetValueByFuzzyMatch = ""
End Function

'————————————————————————————————————
Function GetExactValue(label As String) As String
    Dim tbl As Table, r As Long
    For Each tbl In ActiveDocument.Tables
        For r = 2 To tbl.Rows.Count
            If LCase(Clean(tbl.cell(r, 1).Range.Text)) = LCase(label) Then
                GetExactValue = Clean(tbl.cell(r, 2).Range.Text)
                Exit Function
            End If
        Next r
    Next tbl
    GetExactValue = ""
End Function

'————————————————————————————————————
Function Clean(txt As String) As String
    txt = Replace(txt, Chr(7), "")
    txt = Replace(txt, Chr(13), "")
    txt = Replace(txt, Chr(160), "")
    txt = Replace(txt, ":", "")
    txt = Replace(txt, "=", "")
    Clean = Trim(txt)
End Function


