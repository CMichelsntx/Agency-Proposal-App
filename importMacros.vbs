Option Explicit
Dim objWord, objNormal

' Start Word invisibly
Set objWord = CreateObject("Word.Application")
objWord.Visible = False

' Point at the Normal template
Set objNormal = objWord.NormalTemplate

' Import your two .bas modules
objNormal.VBProject.VBComponents.Import "C:\Agency Proposal\WOrd Proposal Script bkup Main 05-19.bas"
objNormal.VBProject.VBComponents.Import "C:\Agency Proposal\WOrd Proposal Script bkup EMAIL 05-19.bas"

' Save Normal.dotm and quit
objNormal.Save
objWord.Quit
