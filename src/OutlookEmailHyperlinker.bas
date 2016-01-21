VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisOutlookSession"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'########################################
'# This is a Outlook 2013 VBA snippet to replace 
'# text in email (incoming/outgoing/anything - that depends on how you setup your outlook rules)
'# with an equivalent hyperlink - See the sample input/output strings in README.md
'# 
'# Author  : kmmankad@gmail.com
'# License : MIT License
'# Date    : 2016-01-21


Option Explicit

'#########################################
'# InsertBugLinks
'# Desc: This subroutine will open the current email item
'#       and find a particular Bug ID of the form Bug#123456
'#	 and replace that with a hyperlink to http://bug/123456, keeping Bug#123456 as the link text.
Sub InsertBugLinks()

    Dim objMsg As Outlook.MailItem
    Dim objItem As Object
    Dim objExpl As Outlook.Explorer
    Dim objInsp As Outlook.Inspector
    ' Add a reference to the Microsoft Word to your project.
    Dim objDoc As Word.Document
    Dim objRange As Word.Range
    Dim BugID As String
    
    Dim BugIDRegex As String
    BugIDRegex = "Bug#([0-9]{1,7})"
    
    On Error Resume Next
     
     
    Set objExpl = Application.ActiveExplorer
    Set objItem = objExpl.Selection(1)
    If Not objItem Is Nothing Then
        If objItem.Class = olMail Then
            Set objMsg = objItem.Forward
            Set objInsp = objMsg.GetInspector
            If objInsp.EditorType = olEditorWord Then
                ' next statement triggers security prompt
                ' in Outlook 2002 SP3
                Set objDoc = objInsp.WordEditor
                ' Clear Formatting before doing a search
                Set objRange = objDoc.Range
                With objRange.Find
                    .ClearFormatting
                    .Text = BugIDRegex
                    .MatchAllWordForms = False
                    .MatchSoundsLike = False
                    .MatchFuzzy = False
                    .MatchWildcards = True
		    ' Iterate over all the matches in the email body
                    Do While .Execute(Format:=False, Forward:=True) = True
                        ' Get the number part of the text
                        BugID = GetBugID(objRange.Text, BugIDRegex)
                        ' Add a Hyperlink there
                        objDoc.Hyperlinks.Add Anchor:=objRange, _
                        Address:="http://bug/" & BugID, _
                        SubAddress:="", ScreenTip:="", TextToDisplay:=objRange.Text
                        With objRange
                            .End = objRange.Hyperlinks(1).Range.End
                            .Collapse 0
                        End With
                    Loop
                End With
                ' Display the email for quick debug               
                'objMsg.Display
            Else
                MsgBox "Cannot insert text in a formatted message unless Word is the editor", vbCritical
            End If
        End If
    End If
     
    Set objInsp = Nothing
    Set objDoc = Nothing
    Set objExpl = Nothing
    Set objItem = Nothing
    Set objMsg = Nothing


End Sub


'#########################################
'# InsertBugLinks
'# Desc: This function will return text from BugWord as referenced by 
'#       the capture groups used in BugIDPattern
Function GetBugID(BugWord As String, BugIDPattern As String) As String

	Dim myRegExp As Object
	Dim myMatches As Object

	Set myRegExp = New RegExp

	With myRegExp
		.Pattern = BugIDPattern
		.IgnoreCase = True
		.Global = False

		If .Test(BugWord) Then
		    GetBugID = .Execute(BugWord)(0)
		Else
		    GetBugID = "NaN"
		End If

	End With

End Function



