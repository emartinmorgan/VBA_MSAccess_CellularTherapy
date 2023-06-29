VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_F_4100form_ptselect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub find_cupnx_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[cupnx] = " & Me![find_cupnx] & ""
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
    
    Forms!F_4100form_ptselect!find_visit.Requery
End Sub


Private Sub find_visit_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[cupnx] = " & Me![find_cupnx] & " And [visit] = '" & Me![find_visit] & "'"
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
    
    Dim strF As String
    With Me
    strF = "[Visit]='" & .find_visit & "'"
    .F_4100form_s1.Form.Filter = strF
    .F_4100form_s1.Form.FilterOn = True
    .F_4100form_s2_Toxicity.Form.Filter = strF
    .F_4100form_s2_Toxicity.Form.FilterOn = True
    .F_4100form_s3_Infection.Form.Filter = strF
    .F_4100form_s3_Infection.Form.FilterOn = True
    End With
End Sub
