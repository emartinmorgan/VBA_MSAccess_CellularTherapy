VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_F_Products_Infusions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database
Option Explicit


Private Sub Detail_Click()

End Sub

Private Sub find_cupnx_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[cupnx] = " & Me![find_cupnx] & ""
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
    
    Forms!F_Products_Infusions!find_cupnxp.Requery
End Sub


'Private Sub find_cupnxp_AfterUpdate()
'Dim strF As String
'With Me
'strF = "[cupnxp]='" & .find_cupnxp & "'"
'.F_4003form.Form.Filter = strF
'.F_4003form.Form.FilterOn = True
'.F_4006form.Form.Filter = strF
'.F_4006form.Form.FilterOn = True
'End With
'End Sub

Private Sub find_cupnxp_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[cupnxp] = " & Me![find_cupnxp] & ""
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
    
    Forms!F_Products_Infusions!find_cupnxpi.Requery
End Sub

Private Sub find_cupnxpi_AfterUpdate()
    ' Find the record that matches the control.
    Dim rs As Object

    Set rs = Me.Recordset.Clone
    rs.FindFirst "[cupnxpi] = " & Me![find_cupnxpi] & ""
    If Not rs.EOF Then Me.Bookmark = rs.Bookmark
End Sub
