VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFormOptionButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
''' This class tracks changes made in frames in order to perform certain actions

Option Explicit

' in clsFormOptionButton code module

Public WithEvents OptionButton As MSForms.OptionButton
Attribute OptionButton.VB_VarHelpID = -1
Dim myColl As Collection

Private Sub Class_Initialize()
    Set myColl = New Collection
    myColl.Add item:=Me
End Sub

Private Sub OptionButton_Click()
    On Error Resume Next
    CallByName UFParent, OptionButton.Parent.Name & "OptionButton_Click", VbMethod
    On Error GoTo 0
End Sub

Private Function UFParent() As Object
    Set UFParent = OptionButton
    On Error Resume Next
    Do
    Set UFParent = UFParent.Parent
    Loop Until Err
    On Error GoTo 0
End Function
