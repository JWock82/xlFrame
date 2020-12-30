VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LoadCombo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Represents a load combination

'Force explicit variable declaration in this class module
Option Explicit

'Class variables
Public Factors As Object  'A collection containing each load case name and associated load factor

'Adds a load case with its associated factor
Public Sub AddLoadCase(CaseName As String, Factor As Double)
    Call Factors.Add(CaseName, Factor)
End Sub

'Deletes a load case with its associated factor
Public Sub DeleteLoadCase(CaseName As String)
    Call Factors.Remove(CaseName)
End Sub

Private Sub Class_Initialize()
    Set Factors = CreateObject("Scripting.Dictionary")
End Sub
