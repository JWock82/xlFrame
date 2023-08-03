VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Node2D
'A class represending a 2D node in a finite element model

'Node identification
Public Name As String           'User defined node name
Public ID As Long               'Numerical ID used by the program to track the node

'Node coordinates
Public XCoord As Double         'Node X-coordinate
Public YCoord As Double         'Node Y-coordinate

'Nodal loads
Public NodeLoads As EZArray     'An array of loads applied to the node (Direction, P, LoadCase) or (Direction, M, LoadCase)

'Node displacements
Public DX As Object
Public DY As Object
Public RZ As Object

'Node supports
Public SupportDX As Boolean
Public SupportDY As Boolean
Public SupportRZ As Boolean

'Nodal Reactions
Public ReactionX As Object
Public ReactionY As Object
Public ReactionMZ As Object

'Class Initialization
Private Sub Class_Initialize()
    
    Set NodeLoads = New EZArray
    
    Set DX = CreateObject("Scripting.Dictionary")
    Set DY = CreateObject("Scripting.Dictionary")
    Set RZ = CreateObject("Scripting.Dictionary")
    
    Set ReactionX = CreateObject("Scripting.Dictionary")
    Set ReactionY = CreateObject("Scripting.Dictionary")
    Set ReactionMZ = CreateObject("Scripting.Dictionary")
    
End Sub
