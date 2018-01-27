VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Node2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Node identification
Private m_Name As String        'User defined node name
Private m_ID As Long            'Numerical ID used by the program to track the node

'Node coordinates
Private m_XCoord As Double      'Node X-coordinate
Private m_YCoord As Double      'Node Y-coordinate

'Node displacements
Private m_DX As Double
Private m_DY As Double
Private m_RZ As Double

'Node supports
Private m_SupportDX As Boolean
Private m_SupportDY As Boolean
Private m_SupportRZ As Boolean

'Nodal loads
Private m_ForceX As Double
Private m_ForceY As Double
Private m_MomentZ As Double

'Nodal Reactions
Private m_ReactionX As Double
Private m_ReactionY As Double
Private m_ReactionMZ As Double

'Node identification
'********************************
Property Let Name(Value As String)
    m_Name = Value
End Property

Property Get Name() As String
    Name = m_Name
End Property

Property Let ID(Value As Long)
    m_ID = Value
End Property

Property Get ID() As Long
    ID = m_ID
End Property

'Node coordinates
'******************************
Property Let XCoord(Value As Double)
    m_XCoord = Value
End Property

Property Get XCoord() As Double
    XCoord = m_XCoord
End Property

Property Let YCoord(Value As Double)
    m_YCoord = Value
End Property

Property Get YCoord() As Double
    YCoord = m_YCoord
End Property

'Node displacements
'*******************************
Property Let DX(Value As Double)
    m_DX = Value
End Property

Property Get DX() As Double
    DX = m_DX
End Property

Property Let DY(Value As Double)
    m_DY = Value
End Property

Property Get DY() As Double
    DY = m_DY
End Property

Property Let RZ(Value As Double)
    m_RZ = Value
End Property

Property Get RZ() As Double
    RZ = m_RZ
End Property

'Supports
'************************************
Property Let SupportDX(Value As Boolean)
    m_SupportDX = Value
End Property

Property Get SupportDX() As Boolean
    SupportDX = m_SupportDX
End Property

Property Let SupportDY(Value As Boolean)
    m_SupportDY = Value
End Property

Property Get SupportDY() As Boolean
    SupportDY = m_SupportDY
End Property

Property Let SupportRZ(Value As Boolean)
    m_SupportRZ = Value
End Property

Property Get SupportRZ() As Boolean
    SupportRZ = m_SupportRZ
End Property

Public Sub AddForceX(Value As Double)
    m_ForceX = m_ForceX + Value
End Sub

Public Sub AddForceY(Value As Double)
    m_ForceY = m_ForceY + Value
End Sub

Public Sub AddMomentZ(Value As Double)
    m_MomentZ = m_MomentZ + Value
End Sub

Public Function GetForceX() As Double
    GetForceX = m_ForceX
End Function

Public Function GetForceY() As Double
    GetForceY = m_ForceY
End Function

Public Function GetMomentZ() As Double
    GetMomentZ = m_MomentZ
End Function

Public Sub ClearForces()
    m_ForceX = 0
    m_ForceY = 0
    m_MomentZ = 0
End Sub

Public Sub ClearDisplacements()
    m_DX = 0
    m_DY = 0
    m_RZ = 0
End Sub

Public Sub ClearReactions()
    m_ReactionX = 0
    m_ReactionY = 0
    m_ReactionMZ = 0
End Sub

Private Sub Class_Initialize()
    Call ClearForces
    Call ClearDisplacements
    Call ClearReactions
End Sub
