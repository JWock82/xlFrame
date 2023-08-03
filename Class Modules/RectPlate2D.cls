VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RectPlate2D"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Represents a 2D plate element

'Dependencies:

'Force explicit variable declaration in this class module
Option Explicit

'Class variables
'Plate/node ID's
Private m_Name As String
Private m_PlateID As Integer
Private m_iNode As Node2D
Private m_jNode As Node2D
Private m_mNode As Node2D
Private m_nNode As Node2D

'Plate properties
Private m_E As Double
Private m_t As Double
Private m_nu As Double

'Class properties
Property Let Name(value As String)
    m_Name = value
End Property

Property Get Name() As String
    Name = m_Name
End Property

Property Let PlateID(value As Integer)
    m_PlateID = value
End Property

Property Get PlateID() As Integer
    PlateID = m_PlateID
End Property

Property Set iNode(value as Node2D)
    Set m_iNode = value
End Property

Property Get iNode() as Node2D
    Set iNode = m_iNode
End Property

Property Set jNode(value as Node2D)
    Set m_jNode = value
End Property

Property Get jNode() as Node2D
    Set jNode = m_jNode
End Property

Property Set mNode(value as Node2D)
    Set m_mNode = value
End Property

Property Get mNode() as Node2D
    Set mNode = m_mNode
End Property

Property Set nNode(value as Node2D)
    Set m_nNode = value
End Property

Property Get nNode() as Node2D
    Set nNode = m_nNode
End Property

Property Let E(value as Double)
    m_E = value
End Property

Property Get E() as Double
    E = m_E
End Property

Property Let t(value as double)
    m_t = value
End Property

Property Get t() as double
    t = m_t
End Property

Property Let nu(value as double)
    m_nu = value
End Property

Property Get nu(value as double)
    nu = m_nu
End Property

Public Function Width() as double
    Width = sqr((m_nNode.X - m_iNode.X)^2 + (m_nNode.Y - m_iNode.Y)^2)
End Function

Public Function Height() as double
    Height = sqr((m_jNode.X - m_iNode.X)^2 + (m_jNode.Y - m_iNode.Y)^2)
End Function

Public Function LocalStiff() as matrix

end Function