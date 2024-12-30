VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IALevelData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Private member variables
Private m_GCI As String
Private m_Region As String
Private m_Manager As String
Private m_TriggerStatus As String
Private m_NavSources As Collection
Private m_ClientContacts As String
Private m_TriggerCount As Long
Private m_NonTriggerCount As Long
Private m_MissingTriggerCount As Long
Private m_MissingNonTriggerCount As Long
Private m_ManualData As Variant

' Initialize the class
Private Sub Class_Initialize()
    Set m_NavSources = New Collection
End Sub

' Clean up
Private Sub Class_Terminate()
    Set m_NavSources = Nothing
End Sub

' Property Get/Let methods
Public Property Get GCI() As String
    GCI = m_GCI
End Property

Public Property Let GCI(value As String)
    m_GCI = value
End Property

Public Property Get Region() As String
    Region = m_Region
End Property

Public Property Let Region(value As String)
    m_Region = value
End Property

Public Property Get Manager() As String
    Manager = m_Manager
End Property

Public Property Let Manager(value As String)
    m_Manager = value
End Property

Public Property Get TriggerStatus() As String
    TriggerStatus = m_TriggerStatus
End Property

Public Property Let TriggerStatus(value As String)
    m_TriggerStatus = value
End Property

Public Property Get NavSources() As Collection
    Set NavSources = m_NavSources
End Property

Public Property Set NavSources(value As Collection)
    Set m_NavSources = value
End Property

Public Property Get ClientContacts() As String
    ClientContacts = m_ClientContacts
End Property

Public Property Let ClientContacts(value As String)
    m_ClientContacts = value
End Property

Public Property Get TriggerCount() As Long
    TriggerCount = m_TriggerCount
End Property

Public Property Let TriggerCount(value As Long)
    m_TriggerCount = value
End Property

Public Property Get NonTriggerCount() As Long
    NonTriggerCount = m_NonTriggerCount
End Property

Public Property Let NonTriggerCount(value As Long)
    m_NonTriggerCount = value
End Property

Public Property Get MissingTriggerCount() As Long
    MissingTriggerCount = m_MissingTriggerCount
End Property

Public Property Let MissingTriggerCount(value As Long)
    m_MissingTriggerCount = value
End Property

Public Property Get MissingNonTriggerCount() As Long
    MissingNonTriggerCount = m_MissingNonTriggerCount
End Property

Public Property Let MissingNonTriggerCount(value As Long)
    m_MissingNonTriggerCount = value
End Property

Public Property Get ManualData() As Variant
    If IsObject(m_ManualData) Then
        Set ManualData = m_ManualData
    Else
        ManualData = m_ManualData
    End If
End Property

Public Property Let ManualData(value As Variant)
    m_ManualData = value
End Property

Public Property Set ManualData(value As Variant)
    Set m_ManualData = value
End Property
