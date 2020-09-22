VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cObserverTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'cObserverTool contains mechanism to implement the observer or
'OCV (Object Controler View) pattern

Option Explicit

Private mObserver As New Dictionary

'The internal storage is searched for paramName and the related entry (object)
'is returned. If there is no related entry NOTHING is returned.
Public Function getDataObject(vName As String) As cData

    If Not mObserver.Exists(vName) Then
        createDataObject vName
    End If
    Set getDataObject = mObserver(vName)

End Function

'Create a data object and add it to the internal storage
Private Sub createDataObject(vName As String)

    Dim objNew As cData
    Set objNew = New cData

    If mObserver.Exists(vName) Then
        Debug.Print vName & " already exists"
    Else
        mObserver.Add vName, objNew
        Debug.Print vName & " registered"
    End If

End Sub

'Init Observer and create internal storage
Private Sub Class_Initialize()
    
    Set mObserver = New Dictionary

End Sub

'Destroy internal storage
Private Sub Class_Terminate()
    
    Set mObserver = Nothing

End Sub
