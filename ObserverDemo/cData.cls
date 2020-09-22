VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'The cData class raises the events that can be watched
Option Explicit

Public Event PreChange(vData As Variant)
Public Event PostChange(vData As Variant)

Private mData As Variant

Public Property Let vData(ByVal vData As Variant)
Attribute vData.VB_Description = "The data object."
Attribute vData.VB_UserMemId = 0
    If mData <> vData Then
        Debug.Print ">>> cData: RaiseEvent PreChange(" & vData & ")"
        RaiseEvent PreChange(vData)
        mData = vData
        Debug.Print ">>> cData: RaiseEvent PostChange(" & vData & ")"
        RaiseEvent PostChange(mData)
    End If
End Property

Public Property Get vData() As Variant
    vData = mData
End Property

Private Sub Class_Initialize()
    Debug.Print ">>> cData created"
End Sub

Private Sub Class_Terminate()
    Debug.Print ">>> cData destroyed"
End Sub
