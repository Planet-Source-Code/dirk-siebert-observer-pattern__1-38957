Attribute VB_Name = "modMain"
Option Explicit

'List of things to be observed
Public Const ODATA1 As String = "A"
Public Const ODATA2 As String = "B"
Public Const OTERMINATE As String = "Terminate"

'The observer helper object
Public glObserverTool As cObserverTool

'Only to initialize the data ...
Private mData1 As cData
Private mData2 As cData

Public Sub Main()

    Debug.Print ""
    Debug.Print ">>>>>>>>>>>>>>>>>>>>>>>>>>>"
    Debug.Print ">>>                     >>>"
    Debug.Print ">>> Observer Demo Start >>>"
    Debug.Print ">>>                     >>>"
    Debug.Print ">>>>>>>>>>>>>>>>>>>>>>>>>>>"

    Set glObserverTool = New cObserverTool
    
    Set mData1 = glObserverTool.getDataObject(modMain.ODATA1)
    Set mData2 = glObserverTool.getDataObject(modMain.ODATA2)
    mData1 = 10
    mData2 = 333

    frmMain.Show
    frm1.Show
    frm2.Show
    frm3.Show
    
End Sub

Public Sub Terminate()
    
    Set mData2 = Nothing
    Set mData1 = Nothing

    Set glObserverTool = Nothing
    
    Debug.Print "<<<<<<<<<<<<<<<<<<<<<<<<<<<"
    Debug.Print "<<<                     <<<"
    Debug.Print "<<< Observer Demo End   <<<"
    Debug.Print "<<<                     <<<"
    Debug.Print "<<<<<<<<<<<<<<<<<<<<<<<<<<<"
    Debug.Print ""

End Sub

Public Sub txtBoxUnsignedInt(ByRef KeyAscii As Integer)
    Select Case KeyAscii
        Case vbKeyReturn, vbKeyDelete, vbKeyBack, vbKey0 To vbKey9
        Case Else
            KeyAscii = 0
    End Select
End Sub


