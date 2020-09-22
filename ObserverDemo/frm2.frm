VERSION 5.00
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Begin VB.Form frm2 
   Caption         =   "Form 2"
   ClientHeight    =   4080
   ClientLeft      =   5415
   ClientTop       =   2565
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   ScaleHeight     =   4080
   ScaleWidth      =   3750
   ShowInTaskbar   =   0   'False
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3735
      Left            =   600
      OleObjectBlob   =   "frm2.frx":0000
      TabIndex        =   0
      Top             =   120
      Width           =   2535
   End
End
Attribute VB_Name = "frm2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mData1 As cData
Attribute mData1.VB_VarHelpID = -1
Private WithEvents mData2 As cData
Attribute mData2.VB_VarHelpID = -1
Private WithEvents mTerminate As cData
Attribute mTerminate.VB_VarHelpID = -1

Private Sub Form_Load()
    On Error Resume Next
    Set mData1 = glObserverTool.getDataObject(modMain.ODATA1)
    Set mData2 = glObserverTool.getDataObject(modMain.ODATA2)
    Set mTerminate = glObserverTool.getDataObject(modMain.OTERMINATE)
    With MSChart1
        .Row = 1
        .Column = 1
        .Data = CInt(mData1)
        .Column = 2
        .Data = CInt(mData2)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mTerminate = Nothing
    Set mData2 = Nothing
    Set mData1 = Nothing
End Sub

Private Sub mData1_PostChange(vData As Variant)
    On Error Resume Next
    With MSChart1
        .Row = 1
        .Column = 1
        .Data = CInt(mData1)
    End With
End Sub

Private Sub mData2_PostChange(vData As Variant)
    On Error Resume Next
    With MSChart1
        .Row = 1
        .Column = 2
        .Data = CInt(mData2)
    End With
End Sub

Private Sub mTerminate_PostChange(vData As Variant)
    If vData = True Then
        Unload Me
    End If
End Sub

