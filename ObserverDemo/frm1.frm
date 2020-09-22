VERSION 5.00
Begin VB.Form frm1 
   Caption         =   "Form 1"
   ClientHeight    =   2085
   ClientLeft      =   1095
   ClientTop       =   2565
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   2085
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtData2 
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   1260
      Width           =   2055
   End
   Begin VB.TextBox txtData1 
      Height          =   375
      Left            =   1440
      TabIndex        =   0
      Top             =   420
      Width           =   2055
   End
   Begin VB.Label lblData2 
      Alignment       =   1  'Rechts
      Caption         =   "Data 2:"
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.Label lblData1 
      Alignment       =   1  'Rechts
      Caption         =   "Data 1:"
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
End
Attribute VB_Name = "frm1"
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
    Set mData1 = glObserverTool.getDataObject(modMain.ODATA1)
    Set mData2 = glObserverTool.getDataObject(modMain.ODATA2)
    Set mTerminate = glObserverTool.getDataObject(modMain.OTERMINATE)
    txtData1 = CStr(mData1)
    txtData2 = CStr(mData2)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mTerminate = Nothing
    Set mData2 = Nothing
    Set mData1 = Nothing
End Sub

Private Sub txtData1_KeyPress(KeyAscii As Integer)
    txtBoxUnsignedInt KeyAscii
End Sub

Private Sub txtData1_Change()
    On Error Resume Next
    If txtData1 = "" Then
        txtData1 = "0"
    ElseIf CInt(txtData1) > 100 Then
        txtData1 = "100"
    End If
    mData1 = CInt(txtData1)
End Sub

Private Sub txtData2_KeyPress(KeyAscii As Integer)
    txtBoxUnsignedInt KeyAscii
End Sub

Private Sub txtData2_Change()
    On Error Resume Next
    If txtData2 = "" Then
        txtData2 = "0"
    ElseIf CInt(txtData2) > 1000 Then
        txtData2 = "1000"
    End If
    mData2 = CInt(txtData2)
End Sub

Private Sub mData1_PostChange(vData As Variant)
    If txtData1 <> CStr(vData) Then
        txtData1 = CStr(vData)
    End If
End Sub

Private Sub mData2_PostChange(vData As Variant)
    If txtData2 <> CStr(vData) Then
        txtData2 = CStr(vData)
    End If
End Sub

Private Sub mTerminate_PostChange(vData As Variant)
    If vData = True Then
        Unload Me
    End If
End Sub

