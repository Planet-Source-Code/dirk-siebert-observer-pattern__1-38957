VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frm3 
   Caption         =   "Form 3"
   ClientHeight    =   1635
   ClientLeft      =   1095
   ClientTop       =   4995
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   ScaleHeight     =   1635
   ScaleWidth      =   4185
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.Slider slData1 
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      Max             =   100
   End
   Begin MSComctlLib.Slider slData2 
      Height          =   375
      Left            =   960
      TabIndex        =   2
      Top             =   960
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   661
      _Version        =   393216
      Max             =   1000
   End
   Begin VB.Label lblData2 
      Alignment       =   1  'Rechts
      Caption         =   "Data 2:"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1020
      Width           =   615
   End
   Begin VB.Label lblData1 
      Alignment       =   1  'Rechts
      Caption         =   "Data 1:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   300
      Width           =   615
   End
End
Attribute VB_Name = "frm3"
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
    slData1 = CInt(mData1) Mod (slData1.Max + 1)
    slData2 = CInt(mData2) Mod (slData2.Max + 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mTerminate = Nothing
    Set mData2 = Nothing
    Set mData1 = Nothing
End Sub

Private Sub slData1_Change()
    mData1 = slData1.Value
End Sub

Private Sub slData2_Change()
    mData2 = slData2.Value
End Sub

Private Sub mData1_PostChange(vData As Variant)
    If slData1 <> vData Then
        slData1 = vData Mod (slData1.Max + 1)
    End If
End Sub

Private Sub mData2_PostChange(vData As Variant)
    If slData2 <> vData Then
        slData2 = vData Mod (slData2.Max + 1)
    End If
End Sub

Private Sub mTerminate_PostChange(vData As Variant)
    If vData = True Then
        Unload Me
    End If
End Sub

