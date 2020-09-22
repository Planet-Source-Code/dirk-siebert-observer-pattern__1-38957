VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Callback"
   ClientHeight    =   1380
   ClientLeft      =   2670
   ClientTop       =   795
   ClientWidth     =   5475
   LinkTopic       =   "Form1"
   ScaleHeight     =   1380
   ScaleWidth      =   5475
   Begin VB.CommandButton Command3 
      Caption         =   "Open 3"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Open 2"
      Height          =   495
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open 1"
      Height          =   495
      Left            =   600
      TabIndex        =   0
      Top             =   480
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mTerminate As cData
Attribute mTerminate.VB_VarHelpID = -1

Private Sub Form_Load()
    Set mTerminate = glObserverTool.getDataObject(modMain.OTERMINATE)
End Sub

Private Sub Command1_Click()
    frm1.Show
End Sub

Private Sub Command2_Click()
    frm2.Show
End Sub

Private Sub Command3_Click()
    frm3.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mTerminate = True
    Set mTerminate = Nothing
    Terminate
End Sub
