VERSION 5.00
Begin VB.Form frmAdminDashboard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dashboard"
   ClientHeight    =   7950
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7950
   ScaleWidth      =   17310
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdInv 
      Caption         =   "View Medicine Inventory"
      Height          =   615
      Left            =   480
      TabIndex        =   4
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame fraCtrlPanel 
      Height          =   7935
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton cmdReport 
         Caption         =   "Create Patient Report"
         Height          =   615
         Left            =   480
         TabIndex        =   5
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton cmdHistory 
         Caption         =   "View Patient History"
         Height          =   615
         Left            =   480
         TabIndex        =   3
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton cmdLogout 
         Caption         =   "Logout"
         Height          =   495
         Left            =   480
         TabIndex        =   2
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   960
         Shape           =   3  'Circle
         Top             =   480
         Width           =   495
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   720
         Shape           =   2  'Oval
         Top             =   960
         Width           =   975
      End
   End
   Begin VB.Frame fraMain 
      Height          =   7935
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin VB.Frame fraHistory 
         Height          =   7935
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   14775
      End
      Begin VB.Frame fraReport 
         Height          =   7935
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   14775
         Begin VB.TextBox Text1 
            Height          =   615
            Left            =   360
            TabIndex        =   9
            Text            =   "Report"
            Top             =   480
            Width           =   2175
         End
      End
      Begin VB.Frame fraInv 
         Height          =   7935
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   14775
         Begin VB.TextBox Text2 
            Height          =   615
            Left            =   360
            TabIndex        =   10
            Text            =   "Inventory"
            Top             =   480
            Width           =   2175
         End
      End
   End
End
Attribute VB_Name = "frmAdminDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdLogout_Click()
Unload Me
frmLogin.Show
MsgBox "Successfully Logged Out!"

End Sub

Private Sub ShowFrame(fra As Frame)
    fraHistory.Visible = False
    fraInv.Visible = False
    fraReport.Visible = False
    
    fra.Visible = True
End Sub

Private Sub cmdHistory_Click()
    ShowFrame fraHistory
End Sub

Private Sub cmdInv_Click()
    ShowFrame fraInv
End Sub

Private Sub cmdReport_Click()
    ShowFrame fraReport
End Sub

Private Sub Form_Load()
    fraHistory.Visible = False
    fraInv.Visible = False
    fraReport.Visible = False
End Sub
