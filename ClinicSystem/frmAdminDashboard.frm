VERSION 5.00
Begin VB.Form frmAdminDashboard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dashboard"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17310
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
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
      Height          =   8295
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
      Height          =   8295
      Left            =   2520
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin VB.Frame fraHistory 
         Height          =   8295
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   14775
         Begin VB.TextBox txtToms 
            Height          =   3495
            Left            =   6120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   27
            Top             =   2640
            Width           =   4815
         End
         Begin VB.TextBox txtSex 
            Height          =   615
            Left            =   3000
            TabIndex        =   19
            Top             =   5520
            Width           =   3015
         End
         Begin VB.TextBox txtDOB 
            Height          =   615
            Left            =   3000
            TabIndex        =   18
            Top             =   4800
            Width           =   3015
         End
         Begin VB.TextBox txtAge 
            Height          =   615
            Left            =   3000
            TabIndex        =   17
            Top             =   4080
            Width           =   3015
         End
         Begin VB.TextBox txtAddress 
            Height          =   615
            Left            =   3000
            TabIndex        =   16
            Top             =   3360
            Width           =   3015
         End
         Begin VB.TextBox txtName 
            Height          =   615
            Left            =   3000
            TabIndex        =   15
            Top             =   2640
            Width           =   3015
         End
         Begin VB.TextBox txtID 
            Height          =   615
            Left            =   3000
            TabIndex        =   14
            Top             =   1920
            Width           =   3015
         End
         Begin VB.CommandButton cmdReturn01 
            Caption         =   "Return"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   12600
            TabIndex        =   13
            Top             =   7320
            Width           =   1935
         End
         Begin VB.CommandButton cmdPrev 
            Caption         =   "Previous"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   240
            TabIndex        =   12
            Top             =   7320
            Width           =   1935
         End
         Begin VB.CommandButton cmdNext 
            Caption         =   "Next"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   2280
            TabIndex        =   11
            Top             =   7320
            Width           =   1935
         End
         Begin VB.Label lblToms 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Symptoms"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   6120
            TabIndex        =   28
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblSex 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Patient Sex"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   480
            TabIndex        =   26
            Top             =   5520
            Width           =   1935
         End
         Begin VB.Label lblAge 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Patient Age"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   480
            TabIndex        =   25
            Top             =   4080
            Width           =   2175
         End
         Begin VB.Label lblDOB 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Patient DOB"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   480
            TabIndex        =   24
            Top             =   4800
            Width           =   2655
         End
         Begin VB.Label lblAdress 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Address"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   480
            TabIndex        =   23
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Label lblName 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Patient Name"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   480
            TabIndex        =   22
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Label lblID 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Patient ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   480
            TabIndex        =   21
            Top             =   1920
            Width           =   1815
         End
         Begin VB.Label lblDetail 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Patient History"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   24
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   120
            TabIndex        =   20
            Top             =   360
            Width           =   4575
         End
      End
      Begin VB.Frame fraInv 
         Height          =   8295
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
      Begin VB.Frame fraReport 
         Height          =   8295
         Left            =   -120
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
   End
End
Attribute VB_Name = "frmAdminDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset

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

Private Sub cmdReturn01_Click()
    fraHistory.Visible = False
End Sub

Private Sub Form_Load()
    fraHistory.Visible = False
    fraInv.Visible = False
    fraReport.Visible = False
    
    Set cn = New ADODB.Connection
    cn.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\PatientRecord.mdb"


    Set rs = New ADODB.Recordset
    rs.Open "SELECT * FROM patient_master", cn, adOpenDynamic, adLockOptimistic


    If rs.RecordCount > 0 Then
        rs.MoveFirst
        Call ShowRecord
    End If
End Sub

Private Sub ShowRecord()
    If rs Is Nothing Then Exit Sub
    If rs.EOF Or rs.BOF Then Exit Sub

    txtID.Text = rs!ID
    txtName.Text = rs!Name
    txtAddress.Text = rs!Address
    txtAge.Text = rs!Age
    txtDOB.Text = rs!DOB
    txtSex.Text = rs!Sex
    txtToms.Text = rs!Symptoms
End Sub

Private Sub cmdNext_Click()
    If rs.EOF Then Exit Sub
    rs.MoveNext
    If rs.EOF Then rs.MoveLast
    Call ShowRecord
End Sub

Private Sub cmdPrev_Click()
    If rs.BOF Then Exit Sub
    rs.MovePrevious
    If rs.BOF Then rs.MoveFirst
    Call ShowRecord
End Sub

