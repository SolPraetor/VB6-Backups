VERSION 5.00
Begin VB.Form frmUserDashboard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Dashboard"
   ClientHeight    =   8325
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   17340
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8325
   ScaleWidth      =   17340
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraMain 
      Height          =   8295
      Left            =   2520
      TabIndex        =   6
      Top             =   0
      Width           =   14775
      Begin VB.Frame fraAddPatient 
         Height          =   8295
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   14775
         Begin VB.TextBox txtToms 
            Height          =   3495
            Left            =   6240
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   49
            Top             =   2640
            Width           =   4815
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
            TabIndex        =   30
            Top             =   7320
            Width           =   1935
         End
         Begin VB.TextBox txtID 
            Height          =   615
            Left            =   3000
            TabIndex        =   18
            Top             =   1920
            Width           =   3015
         End
         Begin VB.TextBox txtName 
            Height          =   615
            Left            =   3000
            TabIndex        =   17
            Top             =   2640
            Width           =   3015
         End
         Begin VB.TextBox txtAddress 
            Height          =   615
            Left            =   3000
            TabIndex        =   16
            Top             =   3360
            Width           =   3015
         End
         Begin VB.TextBox txtAge 
            Height          =   615
            Left            =   3000
            TabIndex        =   15
            Top             =   4080
            Width           =   3015
         End
         Begin VB.TextBox txtDOB 
            Height          =   615
            Left            =   3000
            TabIndex        =   14
            Top             =   4800
            Width           =   3015
         End
         Begin VB.TextBox txtSex 
            Height          =   615
            Left            =   3000
            TabIndex        =   13
            Top             =   5520
            Width           =   3015
         End
         Begin VB.CommandButton cmdCon01 
            Caption         =   "Confirm"
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
            Left            =   10440
            TabIndex        =   12
            Top             =   7320
            Width           =   1935
         End
         Begin VB.CommandButton cmdPicture 
            Caption         =   "Add Picture"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   735
            Left            =   9120
            TabIndex        =   11
            Top             =   600
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
            Left            =   6240
            TabIndex        =   50
            Top             =   2040
            Width           =   1815
         End
         Begin VB.Label lblDetail 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Add Patient Details"
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
            TabIndex        =   25
            Top             =   360
            Width           =   4575
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
            TabIndex        =   24
            Top             =   1920
            Width           =   1815
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
            TabIndex        =   23
            Top             =   2640
            Width           =   2175
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
            TabIndex        =   22
            Top             =   3360
            Width           =   1335
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
            TabIndex        =   21
            Top             =   4800
            Width           =   2655
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
            TabIndex        =   20
            Top             =   4080
            Width           =   2175
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
            TabIndex        =   19
            Top             =   5520
            Width           =   1935
         End
         Begin VB.Image ImgPatient 
            Height          =   2775
            Left            =   11280
            Top             =   480
            Width           =   3135
         End
      End
      Begin VB.Frame fraEdit 
         Height          =   8295
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   14775
         Begin VB.TextBox lblSymptoms 
            Height          =   1335
            Left            =   360
            TabIndex        =   47
            Top             =   6840
            Width           =   5655
         End
         Begin VB.CommandButton cmdEdit 
            Caption         =   "Edit Data"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   15
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   3000
            TabIndex        =   46
            Top             =   1200
            Width           =   1455
         End
         Begin VB.TextBox txtTemp 
            Height          =   615
            Left            =   3000
            TabIndex        =   37
            Top             =   5520
            Width           =   3015
         End
         Begin VB.TextBox txtPulse 
            Height          =   615
            Left            =   3000
            TabIndex        =   36
            Top             =   4800
            Width           =   3015
         End
         Begin VB.TextBox txtBP 
            Height          =   615
            Left            =   3000
            TabIndex        =   35
            Top             =   4080
            Width           =   3015
         End
         Begin VB.TextBox txtWeight 
            Height          =   615
            Left            =   3000
            TabIndex        =   34
            Top             =   3360
            Width           =   3015
         End
         Begin VB.TextBox txtDate 
            Height          =   615
            Left            =   3000
            TabIndex        =   33
            Top             =   2640
            Width           =   3015
         End
         Begin VB.TextBox txtNameE 
            Height          =   615
            Left            =   3000
            TabIndex        =   32
            Top             =   1920
            Width           =   3015
         End
         Begin VB.Label Label3 
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
            Left            =   360
            TabIndex        =   48
            Top             =   6360
            Width           =   2175
         End
         Begin VB.Label lblData 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Patient Data"
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
            TabIndex        =   45
            Top             =   1200
            Width           =   2295
         End
         Begin VB.Label lblTemp 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Temperature"
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
            TabIndex        =   44
            Top             =   5520
            Width           =   2175
         End
         Begin VB.Label lblBP 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "BP"
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
            TabIndex        =   43
            Top             =   4080
            Width           =   2175
         End
         Begin VB.Label lblPulse 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Pulse"
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
            TabIndex        =   42
            Top             =   4800
            Width           =   2415
         End
         Begin VB.Label lblWeight 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Weight"
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
            TabIndex        =   41
            Top             =   3360
            Width           =   1335
         End
         Begin VB.Label lblDate 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Date"
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
            TabIndex        =   40
            Top             =   2640
            Width           =   2175
         End
         Begin VB.Label lblNameE 
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
            TabIndex        =   39
            Top             =   1920
            Width           =   2175
         End
         Begin VB.Label Label2 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Add Patient Details"
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
            TabIndex        =   38
            Top             =   360
            Width           =   4575
         End
      End
      Begin VB.Frame fraPrescription 
         Height          =   8295
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   14775
      End
      Begin VB.Frame fraRemovePatient 
         Height          =   8295
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   14775
         Begin VB.CommandButton cmdReturn02 
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
            TabIndex        =   31
            Top             =   7320
            Width           =   1935
         End
         Begin VB.CommandButton cmdCon02 
            Caption         =   "Confirm"
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
            Left            =   10440
            TabIndex        =   29
            Top             =   7320
            Width           =   1935
         End
         Begin VB.TextBox txtRID 
            Height          =   615
            Left            =   3720
            TabIndex        =   28
            Top             =   2400
            Width           =   2895
         End
         Begin VB.Label lblRID 
            Caption         =   "Enter Patient ID"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   18
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   600
            TabIndex        =   27
            Top             =   2400
            Width           =   2895
         End
         Begin VB.Label Label1 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "Remove Patient Details"
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
            TabIndex        =   26
            Top             =   360
            Width           =   5295
         End
      End
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove Patient Data"
      Height          =   615
      Left            =   480
      TabIndex        =   3
      Top             =   3240
      Width           =   1455
   End
   Begin VB.Frame fraCtrlPanel 
      Height          =   8295
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
      Begin VB.CommandButton cmdEditPatient 
         Caption         =   "Edit Patient Data"
         Height          =   615
         Left            =   480
         TabIndex        =   5
         Top             =   4680
         Width           =   1455
      End
      Begin VB.CommandButton cmdMed 
         Caption         =   "Add Patient Prescription"
         Height          =   615
         Left            =   480
         TabIndex        =   4
         Top             =   3960
         Width           =   1455
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add Patient Data"
         Height          =   615
         Left            =   480
         TabIndex        =   2
         Top             =   2520
         Width           =   1455
      End
      Begin VB.CommandButton cmdLogout 
         Caption         =   "Logout"
         Height          =   495
         Left            =   480
         TabIndex        =   1
         Top             =   1440
         Width           =   1455
      End
      Begin VB.Shape Shape2 
         Height          =   735
         Left            =   720
         Shape           =   2  'Oval
         Top             =   960
         Width           =   975
      End
      Begin VB.Shape Shape1 
         Height          =   495
         Left            =   960
         Shape           =   3  'Circle
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmUserDashboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cn As ADODB.Connection
Dim rs As ADODB.Recordset


Private Sub ShowFrame(fra As Frame)
    fraAddPatient.Visible = False
    fraRemovePatient.Visible = False
    fraEdit.Visible = False
    fraPrescription.Visible = False

    fra.Visible = True
End Sub

Private Sub cmdAdd_Click()
    ShowFrame fraAddPatient
End Sub

Private Sub cmdCon02_Click()
    Dim sql As String

    If Trim(txtRID.Text) = "" Then
        MsgBox "Please enter a Patient ID.", vbExclamation
        Exit Sub
    End If

    sql = "SELECT * FROM patient_master WHERE ID = " & txtRID.Text
    Set rs = New ADODB.Recordset
    rs.Open sql, cn, adOpenDynamic, adLockOptimistic

    If rs.EOF Then
        MsgBox "No record found with that ID.", vbInformation
    Else
        If MsgBox("Are you sure you want to delete this record?", vbYesNo + vbQuestion) = vbYes Then
            rs.Delete
            MsgBox "Record deleted successfully!"
        End If
    End If

    rs.Close
    
End Sub

Private Sub cmdReturn01_Click()
    fraAddPatient.Visible = False
End Sub

Private Sub cmdEditPatient_Click()
    ShowFrame fraEdit
End Sub

Private Sub cmdLogout_Click()
Unload Me
frmLogin.Show
MsgBox "Successfully Logged Out!"

End Sub

Private Sub cmdMed_Click()
    ShowFrame fraPrescription
End Sub

Private Sub cmdRemove_Click()
    ShowFrame fraRemovePatient
End Sub

Private Sub cmdReturn02_Click()
    fraRemovePatient.Visible = False
End Sub

Private Sub Form_Load()
    fraAddPatient.Visible = False
    fraRemovePatient.Visible = False
    fraPrescription.Visible = False
    fraEdit.Visible = False


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

Private Sub cmdCon01_Click()
    rs.AddNew
    rs!ID = txtID.Text
    rs!Name = txtName.Text
    rs!Address = txtAddress.Text
    rs!Age = txtAge.Text
    rs!DOB = txtDOB.Text
    rs!Sex = txtSex.Text
    rs!Symptoms = txtToms.Text
    rs.Update
    MsgBox "Patient record saved successfully!"

End Sub

