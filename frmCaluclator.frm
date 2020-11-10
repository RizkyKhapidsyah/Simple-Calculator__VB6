VERSION 5.00
Begin VB.Form frmCaluclator 
   Caption         =   "Simple Calculator"
   ClientHeight    =   3870
   ClientLeft      =   5115
   ClientTop       =   3420
   ClientWidth     =   3990
   Icon            =   "frmCaluclator.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3870
   ScaleWidth      =   3990
   Begin VB.Frame fraOperators 
      BorderStyle     =   0  'None
      Height          =   1695
      Left            =   360
      TabIndex        =   8
      Top             =   2040
      Width           =   1935
      Begin VB.CommandButton cmdADD 
         Caption         =   "+"
         Height          =   615
         Left            =   0
         TabIndex        =   15
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdSUBTRACT 
         Caption         =   "-"
         Height          =   615
         Left            =   600
         TabIndex        =   14
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdMUTIPLY 
         Caption         =   "*"
         Height          =   615
         Left            =   1200
         TabIndex        =   13
         Top             =   0
         Width           =   615
      End
      Begin VB.CommandButton cmdDIVID 
         Caption         =   "/"
         Height          =   615
         Left            =   0
         TabIndex        =   12
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton Command4 
         Caption         =   "\"
         Height          =   615
         Left            =   600
         TabIndex        =   11
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdPOWER 
         Caption         =   "^"
         Height          =   615
         Left            =   1200
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.CommandButton cmdMOD 
         Caption         =   "MOD"
         Height          =   495
         Left            =   0
         TabIndex        =   9
         Top             =   1200
         Width           =   1815
      End
   End
   Begin VB.TextBox txtONE 
      Height          =   375
      Left            =   480
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.TextBox txtTWO 
      Height          =   375
      Left            =   2160
      TabIndex        =   6
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame fraDISPLAY 
      Height          =   1215
      Left            =   360
      TabIndex        =   2
      Top             =   720
      Width           =   3255
      Begin VB.Label lblEQUAL 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1680
         TabIndex        =   5
         Top             =   240
         Width           =   495
      End
      Begin VB.Label lblFORMULA 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblANSWER 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   3015
      End
   End
   Begin VB.CommandButton cmdEND 
      Caption         =   "EXIT"
      Height          =   615
      Left            =   2640
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdCLEAR 
      Caption         =   "CLEAR"
      Height          =   615
      Left            =   2640
      TabIndex        =   0
      Top             =   2160
      Width           =   975
   End
End
Attribute VB_Name = "frmCaluclator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim one, two As Integer
Dim Add, Subtract, Divid, Muti, modules, Power, Ren As Integer

Private Sub cmdADD_Click()
    one = txtONE.Text
    two = txtTWO.Text
    Add = one + two
    
    lblFORMULA.Caption = one & "  +  " & two
    lblEQUAL.Caption = "="
    lblANSWER.Caption = Add
End Sub

Private Sub cmdCLEAR_Click()
    txtONE.Text = ""
    txtTWO.Text = ""
        
    lblFORMULA.Caption = ""
    lblEQUAL.Caption = ""
    lblANSWER.Caption = ""
    End Sub

Private Sub cmdDIVID_Click()
    one = txtONE.Text
    two = txtTWO.Text
    Divid = one / two
    
    lblFORMULA.Caption = one & "  /  " & two
    lblEQUAL.Caption = "="
    lblANSWER.Caption = Divid
End Sub

Private Sub cmdEND_Click()
    Msg = "Are You Sure You Want To Quit ?" ' Creates the mesgbox and what it will state.
    Style = vbYesNo + vbInformation + vbDefaultButton2 ' Create the design of the msg box.
    Title = "Exit" ' Create the title of the msgbox.
             
    Response = MsgBox(Msg, Style, Title) ' states that whatever the user respone.
    If Response = vbYes Then ' if the user chooese the button yes then
       MyString = "Yes"
        End ' exit the program.
    Else
       MyString = "No" ' if the NO button is click simple to the following.
       
    End If
End Sub

Private Sub cmdMOD_Click()
    one = txtONE.Text
    two = txtTWO.Text
    modules = one Mod two
    
    lblFORMULA.Caption = one & "  Mod  " & two
    lblEQUAL.Caption = "="
    lblANSWER.Caption = modules
End Sub

Private Sub cmdMUTIPLY_Click()
    one = txtONE.Text
    two = txtTWO.Text
    Muti = one * two
    
    lblFORMULA.Caption = one & "  *  " & two
    lblEQUAL.Caption = "="
    lblANSWER.Caption = Muti
End Sub

Private Sub cmdPOWER_Click()
    one = txtONE.Text
    two = txtTWO.Text
    Power = one ^ two
    
    lblFORMULA.Caption = one & "  ^  " & two
    lblEQUAL.Caption = "="
    lblANSWER.Caption = Power
End Sub

Private Sub cmdSUBTRACT_Click()
    one = txtONE.Text
    two = txtTWO.Text
    Subtract = one - two
    
    lblFORMULA.Caption = one & "  -  " & two
    lblEQUAL.Caption = "="
    lblANSWER.Caption = Subtract
End Sub

Private Sub Command4_Click()
    one = txtONE.Text
    two = txtTWO.Text
    Ren = one \ two
    
    lblFORMULA.Caption = one & "  \  " & two
    lblEQUAL.Caption = "="
    lblANSWER.Caption = Ren
End Sub

Private Sub Form_Load()
    txtTWO.Enabled = False
    fraOperators.Enabled = False
End Sub

Private Sub txtONE_Change()
    If txtONE.Text = "" Then
           txtTWO.Enabled = False
        Else
           txtTWO.Enabled = True
    End If
End Sub

Private Sub txtONE_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii ' This case statment check that no text can be entered.
            Case vbKeyDelete ' Delete Key
            Case vbKeyBack 'Baskspace key
            Case 48 To 57 ' Number 0-9
            Case Else
            MsgBox "ONLY NUMBERS CAN BE ENTERED, PLEASE ENTER OVER..", vbInformation, "KEY ERROR" ' Error message comes up if you enter text.
            KeyAscii = 0 ' Cancels Keystroke
    End Select
End Sub

Private Sub txtTWO_Change()
    If txtTWO.Text = "" Then
        fraOperators.Enabled = False
    Else
        fraOperators.Enabled = True
    End If
End Sub

Private Sub txtTWO_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii ' This case statment check that no text can be entered.
            Case vbKeyDelete ' Delete Key
            Case vbKeyBack 'Baskspace key
            Case 48 To 57 ' Number 0-9
            Case Else
            MsgBox "ONLY NUMBERS CAN BE ENTERED, PLEASE ENTER OVER..", vbInformation, "KEY ERROR" ' Error message comes up if you enter text.
            KeyAscii = 0 ' Cancels Keystroke
    End Select
End Sub
