VERSION 5.00
Begin VB.Form frmTrialCount 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Mike's Trial Creater"
   ClientHeight    =   1680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   5055
   Icon            =   "frmTrialCount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1680
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Register Me! >>"
      Default         =   -1  'True
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Register!"
      Height          =   375
      Left            =   3840
      TabIndex        =   8
      Top             =   1200
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2640
      MaxLength       =   20
      TabIndex        =   7
      Text            =   "9873797-9FD9FD9E3"
      Top             =   720
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   3120
      MaxLength       =   15
      TabIndex        =   5
      Text            =   "Mike Canejo"
      Top             =   240
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Count Trial Up"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Reset Trial"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1455
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Code#:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   6
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
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
      TabIndex        =   2
      Top             =   0
      Width           =   1455
   End
End
Attribute VB_Name = "frmTrialCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Function TrialTime(TheForm As Form, TrialOverMSG As String, TrialOverMSGTitle As String, TrialOverMSGType As String, TrialCount As Integer, Work As Boolean)

    If Not Work Then SaveSetting TheForm.Name, "Trial", "TimesOpen", ".": End
'If Work = False then reset trial to 0 if Work = True then Count up the Trial

    SaveSetting TheForm.Name, "Trial", "TimesOpen", Val(GetSetting(TheForm.Name, "Trial", "TimesOpen")) + 1
'Write + 1 to the last to the last time opened

    If GetSetting(TheForm.Name, "Trial", "TimesOpen") > TrialCount Then SaveSetting TheForm.Name, "Trial", "TimesOpen", TrialCount: MsgBox TrialOverMSG, TrialOverMSGType, TrialOverMSGTitle: End
'If the amount of times open is > then the TrialCount..
'Reset it to the number in TrialCount specified
'Display a message and terminate the program
End Function

Private Sub Command1_Click()
    SaveSetting Me.Name, "Trial", "TimesOpen", 0
'Resets the trial
    Label1.Caption = 0
'Resets the Label
End Sub

Private Sub Command2_Click()
    TrialTime Me, "The trial of " & Me.Caption & " has expired. Please register this product to get the full version.", "Trial Expired", vbCritical, 50, True
'Activates the trial counter. True to count up and False to reset the Trial count
    Label1.Caption = GetSetting(Me.Name, "Trial", "TimesOpen")
'Display times open
End Sub

Private Sub Command3_Click()
If Command3.Tag = "" Then
    Command3.Tag = "1"
    Command3.Caption = "Register Me! <<"
    Dim x As Integer: For x = 1850 To 5145: Me.Width = x: Next x
'Makes width 5145 to show register part
Exit Sub
Else
    Command3.Tag = ""
    x = 5145: Do: Me.Width = Me.Width - 1: Loop Until Me.Width = 1860
    Command3.Caption = "Register Me! >>"
End If
'Makes the forms width back to normal
End Sub

Private Sub Command4_Click()
    If Text1 = "Mike Canejo" And Text2 = "9873797-9FD9FD9E3" Then
        MsgBox "The name and code you entered was Correct!", vbInformation
        TrialTime Me, "", "", "", 0, False
    Else
        MsgBox "The name and code you entered was In-Correct!", vbExclamation
    End If
    
'Example! this checks for the name and code and disables the Trial Count to register it
End Sub

Private Sub Form_Load()
'Put: TrialTime Me, "The trial of " & Me.Caption & " has expired. Please register this product to get the full version.", "Trial Expired", vbCritical, 50, True
'Here on Form_Load to activate the Trial Counter when the application loads
    Label1.Caption = GetSetting(Me.Name, "Trial", "TimesOpen")
'Display trial count
    Me.Width = 1860
'Make forms width 1860
    If GetSetting(Me.Name, "Trial", "TimesOpen") = "." Then
    Label1.FontSize = 8: Label1.Caption = "Registered!"
    MsgBox "Thanks for registering.", vbInformation
    'this means they registered the program.
    'put the full version code here
    'Example: me.unload  and frmFullVersion.show
    End If
    
End Sub
