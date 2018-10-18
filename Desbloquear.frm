VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Desbloquear 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   2190
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   4230
   ControlBox      =   0   'False
   Icon            =   "Desbloquear.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2190
   ScaleWidth      =   4230
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "Desbloquear.frx":521A
      Top             =   1560
   End
   Begin VB.CommandButton cmdDesbloquear 
      Caption         =   "Desbloquear"
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   1440
      TabIndex        =   1
      Top             =   240
      Width           =   2535
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      IMEMode         =   3  'DISABLE
      Left            =   1440
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   960
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Utilizador:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
End
Attribute VB_Name = "Desbloquear"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim SQL As String

Private Sub cmdDesbloquear_Click()
   
    SQL = " SELECT * FROM login" _
        & " WHERE nome_utilizador LIKE " & "'" & Text1.Text & "'" & " AND password LIKE " & "'" & Text2.Text & "'"
    Set rst = db.OpenRecordset(SQL)
    
    If rst.EOF = True Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "Utilizador/Password Inválido(s)!"
        FormMsgBoxNormal.Caption = "Desbloquear"
        Desbloquear.Enabled = False
    Else
        If Text1.Text = rst("nome_utilizador") And Text2.Text = rst("password") Then
            PainelControlo.Enabled = True
            Unload Me
        Else
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "Utilizador/Password Inválido(s)!"
            FormMsgBoxNormal.Caption = "Desbloquear"
            Desbloquear.Enabled = False
        End If
    End If
    
End Sub

Private Sub Form_Load()

    Skin1.LoadSkin App.Path & "\Skin\dogmas.skn"
    Skin1.ApplySkin Me.hWnd

    Set db = OpenDatabase(App.Path & "\clinica.mdb")
    
End Sub


