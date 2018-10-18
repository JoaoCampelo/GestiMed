VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form Login 
   BackColor       =   &H8000000E&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5895
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6150
   ControlBox      =   0   'False
   Icon            =   "Login.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Login.frx":521A
   ScaleHeight     =   5895
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Novo Utilizador"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   5280
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "Login.frx":D013
      Top             =   0
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3960
      TabIndex        =   4
      Top             =   5280
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Confirmar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1080
      TabIndex        =   3
      Top             =   5280
      Width           =   1095
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
      Left            =   2160
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   4440
      Width           =   2535
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
      Left            =   2160
      TabIndex        =   1
      Top             =   3720
      Width           =   2535
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
      Left            =   960
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
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
      Left            =   960
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
End
Attribute VB_Name = "Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim SQL As String

Private Sub Command1_Click()
   
    SQL = " SELECT * FROM login" _
        & " WHERE nome_utilizador LIKE " & "'" & Text1.Text & "'" & " AND password LIKE " & "'" & Text2.Text & "'"
    Set rst = db.OpenRecordset(SQL)
    
    If rst.EOF = True Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "Utilizador/Password Inválido(s)!"
        FormMsgBoxNormal.Caption = "Login"
        Login.Enabled = False
    Else
        If Text1.Text = rst("nome_utilizador") And Text2.Text = rst("password") Then
            PainelControlo.Show
            PainelControlo.ListConsultas.SetFocus
            SQL = "select * from medicamentos"
            Set rst = db.OpenRecordset(SQL)
            While Not rst.EOF
                If rst("emblagem_disponiveis") <= 10 Then
                    FormMsgBoxNormal.Show
                    FormMsgBoxNormal.Caption = "Stock a Esgotar"
                    FormMsgBoxNormal.Label1.Caption = "Existem medicamentos cujo stock se está a esgotar!"
                    PainelControlo.Enabled = False
                    Exit Sub
                End If
                rst.MoveNext
            Wend
            Unload Me
        Else
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "Utilizador/Password Inválido(s)!"
            FormMsgBoxNormal.Caption = "Login"
            Login.Enabled = False
        End If
    End If
    
End Sub

Private Sub Command2_Click()

    End
    
End Sub

Private Sub Command3_Click()
    
    NovoUtilizador.Show
    NovoUtilizador.Caption = "Novo Utilizador (Login)"
    Unload Me
    
End Sub

Private Sub Form_Load()

    Skin1.LoadSkin App.Path & "\Skin\dogmas.skn"
    Skin1.ApplySkin Me.hWnd
    Set db = OpenDatabase(App.Path & "\clinica.mdb")

End Sub

