VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form NovoUtilizador 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4635
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   6150
   ControlBox      =   0   'False
   Icon            =   "Novo Utilizador.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   6150
   StartUpPosition =   2  'CenterScreen
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
      Left            =   1800
      TabIndex        =   1
      Top             =   480
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
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1440
      Width           =   2535
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
      Left            =   3480
      TabIndex        =   6
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Novo Utilizador"
      Height          =   495
      Left            =   1560
      TabIndex        =   5
      Top             =   4080
      Width           =   1335
   End
   Begin VB.TextBox Text3 
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
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2400
      Width           =   2535
   End
   Begin VB.TextBox Text4 
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
      Left            =   1800
      PasswordChar    =   "*"
      TabIndex        =   4
      Top             =   3360
      Width           =   2535
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "Novo Utilizador.frx":521A
      Top             =   240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Utilizador"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
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
      Left            =   2520
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Confirme a Password"
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
      Left            =   1920
      TabIndex        =   7
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Password do Sistema"
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
      Left            =   1920
      TabIndex        =   0
      Top             =   3000
      Width           =   2415
   End
End
Attribute VB_Name = "NovoUtilizador"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim SQL As String

Private Sub Command2_Click()
    
    If NovoUtilizador.Caption = "Novo Utilizador (Login)" Then
        Login.Show
        Unload Me
    ElseIf NovoUtilizador.Caption = "Novo Utilizador (GestiMed)" Then
        PainelControlo.Enabled = True
        PainelControlo.SetFocus
        Unload Me
    End If
    
End Sub

Private Sub Command1_Click()

    SQL = "select * from login" _
        & " WHERE nome_utilizador LIKE " & "'" & Text1.Text & "'" & ""
    Set rst = db.OpenRecordset(SQL)
    
    If Not (rst.BOF = True And rst.EOF = True) Then
        FormMsgBoxNormal.Caption = "Novo Utilizador (Utilizador)"
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O nome de utilizador que escolheu já existe!"
        NovoUtilizador.Enabled = False
        Exit Sub
    ElseIf Text2.Text <> Text3.Text Then
        FormMsgBoxNormal.Caption = "Novo Utilizador (Password)"
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "A password que escreveu está errada!" & vbCr & "Escreva a password novamente."
        NovoUtilizador.Enabled = False
        Exit Sub
    ElseIf Text4.Text <> "GestiMedSystem" Then
        FormMsgBoxNormal.Caption = "Novo Utilizador (Password do Sistema)"
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "A password do sistema que escreveu está errada!" & vbCr & "Escreva a password do sistema novamente."
        NovoUtilizador.Enabled = False
        Exit Sub
    Else
        SQL = "INSERT INTO login(nome_utilizador,password)" _
                & "VALUES('" & Text1.Text & "'" _
                    & ", '" & Text2.Text & "')"
        db.Execute SQL
        If NovoUtilizador.Caption = "Novo Utilizador (Login)" Then
            FormMsgBoxNormal.Caption = "Novo Utilizador (Login)"
        ElseIf NovoUtilizador.Caption = "Novo Utilizador (GestiMed)" Then
            FormMsgBoxNormal.Caption = "Novo Utilizador (GestiMed)"
        End If
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "Utilizador inserido com êxito!"
        NovoUtilizador.Enabled = False
        Exit Sub
    End If
    
End Sub

Private Sub Form_Load()

    Skin1.LoadSkin App.Path & "\Skin\dogmas.skn"
    Skin1.ApplySkin Me.hWnd
    Set db = OpenDatabase(App.Path & "\clinica.mdb")

End Sub

