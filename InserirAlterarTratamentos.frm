VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form InserirAlterarTratamentos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4485
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5550
   ControlBox      =   0   'False
   Icon            =   "InserirAlterarTratamentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4485
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Height          =   1005
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   2400
      Width           =   3975
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "InserirAlterarTratamentos.frx":521A
      Top             =   3240
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   1
      Top             =   240
      Width           =   3135
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   840
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2280
      TabIndex        =   2
      Top             =   1440
      Width           =   1695
   End
   Begin VB.CommandButton cmdInserir 
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
      Left            =   1200
      TabIndex        =   4
      Top             =   3720
      Width           =   1095
   End
   Begin VB.CommandButton CmdSair 
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
      Left            =   3240
      TabIndex        =   5
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label4 
      Caption         =   "O preço tem de ser introduzido em euros."
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   1800
      Width           =   3375
   End
   Begin VB.Label Label3 
      Caption         =   "Descrição:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label0 
      Caption         =   "Nome do Tratamento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   240
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Código do Tratamento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   2055
   End
   Begin VB.Label Label2 
      Caption         =   "Preço do Tratamento:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1440
      Width           =   1935
   End
End
Attribute VB_Name = "InserirAlterarTratamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim SQL As String
Dim SQL1 As String

Private Sub cmdInserir_Click()
    
    If Len(Trim(Text1.Text)) = 0 Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Nome do Tratamento» é obrigatório."
        FormMsgBoxNormal.Caption = "Tratamentos (Nome)"
        InserirAlterarTratamentos.Enabled = False
        Exit Sub
    End If
    contnome = Len(Text1.Text)
    For i = 1 To contnome
        If IsNumeric(Mid(Text1.Text, i, 1)) Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo «Nome do Tratamento» não pode conter números!"
            FormMsgBoxNormal.Caption = "Tratamentos (Nome)"
            InserirAlterarTratamentos.Enabled = False
            Exit Sub
        End If
    Next i
    
    If Len(Trim(Text3.Text)) = 0 Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Preço do Tratamento» é obrigatório."
        FormMsgBoxNormal.Caption = "Tratamentos (Preço do Tratamento)"
        InserirAlterarTratamentos.Enabled = False
        Exit Sub
    End If
    
    If Len(Trim(Text4.Text)) = 0 Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Descrição» é obrigatório."
        FormMsgBoxNormal.Caption = "Tratamentos (Descrição)"
        InserirAlterarTratamentos.Enabled = False
        Exit Sub
    End If
    
    If cmdInserir.Caption = "&Inserir" Then
        SQL = "INSERT INTO tratamentos(nome_tratamento, cod_Tratamento, preco_tratamento, descricao) " _
                    & "VALUES('" & Text1.Text & "'" _
                        & ", '" & Text2.Text & "'" _
                        & ", '" & Text3.Text & "'" _
                        & ", '" & Text4.Text & "')"
        db.Execute SQL
        
        SQL1 = "INSERT INTO seg_tratamentos(nome_tratamento, cod_Tratamento, preco_tratamento, descricao) " _
                    & "VALUES('" & Text1.Text & "'" _
                        & ", '" & Text2.Text & "'" _
                        & ", '" & Text3.Text & "'" _
                        & ", '" & Text4.Text & "')"
        db.Execute SQL1
        
        FormMsgBoxSimNao.Show
        FormMsgBoxSimNao.Label1.Caption = "O tratamento foi inserido com exito!" & vbCr & "Deseja inserir outro tratamento?"
        FormMsgBoxSimNao.Caption = "Inserir (Tratamentos)"
        InserirAlterarPacientes.Enabled = False

    Else
        SQL = "UPDATE Tratamentos SET preco_tratamento = '" & Text3.Text & "'" _
            & ", descricao = '" & Text4.Text & "'" _
            & " WHERE cod_tratamento = " & PainelControlo.ListTratamentos.SelectedItem.Tag
        db.Execute SQL
        
        SQL1 = "UPDATE seg_Tratamentos SET preco_tratamento = '" & Text3.Text & "'" _
            & ", descricao = '" & Text4.Text & "'" _
            & " WHERE cod_tratamento = " & PainelControlo.ListTratamentos.SelectedItem.Tag
        db.Execute SQL1
        
        ListTratamentos_Ordena_SetUp
        
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "Os dados do tratamento foram alterados com êxito!"
        FormMsgBoxNormal.Caption = "Tratamentos (Alterar)"
        InserirAlterarTratamentos.Enabled = False
    End If
  
End Sub

Private Sub Form_Load()

    Skin1.LoadSkin App.Path & "\Skin\dogmas.skn"
    Skin1.ApplySkin Me.hWnd
    
    Set db = OpenDatabase(App.Path & "\Clinica.mdb")
    
End Sub

Private Sub cmdSair_Click()
    
    PainelControlo.Enabled = True
    ListTratamentos_Ordena_SetUp
    PainelControlo.SetFocus
    PainelControlo.ListTratamentos.SetFocus
    SQL = "select * from tratamentos"
    Set rst = db.OpenRecordset(SQL)
    If Not (rst.BOF = True And rst.EOF = True) Then
        PainelControlo.cmdAlterar5.Enabled = True
        PainelControlo.cmdDel5.Enabled = True
    End If
    Unload Me
    
End Sub
Private Sub ListTratamentos_Ordena_SetUp()

    Dim itemx As ListItem
  
    PainelControlo.ListTratamentos.ListItems.Clear
    
    SQL = " SELECT * " _
        & " FROM Tratamentos" _
        & " ORDER BY nome_tratamento "
        
    Set rst = db.OpenRecordset(SQL)

    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = PainelControlo.ListTratamentos.ListItems.Add(, , rst("cod_tratamento"))
        itemx.SubItems(1) = rst("nome_tratamento")
        itemx.SubItems(2) = rst("preco_tratamento")
        itemx.Tag = rst("cod_Tratamento")
        
        rst.MoveNext
    Wend
    
End Sub

