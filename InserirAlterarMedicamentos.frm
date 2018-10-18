VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form InserirAlterarMedicamentos 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   6150
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5550
   ControlBox      =   0   'False
   Icon            =   "InserirAlterarMedicamentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   120
      OleObjectBlob   =   "InserirAlterarMedicamentos.frx":521A
      Top             =   5160
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2400
      TabIndex        =   5
      Top             =   3600
      Width           =   1935
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   4
      Top             =   3000
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   3
      Top             =   2400
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Height          =   1005
      Left            =   1320
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Top             =   4200
      Width           =   3975
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   0
      Top             =   240
      Width           =   3015
   End
   Begin VB.TextBox Text2 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2400
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2400
      TabIndex        =   2
      Top             =   1440
      Width           =   1935
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
      TabIndex        =   7
      Top             =   5400
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
      TabIndex        =   8
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label7 
      Caption         =   "Embalagens Disponiveis:"
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
      TabIndex        =   16
      Top             =   3600
      Width           =   2175
   End
   Begin VB.Label Label6 
      Caption         =   "Tipo de Medicamento:"
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
      TabIndex        =   15
      Top             =   3000
      Width           =   1935
   End
   Begin VB.Label Label5 
      Caption         =   "Comprimidos por Caixa:"
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
      TabIndex        =   14
      Top             =   2400
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "O preço tem de ser introduzido em euros."
      Height          =   255
      Left            =   240
      TabIndex        =   13
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
      TabIndex        =   12
      Top             =   4560
      Width           =   975
   End
   Begin VB.Label Label0 
      Caption         =   "Nome do Medicamento:"
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
      TabIndex        =   11
      Top             =   240
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Código do Medicamento:"
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
      TabIndex        =   10
      Top             =   840
      Width           =   2175
   End
   Begin VB.Label Label2 
      Caption         =   "Preço do Medicamento:"
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
      Top             =   1440
      Width           =   2055
   End
End
Attribute VB_Name = "InserirAlterarMedicamentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim SQL As String

Private Sub cmdInserir_Click()
    
    If Len(Trim(Text1.Text)) = 0 Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Nome do Medicamento» é obrigatório."
        FormMsgBoxNormal.Caption = "Medicamentos (Nome)"
        InserirAlterarMedicamentos.Enabled = False
        Exit Sub
    End If
    
    If Len(Trim(Text3.Text)) = 0 Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Preço do Medicamento» é obrigatório."
        FormMsgBoxNormal.Caption = "Medicamentos (Preço do Medicamento)"
        InserirAlterarMedicamentos.Enabled = False
        Exit Sub
    End If
    
    If Text4.Text <> "" Then
        contquant = Len(Text4.Text)
        For i = 1 To contquant
            If Not IsNumeric(Mid(Text4.Text, i, 1)) Then
                FormMsgBoxNormal.Show
                FormMsgBoxNormal.Label1.Caption = "O campo «Comprimidos por Caixa» não pode conter letras!"
                FormMsgBoxNormal.Caption = "Medicamentos (Comprimidos por Caixa)"
                InserirAlterarMedicamentos.Enabled = False
                Exit Sub
            End If
        Next i
    End If
    
    If Len(Trim(Text5.Text)) = 0 Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Tipo de Medicamento» é obrigatório."
        FormMsgBoxNormal.Caption = "Medicamentos (Tipo de Medicamento)"
        InserirAlterarMedicamentos.Enabled = False
        Exit Sub
    End If
    contcaixa = Len(Text5.Text)
    For i = 1 To contcaixa
        If IsNumeric(Mid(Text5.Text, i, 1)) Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo «Tipo de Medicamento» não pode conter números!"
            FormMsgBoxNormal.Caption = "Medicamentos (Tipo de Medicamento)"
            InserirAlterarMedicamentos.Enabled = False
            Exit Sub
        End If
    Next i
    
    If Len(Trim(Text6.Text)) = 0 Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Embalagens Disponiveis» é obrigatório."
        FormMsgBoxNormal.Caption = "Medicamentos (Embalagens Disponiveis)"
        InserirAlterarMedicamentos.Enabled = False
        Exit Sub
    End If
    contstock = Len(Text6.Text)
    For i = 1 To contstock
        If Not IsNumeric(Mid(Text6.Text, i, 1)) Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo «Embalagens Disponiveis» não pode conter letras!"
            FormMsgBoxNormal.Caption = "Medicamentos (Embalagens Disponiveis)"
            InserirAlterarMedicamentos.Enabled = False
            Exit Sub
        End If
    Next i
    
    If Len(Trim(Text7.Text)) = 0 Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Descrição» é obrigatório."
        FormMsgBoxNormal.Caption = "Medicamentos (Descrição)"
        InserirAlterarMedicamentos.Enabled = False
        Exit Sub
    End If
    
    If cmdInserir.Caption = "&Inserir" Then
        SQL = "INSERT INTO Medicamentos(nome, cod_Medicamento, preco_Medicamento, comprimidos_caixa, tipo_medicamento, emblagem_disponiveis, descricao) " _
                    & "VALUES('" & Text1.Text & "'" _
                        & ", '" & Text2.Text & "'" _
                        & ", '" & Text3.Text & "'" _
                        & ", '" & Text4.Text & "'" _
                        & ", '" & Text5.Text & "'" _
                        & ", '" & Text6.Text & "'" _
                        & ", '" & Text7.Text & "')"
        db.Execute SQL
        
        FormMsgBoxSimNao.Show
        FormMsgBoxSimNao.Label1.Caption = "O medicamento foi inserido com exito!" & vbCr & "Deseja inserir outro medicamento?"
        FormMsgBoxSimNao.Caption = "Inserir (Medicamentos)"
        InserirAlterarMedicamentos.Enabled = False

    Else
        SQL = "UPDATE Medicamentos SET preco_Medicamento = '" & Text3.Text & "'" _
            & ", emblagem_disponiveis = '" & Text6.Text & "'" _
            & ", descricao = '" & Text7.Text & "'" _
            & " WHERE cod_Medicamento = " & PainelControlo.ListMedicamentos.SelectedItem.Tag
        db.Execute SQL
        
        ListMedicamentos_Ordena_SetUp
        
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "Os dados do medicamento foram alterados com êxito!"
        FormMsgBoxNormal.Caption = "Medicamentos (Alterar)"
        InserirAlterarMedicamentos.Enabled = False
    End If
  
End Sub

Private Sub Form_Load()

    Skin1.LoadSkin App.Path & "\Skin\dogmas.skn"
    Skin1.ApplySkin Me.hWnd
    
    Set db = OpenDatabase(App.Path & "\Clinica.mdb")
    
End Sub

Private Sub cmdSair_Click()
    
    PainelControlo.Enabled = True
    ListMedicamentos_Ordena_SetUp
    PainelControlo.SetFocus
    PainelControlo.ListMedicamentos.SetFocus
    SQL = "select * from medicamentos"
    Set rst = db.OpenRecordset(SQL)
    If Not (rst.BOF = True And rst.EOF = True) Then
        PainelControlo.cmdAlterar6.Enabled = True
        PainelControlo.cmdDel6.Enabled = True
        PainelControlo.cmdVender.Enabled = True
    End If
    Unload Me
    
End Sub
Private Sub ListMedicamentos_Ordena_SetUp()

    Dim itemx As ListItem
  
    PainelControlo.ListMedicamentos.ListItems.Clear
    
    SQL = " SELECT * " _
        & " FROM Medicamentos" _
        & " ORDER BY nome "
        
    Set rst = db.OpenRecordset(SQL)

    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = PainelControlo.ListMedicamentos.ListItems.Add(, , rst("cod_Medicamento"))
        itemx.SubItems(1) = rst("nome")
        itemx.SubItems(2) = rst("preco_Medicamento")
        itemx.Tag = rst("cod_Medicamento")
        
        rst.MoveNext
    Wend
    
End Sub

