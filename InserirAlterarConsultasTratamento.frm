VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form InserirAlterarConsultasTratamento 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7695
   ClientLeft      =   15
   ClientTop       =   15
   ClientWidth     =   5550
   ControlBox      =   0   'False
   Icon            =   "InserirAlterarConsultasTratamento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7695
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text4 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2280
      TabIndex        =   8
      Top             =   2040
      Width           =   3015
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   0
      Top             =   2640
      Width           =   3135
   End
   Begin VB.ComboBox cmbTratamento 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   315
      Left            =   2280
      TabIndex        =   3
      Text            =   "Selecione o Tratamento"
      Top             =   1440
      Width           =   3015
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   9
      Top             =   240
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Top             =   3840
      Width           =   3135
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   13
      Top             =   4440
      Width           =   1335
   End
   Begin VB.TextBox Text8 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   12
      Top             =   4440
      Width           =   1455
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   11
      Text            =   " "
      Top             =   6240
      Width           =   1815
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
      TabIndex        =   6
      Top             =   6840
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
      TabIndex        =   7
      Top             =   6840
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   10
      Top             =   5040
      Width           =   4215
   End
   Begin VB.ComboBox cmbPacientes 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "InserirAlterarConsultasTratamento.frx":521A
      Left            =   2040
      List            =   "InserirAlterarConsultasTratamento.frx":521C
      TabIndex        =   4
      Text            =   "Selecione o Paciente"
      Top             =   3240
      Width           =   3255
   End
   Begin VB.ComboBox cmbMedicos 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      TabIndex        =   5
      Text            =   "Selecione o M�dico"
      Top             =   5640
      Width           =   3495
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "InserirAlterarConsultasTratamento.frx":521E
      Top             =   7080
   End
   Begin VB.Label Label4 
      Caption         =   "C�digo do Tratamento:"
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
      TabIndex        =   27
      Top             =   2040
      Width           =   2055
   End
   Begin VB.Label Label5 
      Caption         =   "Pre�o do Tratamento:"
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
      TabIndex        =   26
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label3 
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
      TabIndex        =   25
      Top             =   1440
      Width           =   1935
   End
   Begin VB.Label Label6 
      Caption         =   "Nome do Paciente:"
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
      TabIndex        =   24
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label Label0 
      Caption         =   "C�digo da Consulta:"
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
      TabIndex        =   23
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Data:"
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
      TabIndex        =   22
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Hora:"
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
      Left            =   2880
      TabIndex        =   21
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label7 
      Caption         =   "C�digo do Paceinte:"
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
      TabIndex        =   20
      Top             =   3840
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Telefone:"
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
      TabIndex        =   19
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label9 
      Caption         =   "Telem�vel:"
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
      Left            =   2760
      TabIndex        =   18
      Top             =   4440
      Width           =   975
   End
   Begin VB.Label Label11 
      Caption         =   "Nome do M�dico:"
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
      TabIndex        =   17
      Top             =   5640
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "C�digo do M�dico:"
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
      Top             =   6240
      Width           =   1815
   End
   Begin VB.Label Label10 
      Caption         =   "E - Mail:"
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
      Top             =   5040
      Width           =   855
   End
End
Attribute VB_Name = "InserirAlterarConsultasTratamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim SQL As String
Dim tratamento As String
Dim pacientes As String
Dim medicos As String

Private Sub cmbMedicos_Click()

    medicos = cmbMedicos.Text
    SQL = " SELECT * FROM medicos" _
        & " WHERE nome_medico LIKE " & "'" & medicos & "'"
        
    Set rst = db.OpenRecordset(SQL)
    
    Text10.Text = rst("cod_medico")
    
End Sub

Private Sub cmbPacientes_Click()

    pacientes = cmbPacientes.Text
    SQL = " SELECT * FROM pacientes" _
        & " WHERE nome_paciente LIKE " & "'" & pacientes & "'"
        
    Set rst = db.OpenRecordset(SQL)
    
    Text6.Text = rst("cod_paciente")
   
    If rst("telefone") = Null Then
        Text7.Text = rst("telefone")
    Else
        x = rst("telefone")
        Text7.Text = "" & x
    End If
    
    If rst("telemovel") = Null Then
        Text8.Text = rst("telemovel")
    Else
        x = rst("telemovel")
        Text8.Text = "" & x
    End If
    
    If rst("email") = Null Then
        Text9.Text = rst("email")
    Else
        x = rst("email")
        Text9.Text = "" & x
    End If
    
End Sub

Private Sub cmbTratamento_Click()

    tratamento = cmbTratamento.Text
    SQL = " SELECT * FROM Tratamentos" _
        & " WHERE nome_tratamento LIKE " & "'" & tratamento & "'"
        
    Set rst = db.OpenRecordset(SQL)
    
    Text4.Text = rst("cod_tratamento")
    Text5.Text = rst("preco_tratamento")
    
End Sub

Private Sub cmdInserir_Click()

    If Not IsDate(Text2.Text) Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo �Data� tem que ser uma data (dd/mm/aa)!"
        FormMsgBoxNormal.Caption = "Consultas de Tratamento (Data)"
        InserirAlterarConsultas.Enabled = False
        Exit Sub
    End If
 
    If Not IsDate(Text3.Text) Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo �Hora� tem que ser uma hora (hh:mm:ss)!"
        FormMsgBoxNormal.Caption = "Consultas de Tratamento (Hora)"
        InserirAlterarConsultas.Enabled = False
        Exit Sub
    End If
    
    If cmdInserir.Caption = "&Inserir" Then
        If cmbTratamento.ListIndex = -1 Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo �Nome do Tratamento� � obrigat�rio."
            FormMsgBoxNormal.Caption = "Consultas de Tratamento (Nome do Tratamento)"
            InserirAlterarConsultas.Enabled = False
            Exit Sub
            cmbTratamento.SetFocus
        End If
        
        If cmbPacientes.ListIndex = -1 Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo �Nome do Paciente� � obrigat�rio."
            FormMsgBoxNormal.Caption = "Consultas de Tratamento (Nome do Paciente)"
            InserirAlterarConsultas.Enabled = False
            Exit Sub
            cmbPacientes.SetFocus
        End If
        
        If cmbMedicos.ListIndex = -1 Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo �Nome do M�dico� � obrigat�rio."
            FormMsgBoxNormal.Caption = "Consultas de Tratamento (Nome do M�dico)"
            InserirAlterarConsultas.Enabled = False
            Exit Sub
            cmbMedicos.SetFocus
        End If
        
        SQL = "INSERT INTO Consultas_Tratamentos(cod_consultatratamento,data,hora,cod_tratamento,cod_paciente,cod_medico)" _
                    & "VALUES('" & Text1.Text & "'" _
                        & ", '" & Text2.Text & "'" _
                        & ", '" & Text3.Text & "'" _
                        & ", '" & Text4.Text & "'" _
                        & ", '" & Text6.Text & "'" _
                        & ", '" & Text10.Text & "')"
        db.Execute SQL
        
        FormMsgBoxSimNao.Show
        FormMsgBoxSimNao.Label1.Caption = "A consulta foi inserida com exito!" & vbCr & "Deseja inserir outra?"
        FormMsgBoxSimNao.Caption = "Inserir (Consultas de Tratamento)"
        InserirAlterarConsultasTratamento.Enabled = False
        
    Else
        SQL = "UPDATE Consultas_Tratamentos SET cod_consultatratamento = '" & Text1.Text & "'" _
            & ", data = '" & Text2.Text & "'" _
            & ", hora = '" & Text3.Text & "'" _
            & " WHERE cod_consultatratamento = " & PainelControlo.ListConsultasTratamento.SelectedItem.Tag
        db.Execute SQL
        
        ListConsultasTratamentos_Ordena_SetUp
               
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "Os dados da consulta foram alterados com �xito!"
        FormMsgBoxNormal.Caption = "Consultas de Tratamento (Alterar)"
        InserirAlterarConsultas.Enabled = False
    End If
    
End Sub

Private Sub cmdSair_Click()
        
    PainelControlo.Enabled = True
    ListConsultasTratamentos_Ordena_SetUp
    PainelControlo.SetFocus
    PainelControlo.ListConsultasTratamento.SetFocus
    SQL = "select * from consultas_tratamentos"
    Set rst = db.OpenRecordset(SQL)
    If Not (rst.BOF = True And rst.EOF = True) Then
        PainelControlo.cmdAlterar2.Enabled = True
        PainelControlo.cmdDel2.Enabled = True
    End If
    Unload Me
    
End Sub

Private Sub Form_Load()

    Skin1.LoadSkin App.Path & "\Skin\dogmas.skn"
    Skin1.ApplySkin Me.hWnd
    
    Set db = OpenDatabase(App.Path & "\Clinica.mdb")
    
End Sub

Private Sub ListConsultasTratamentos_Ordena_SetUp()

    Dim itemx As ListItem

    PainelControlo.ListConsultasTratamento.ListItems.Clear
    
    SQL = " SELECT consultas_tratamentos.cod_consultatratamento, consultas_tratamentos.data, consultas_tratamentos.hora, consultas_tratamentos.cod_tratamento, pacientes.nome_paciente " _
        & " FROM consultas_tratamentos, pacientes" _
        & " WHERE consultas_tratamentos.cod_paciente=pacientes.cod_paciente" _
        & " ORDER BY consultas_tratamentos.data desc"
        
    Set rst = db.OpenRecordset(SQL)

    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = PainelControlo.ListConsultasTratamento.ListItems.Add(, , rst("cod_consultatratamento"))
        itemx.SubItems(1) = rst("data")
        itemx.SubItems(2) = rst("hora")
        itemx.SubItems(3) = rst("cod_tratamento")
        itemx.SubItems(4) = rst("nome_paciente")
        itemx.Tag = rst("cod_consultatratamento")
        
        rst.MoveNext
    Wend
    
End Sub

