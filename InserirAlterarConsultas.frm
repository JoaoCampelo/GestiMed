VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Begin VB.Form InserirAlterarConsultas 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   5790
   ClientLeft      =   15
   ClientTop       =   -15
   ClientWidth     =   5550
   ControlBox      =   0   'False
   Icon            =   "InserirAlterarConsultas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   5550
   StartUpPosition =   2  'CenterScreen
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   0
      OleObjectBlob   =   "InserirAlterarConsultas.frx":521A
      Top             =   5280
   End
   Begin VB.ComboBox cmbMedicos 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1800
      TabIndex        =   4
      Text            =   "Selecione o Médico"
      Top             =   3840
      Width           =   3495
   End
   Begin VB.ComboBox cmbPacientes 
      CausesValidation=   0   'False
      Enabled         =   0   'False
      Height          =   315
      Left            =   2040
      TabIndex        =   3
      Text            =   "Selecione o Paciente"
      Top             =   1440
      Width           =   3255
   End
   Begin VB.TextBox Text10 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1080
      TabIndex        =   21
      Top             =   3240
      Width           =   4215
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
      TabIndex        =   6
      Top             =   5040
      Width           =   1095
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
      TabIndex        =   5
      Top             =   5040
      Width           =   1095
   End
   Begin VB.TextBox Text9 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   19
      Text            =   " "
      Top             =   4440
      Width           =   1815
   End
   Begin VB.TextBox Text7 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3840
      TabIndex        =   18
      Top             =   2640
      Width           =   1455
   End
   Begin VB.TextBox Text6 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1200
      TabIndex        =   17
      Top             =   2640
      Width           =   1335
   End
   Begin VB.TextBox Text5 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2160
      TabIndex        =   16
      Top             =   2040
      Width           =   3135
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   3480
      TabIndex        =   2
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   840
      TabIndex        =   1
      Top             =   840
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2040
      TabIndex        =   15
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label9 
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
      TabIndex        =   20
      Top             =   3240
      Width           =   855
   End
   Begin VB.Label Label8 
      Caption         =   "Código do Médico:"
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
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label Label7 
      Caption         =   "Nome do Médico:"
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
      TabIndex        =   13
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Telemóvel:"
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
      TabIndex        =   12
      Top             =   2640
      Width           =   975
   End
   Begin VB.Label Label5 
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
      TabIndex        =   11
      Top             =   2640
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "Código do Paceinte:"
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
      Top             =   2040
      Width           =   1815
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
      TabIndex        =   9
      Top             =   840
      Width           =   615
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
      TabIndex        =   8
      Top             =   840
      Width           =   615
   End
   Begin VB.Label Label0 
      Caption         =   "Código da Consulta:"
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
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label3 
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
      TabIndex        =   0
      Top             =   1440
      Width           =   1695
   End
End
Attribute VB_Name = "InserirAlterarConsultas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim SQL As String
Dim pacientes As String
Dim medicos As String

Private Sub cmbMedicos_Click()

    medicos = cmbMedicos.Text
    SQL = " SELECT * FROM medicos" _
        & " WHERE nome_medico LIKE " & "'" & medicos & "'"
        
    Set rst = db.OpenRecordset(SQL)
    
    Text9.Text = rst("cod_medico")
    
End Sub

Private Sub cmbPacientes_Click()

    pacientes = cmbPacientes.Text
    SQL = " SELECT * FROM pacientes" _
        & " WHERE nome_paciente LIKE " & "'" & pacientes & "'"
        
    Set rst = db.OpenRecordset(SQL)
    
    Text5.Text = rst("cod_paciente")
   
    If rst("telefone") = Null Then
        Text6.Text = rst("telefone")
    Else
        x = rst("telefone")
        Text6.Text = "" & x
    End If
    
    If rst("telemovel") = Null Then
        Text7.Text = rst("telemovel")
    Else
        x = rst("telemovel")
        Text7.Text = "" & x
    End If
    
    If rst("email") = Null Then
        Text10.Text = rst("email")
    Else
        x = rst("email")
        Text10.Text = "" & x
    End If
    
End Sub

Private Sub cmdInserir_Click()
       
    If Not IsDate(Text2.Text) Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Data» tem que ser uma data (dd/mm/aa)!"
        FormMsgBoxNormal.Caption = "Consultas (Data)"
        InserirAlterarConsultas.Enabled = False
        Exit Sub
    End If
 
    If Not IsDate(Text3.Text) Then
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "O campo «Hora» tem que ser uma hora (hh:mm:ss)!"
        FormMsgBoxNormal.Caption = "Consultas (Hora)"
        InserirAlterarConsultas.Enabled = False
        Exit Sub
    End If
    
    If cmdInserir.Caption = "&Inserir" Then
        If cmbPacientes.ListIndex = -1 Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo «Nome do Paciente» é obrigatório."
            FormMsgBoxNormal.Caption = "Consultas (Nome do Paciente)"
            InserirAlterarConsultas.Enabled = False
            Exit Sub
        End If
        
        If cmbMedicos.ListIndex = -1 Then
            FormMsgBoxNormal.Show
            FormMsgBoxNormal.Label1.Caption = "O campo «Nome do Médico» é obrigatório."
            FormMsgBoxNormal.Caption = "Consultas (Nome do Médico)"
            InserirAlterarConsultas.Enabled = False
            Exit Sub
        End If
        
        SQL = "INSERT INTO Consultas(cod_cunsulta,data,hora,cod_paciente,cod_medico)" _
                    & "VALUES('" & Text1.Text & "'" _
                        & ", '" & Text2.Text & "'" _
                        & ", '" & Text3.Text & "'" _
                        & ", '" & Text5.Text & "'" _
                        & ", '" & Text9.Text & "')"
        db.Execute SQL
        
        FormMsgBoxSimNao.Show
        FormMsgBoxSimNao.Label1.Caption = "A consulta foi inserida com exito!" & vbCr & "Deseja inserir outra?"
        FormMsgBoxSimNao.Caption = "Inserir (Consultas)"
        InserirAlterarConsultas.Enabled = False
        
    Else
        SQL = "UPDATE Consultas SET cod_cunsulta = '" & Text1.Text & "'" _
            & ", data = '" & Text2.Text & "'" _
            & ", hora = '" & Text3.Text & "'" _
            & " WHERE cod_cunsulta = " & PainelControlo.ListConsultas.SelectedItem.Tag
        db.Execute SQL
               
        ListConsultas_Ordena_SetUp
        
        FormMsgBoxNormal.Show
        FormMsgBoxNormal.Label1.Caption = "Os dados da consulta foram alterados com êxito!"
        FormMsgBoxNormal.Caption = "Consultas (Alterar)"
        InserirAlterarConsultas.Enabled = False
    End If
    
End Sub

Private Sub cmdSair_Click()
        
    PainelControlo.Enabled = True
    ListConsultas_Ordena_SetUp
    PainelControlo.SetFocus
    PainelControlo.ListConsultas.SetFocus
    SQL = "select * from consultas"
    Set rst = db.OpenRecordset(SQL)
    If Not (rst.BOF = True And rst.EOF = True) Then
        PainelControlo.cmdAlterar.Enabled = True
        PainelControlo.cmdDel.Enabled = True
    End If
    Unload Me
    
End Sub

Private Sub Form_Load()

    Skin1.LoadSkin App.Path & "\Skin\dogmas.skn"
    Skin1.ApplySkin Me.hWnd
    
    Set db = OpenDatabase(App.Path & "\Clinica.mdb")
    
End Sub

Private Sub ListConsultas_Ordena_SetUp()
    
    PainelControlo.ListConsultas.ListItems.Clear
    
    SQL = " SELECT consultas.cod_cunsulta, consultas.data, consultas.hora, pacientes.nome_paciente " _
        & " FROM consultas, pacientes" _
        & " WHERE consultas.cod_paciente=pacientes.cod_paciente" _
        & " ORDER BY data desc"
    Set rst = db.OpenRecordset(SQL)

    If rst.BOF = True And rst.EOF = True Then Exit Sub
    While Not rst.EOF
        Set itemx = PainelControlo.ListConsultas.ListItems.Add(, , rst("cod_cunsulta"))
        itemx.SubItems(1) = rst("data")
        itemx.SubItems(2) = rst("hora")
        itemx.SubItems(3) = rst("nome_paciente")
        itemx.Tag = rst("cod_cunsulta")
        rst.MoveNext
    Wend
    
End Sub
