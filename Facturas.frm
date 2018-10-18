VERSION 5.00
Object = "{74848F95-A02A-4286-AF0C-A3C755E4A5B3}#1.0#0"; "actskn43.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Facturas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Facturas"
   ClientHeight    =   4755
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10830
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4755
   ScaleWidth      =   10830
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdRepor 
      Caption         =   "Repor"
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmdProcurar 
      Caption         =   "Procurar"
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   4200
      Width           =   1335
   End
   Begin ACTIVESKINLibCtl.Skin Skin1 
      Left            =   240
      OleObjectBlob   =   "Facturas.frx":0000
      Top             =   5640
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5160
      TabIndex        =   0
      Top             =   4200
      Width           =   1335
   End
   Begin MSComctlLib.ListView ListFacturas 
      Height          =   3375
      Left            =   0
      TabIndex        =   1
      Top             =   120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   5953
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label1 
      Caption         =   "Os preços acima apresentados encontram-se em euros."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   4095
   End
End
Attribute VB_Name = "Facturas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim db As DAO.Database
Dim rst As DAO.Recordset
Dim SQL As String

Private Sub cmdOK_Click()
    
    PainelControlo.Enabled = True
    PainelControlo.SetFocus
    Unload Me
    
End Sub

Private Sub cmdProcurar_Click()
        
    InputBoxProcurar.Caption = "Procurar (Facturas)"
    InputBoxProcurar.Label1.Caption = "Indique o nome do proprietário da factura que pretende procurar."
    Facturas.Enabled = False
    InputBoxProcurar.Show
    
End Sub

Private Sub cmdRepor_Click()
    
    Dim itemx As ListItem
    
    SQL = " SELECT * " _
        & " FROM faturas" _
        & " ORDER BY cod_fatura"
        
    Set rst = db.OpenRecordset(SQL)
    
    ListFacturas.ListItems.Clear
 
    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = ListFacturas.ListItems.Add(, , rst("cod_fatura"))
        itemx.SubItems(1) = rst("cod_paciente")
        itemx.SubItems(2) = rst("nome_paciente")
        itemx.SubItems(3) = rst("data")
        itemx.SubItems(4) = rst("cod_medicamento")
        itemx.SubItems(5) = rst("nome_medicamento")
        itemx.SubItems(6) = rst("quantidade")
        itemx.SubItems(7) = rst("preco_unidade")
        itemx.Tag = rst("cod_fatura")
        rst.MoveNext
    Wend
    
    cmdProcurar.Visible = True
    cmdRepor.Visible = False
    
End Sub

Private Sub Form_Load()

    Set db = OpenDatabase(App.Path & "\Clinica.mdb")
    
    Skin1.LoadSkin App.Path & "\Skin\dogmas.skn"
    Skin1.ApplySkin Me.hWnd
    
    ListFacturas_SetUp
    
End Sub

Private Sub ListFacturas_SetUp()
    
    SQL = "select * from faturas order by cod_fatura"
    Set rst = db.OpenRecordset(SQL)
    
    ListFacturas.ListItems.Clear
    ListFacturas.ColumnHeaders.Add , , "Cod. Factura", ListFacturas.Width - 9700
    ListFacturas.ColumnHeaders.Add , , "Cod. Paciente", ListFacturas.Width - 9600
    ListFacturas.ColumnHeaders.Add , , "Nome Paciente", ListFacturas.Width - 8900
    ListFacturas.ColumnHeaders.Add , , "Data", ListFacturas.Width - 9800
    ListFacturas.ColumnHeaders.Add , , "Cod. Medicamento", ListFacturas.Width - 9300
    ListFacturas.ColumnHeaders.Add , , "Medicamento", ListFacturas.Width - 9200
    ListFacturas.ColumnHeaders.Add , , "Quantidade", ListFacturas.Width - 9800
    ListFacturas.ColumnHeaders.Add , , "Preço Unidade", ListFacturas.Width - 9500
    ListFacturas.View = lvwReport
    
    If rst.BOF = True And rst.EOF = True Then Exit Sub
    
    While Not rst.EOF
        Set itemx = ListFacturas.ListItems.Add(, , rst("cod_fatura"))
        itemx.SubItems(1) = rst("cod_paciente")
        itemx.SubItems(2) = rst("nome_paciente")
        itemx.SubItems(3) = rst("data")
        itemx.SubItems(4) = rst("cod_medicamento")
        itemx.SubItems(5) = rst("nome_medicamento")
        itemx.SubItems(6) = rst("quantidade")
        itemx.SubItems(7) = rst("preco_unidade")
        itemx.Tag = rst("cod_fatura")
        rst.MoveNext
    Wend
 
End Sub
