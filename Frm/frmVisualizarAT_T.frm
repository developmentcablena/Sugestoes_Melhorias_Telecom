VERSION 5.00
Begin VB.Form frmVisualizarAT 
   Caption         =   "Sugestçoes de Melhoria - (Visualização)"
   ClientHeight    =   6450
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7545
   Icon            =   "frmVisualizarAT_T.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6450
   ScaleWidth      =   7545
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cboTipo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   360
      Width           =   2295
   End
   Begin VB.ComboBox cboElemento 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2520
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   6
      Top             =   360
      Width           =   4695
   End
   Begin VB.TextBox txt01 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      MaxLength       =   1000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Top             =   1920
      Width           =   7335
   End
   Begin VB.CommandButton cmdCadastrar 
      Appearance      =   0  'Flat
      Caption         =   "&Cadastrar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   6000
      Width           =   1335
   End
   Begin VB.CommandButton cmdFechar 
      Appearance      =   0  'Flat
      Caption         =   "Fechar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   3
      Top             =   6000
      Width           =   1335
   End
   Begin VB.ComboBox cboAtendente 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      ItemData        =   "frmVisualizarAT_T.frx":12CA
      Left            =   2520
      List            =   "frmVisualizarAT_T.frx":12CC
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1035
      Width           =   4695
   End
   Begin VB.ComboBox cboLocal 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   120
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox txt02 
      Appearance      =   0  'Flat
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   4080
      Width           =   7335
   End
   Begin VB.Label lblDivisao 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Elemento 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Elemento"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2520
      TabIndex        =   12
      Top             =   120
      Width           =   915
   End
   Begin VB.Label lblDescricaoServico 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reporte Tecnico:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   11
      Top             =   1680
      Width           =   1635
   End
   Begin VB.Label lblAtendente 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Atendente"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   2520
      TabIndex        =   10
      Top             =   795
      Width           =   1005
   End
   Begin VB.Label Local 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Local"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   840
      Width           =   510
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Reporte Tecnico - Concluido:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   120
      TabIndex        =   8
      Top             =   3840
      Width           =   2790
   End
End
Attribute VB_Name = "frmVisualizarAT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private strRelatorio As String

Private Sub cmdFechar_Click()
Call Unload(Me)
End Sub

Private Sub Form_Load()
Call suListarTipo
Call suListarElemento
Call suListarLocal
Call suListarAtendimento

Me.cboTipo.Enabled = False
Me.cboElemento.Enabled = False
Me.cboLocal.Enabled = False
Me.cboAtendente.Enabled = False
Me.txt01.Enabled = False
Me.txt02.Enabled = False

Me.cboTipo.Text = ValorTipo
Me.cboLocal.Text = ValorLocal

Me.cboAtendente.Text = ValorAtendente

Me.txt01 = ValorResultado
Me.txt02 = ValorFinalizado

If ValorElemento <> "" Then
    Me.cboElemento.Text = ValorElemento
Else
End If
End Sub

Private Sub suListarTipo()
Dim rs As New ADODB.Recordset
Call ConectarBD
    strSQL = "SELECT * FROM tb_Tipo"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        Me.cboTipo.AddItem rs!Tipo
        rs.MoveNext
    Loop
    Set rs = Nothing
  
End Sub

Private Sub suListarElemento()
    strSQL = "SELECT * FROM tb_Elemento"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        Me.cboElemento.AddItem rs!Elemento
        rs.MoveNext
    Loop
    Set rs = Nothing
    
End Sub

Private Sub suListarLocal()
    strSQL = "SELECT * FROM tb_Local"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        Me.cboLocal.AddItem rs!Local
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub

Private Sub suListarAtendimento()
    strSQL = "SELECT * FROM tb_Atendente"
    Set rs = New ADODB.Recordset
    rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly
    
    Do While Not rs.EOF
        Me.cboAtendente.AddItem rs!Atendente
        rs.MoveNext
    Loop
    Set rs = Nothing
End Sub


