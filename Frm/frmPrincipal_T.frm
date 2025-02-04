VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmPrincipal 
   Appearance      =   0  'Flat
   Caption         =   "Sugestões de Melhorias - Telecom"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11010
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPrincipal_T.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11010
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar stbPrincipal 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   8220
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3969
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "03/02/2025"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            Object.Width           =   2117
            MinWidth        =   2117
            TextSave        =   "17:46"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   3528
            MinWidth        =   3528
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlMenu 
      Left            =   10320
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal_T.frx":12CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal_T.frx":15E4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal_T.frx":32BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal_T.frx":3B98
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal_T.frx":47EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal_T.frx":365F0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal_T.frx":376CA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal_T.frx":387A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal_T.frx":38D3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal_T.frx":39E18
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPrincipal_T.frx":3AEF2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbMenu 
      Align           =   1  'Align Top
      Height          =   810
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11010
      _ExtentX        =   19420
      _ExtentY        =   1429
      ButtonWidth     =   3493
      ButtonHeight    =   1429
      Style           =   1
      ImageList       =   "imlMenu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sugestões de Melhorias"
            Key             =   "Sistemas"
            Object.ToolTipText     =   "Sugestões de Melhorias"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Acompanhar Sugestões"
            Key             =   "Acompanhar"
            Object.ToolTipText     =   "Acompanhar Sugestões"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Verificar sugestões"
            Key             =   "Chamados"
            Object.ToolTipText     =   "Verificar sugestões"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gerar Relatório"
            Key             =   "relatorio"
            Object.ToolTipText     =   "Gerar Relatório"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Usuários"
            Key             =   "Usuarios"
            Object.ToolTipText     =   "Usuários"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Usuários"
            Key             =   "u"
            Object.ToolTipText     =   "Usuários"
            Style           =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Sair"
            Key             =   "Sair"
            Object.ToolTipText     =   "Sair"
            ImageIndex      =   10
         EndProperty
      EndProperty
   End
   Begin VB.Image imgCablena 
      Appearance      =   0  'Flat
      Height          =   1560
      Left            =   8880
      Picture         =   "frmPrincipal_T.frx":3D4CC
      Top             =   6360
      Width           =   1875
   End
   Begin VB.Shape shpPrincipal 
      BackStyle       =   1  'Opaque
      Height          =   8175
      Left            =   0
      Top             =   360
      Width           =   11175
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private rs As ADODB.Recordset
Private strSQL As String
Private intTotalOS As Integer
Private intTotalOSPendente As Integer
Private intTotalOSAtendimento As Integer
Private intTotalOSNaoValidada As Integer

Private Sub Form_Load()
    With Me
        .tlbMenu.Buttons(1).Visible = False
        .tlbMenu.Buttons(2).Visible = False
        .tlbMenu.Buttons(3).Visible = False
        .tlbMenu.Buttons(4).Visible = False
        .tlbMenu.Buttons(5).Visible = False
    
    End With
    
Call ConectarBD
Call suVerificarPermissaoAD(PermissaoAD)
Call suVerificarPermissoes(gintUsuarioID)
End Sub

Private Sub Form_Resize()

    shpPrincipal.Width = Me.Width
    shpPrincipal.Height = Me.Height
    If Len(gstrNome) > 0 Then
        stbPrincipal.Panels(1).Text = "Usuário: " & gstrNome
    Else
        stbPrincipal.Panels(1).Text = "Usuário: " & sFullName
    End If
    
    If Len(gstrDepto) > 0 Then
        stbPrincipal.Panels(2).Text = "Departamento: " & gstrDepto
    Else
        stbPrincipal.Panels(2).Text = "Departamento: " & strDescription
    End If
    stbPrincipal.Panels(5).Text = "Versão: " & App.Major & "." & App.Minor & "." & App.Revision
    imgCablena.Left = Me.Width - 2400
    imgCablena.Top = Me.Height - 2700
End Sub

Private Sub Form_Unload(cancel As Integer)
  
    Set cn = Nothing
End Sub
Private Sub tlbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case Is = "Sistemas"
            Call frmSuporteSistemas.Show(vbModal)
        Case Is = "Chamados"
            Call frmChamados.Show(vbModal)
        Case Is = "Usuarios"
            Call frmUsuarios.Show(vbModal)
        Case Is = "Acompanhar"
            Call frmAcompanharOS.Show(vbModal)
        Case Is = "relatorio"
            Call frmRelatorio.Show(vbModal)
       
        Case Is = "Sair"
            Call Unload(Me)
    End Select
End Sub

Private Sub suVerificarPermissaoAD(ByVal vPermissaoAD As Boolean)
If vPermissaoAD = True Then
    tlbMenu.Buttons(1).Visible = True
    tlbMenu.Buttons(2).Visible = True
    If vUsuario = "guilherme.souza" Then
        tlbMenu.Buttons(1).Visible = True
        tlbMenu.Buttons(2).Visible = True
        tlbMenu.Buttons(3).Visible = True
        tlbMenu.Buttons(4).Visible = True
        tlbMenu.Buttons(5).Visible = True
        Me.Caption = "Sugestões de Melhorias - Telecom - (administrador)"
        
    ElseIf vUsuario = "kalves" Then
        tlbMenu.Buttons(1).Visible = True
        tlbMenu.Buttons(2).Visible = True
        tlbMenu.Buttons(3).Visible = True
        tlbMenu.Buttons(4).Visible = True
        tlbMenu.Buttons(5).Visible = True
        Me.Caption = "Sugestões de Melhorias - Telecom - (administrador)"
        
    ElseIf vUsuario = "emperes" Then
        tlbMenu.Buttons(1).Visible = True
        tlbMenu.Buttons(2).Visible = True
        tlbMenu.Buttons(3).Visible = True
        tlbMenu.Buttons(4).Visible = True
        tlbMenu.Buttons(5).Visible = True
        Me.Caption = "Sugestões de Melhorias - Telecom - (administrador)"
    End If
Else
    tlbMenu.Buttons(1).Visible = False
    tlbMenu.Buttons(2).Visible = False
    
End If

End Sub


Private Sub suVerificarPermissoes(ByVal vUsuarioID As Integer)
strSQL = "SELECT * FROM vw_Usuarios WHERE UsuarioID = " & vUsuarioID & ""

Set rs = New ADODB.Recordset
rs.Open strSQL, cn, adOpenForwardOnly, adLockReadOnly

Do While Not rs.EOF
    Select Case rs!Permissao
        Case Is = "Administrador"
            If rs!Permissao = "Administrador" Then
                tlbMenu.Buttons(1).Visible = True
                tlbMenu.Buttons(2).Visible = True
                tlbMenu.Buttons(3).Visible = True
                tlbMenu.Buttons(4).Visible = True
                tlbMenu.Buttons(5).Visible = False
            Else
                tlbMenu.Buttons(1).Visible = False
                tlbMenu.Buttons(2).Visible = False
                tlbMenu.Buttons(3).Visible = False
                tlbMenu.Buttons(4).Visible = False
                tlbMenu.Buttons(5).Visible = False
            End If
               
        Case Is = "Usuário"
            If rs!Permissao = "Usuário" Then
                tlbMenu.Buttons(1).Visible = True
                tlbMenu.Buttons(2).Visible = True
            Else
                tlbMenu.Buttons(1).Visible = False
                tlbMenu.Buttons(2).Visible = False
            End If
    End Select
    
    rs.MoveNext
Loop
   
End Sub










