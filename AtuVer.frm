VERSION 5.00
Begin VB.Form formAtuVer 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atualização automática de Versão"
   ClientHeight    =   5790
   ClientLeft      =   2760
   ClientTop       =   1755
   ClientWidth     =   9420
   ControlBox      =   0   'False
   Icon            =   "AtuVer.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5790
   ScaleWidth      =   9420
   Begin VB.Frame Frame1 
      Height          =   45
      Left            =   240
      TabIndex        =   11
      Top             =   5220
      Width           =   8955
   End
   Begin VB.CommandButton fcmbF03Atu 
      Caption         =   "Atualizar (F3)"
      Height          =   315
      Left            =   8010
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5370
      Width           =   1120
   End
   Begin VB.Label flblMsg010 
      Height          =   225
      Left            =   3570
      TabIndex        =   10
      Top             =   3090
      Width           =   5505
   End
   Begin VB.Label flblMsg009 
      Height          =   225
      Left            =   3570
      TabIndex        =   9
      Top             =   2820
      Width           =   5505
   End
   Begin VB.Label flblMsg008 
      Height          =   225
      Left            =   3570
      TabIndex        =   8
      Top             =   2400
      Width           =   5505
   End
   Begin VB.Label flblMsg007 
      Height          =   225
      Left            =   3570
      TabIndex        =   7
      Top             =   2130
      Width           =   5505
   End
   Begin VB.Label flblMsg006 
      Height          =   225
      Left            =   3570
      TabIndex        =   6
      Top             =   1860
      Width           =   5505
   End
   Begin VB.Label flblMsg001 
      Height          =   225
      Left            =   3570
      TabIndex        =   5
      Top             =   360
      Width           =   5505
   End
   Begin VB.Label flblMsg002 
      Height          =   225
      Left            =   3570
      TabIndex        =   4
      Top             =   780
      Width           =   5505
   End
   Begin VB.Label flblMsg003 
      Height          =   225
      Left            =   3570
      TabIndex        =   3
      Top             =   1050
      Width           =   5505
   End
   Begin VB.Label flblMsg004 
      Height          =   225
      Left            =   3570
      TabIndex        =   2
      Top             =   1320
      Width           =   5505
   End
   Begin VB.Label flblMsg005 
      Height          =   225
      Left            =   3570
      TabIndex        =   1
      Top             =   1590
      Width           =   5505
   End
   Begin VB.Image Image1 
      BorderStyle     =   1  'Fixed Single
      Height          =   4815
      Left            =   300
      Picture         =   "AtuVer.frx":08CA
      Stretch         =   -1  'True
      Top             =   300
      Width           =   3120
   End
End
Attribute VB_Name = "formAtuVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 Option Explicit

Private fstrModulo As String, fstrArgmto(1 To 3) As String
Private Sub fcmbF03Atu_Click()
        rlsLimparMensagens

        fcmbF03Atu.Enabled = False

        flblMsg002.Caption = "Copiando "
        flblMsg002.Refresh
        flblMsg004.Caption = fstrArgmto(2)
        flblMsg004.Refresh
        flblMsg006.Caption = "para"
        flblMsg006.Refresh
        flblMsg008.Caption = fstrArgmto(3) & " aguarde..."
        flblMsg008.Refresh

        FileCopy fstrArgmto(2), fstrArgmto(3)

        MsgBox "Atualização concluída com sucesso!", vbInformation

        rlsLimparMensagens

        flblMsg002.Caption = "Reexecutando o Módulo " & fstrModulo & "..."
        flblMsg002.Refresh

        Shell App.Path & "\" & fstrArgmto(1) & ".exe", vbNormalFocus
        Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        Dim lintNumItm As Integer

        If (KeyCode < 112 Or KeyCode > 123) And KeyCode <> 27 Then Exit Sub

        For lintNumItm = 0 To Me.Controls.Count - 1
        If (InStr(1, Me.Controls(lintNumItm).Name, "fcmbF" & Format(KeyCode - 111, "00"), vbTextCompare)) Then
        If (Me.Controls(lintNumItm).Enabled) Then
            Me.Controls(lintNumItm).Value = True
        End If
            Exit For
        End If
        Next
End Sub
Private Sub Form_Load()
        Dim ldatDatLoc As Date

        Dim ldatDatRde As Date

        Me.Top = ((Screen.Height - Me.Height)) / 2
        Me.Left = (Screen.Width - (Me.Width)) / 2

        rlsObterArgumentos

        Select Case fstrArgmto(1)
               Case "Contabil"
                    fstrModulo = "Contabilidade"
               Case "Mantem"
                    fstrModulo = "Manutenção"
               Case Else
                    fstrModulo = fstrArgmto(1)
        End Select

        ldatDatLoc = FileDateTime(fstrArgmto(3))
        ldatDatRde = FileDateTime(fstrArgmto(2))

        flblMsg001.Caption = "Atualização do Módulo " & fstrModulo
        flblMsg001.Refresh
        flblMsg002.Caption = "Há uma Versão mais recente " & IIf(Mid(fstrArgmto(3), Len(fstrArgmto(3)) - 2, 3) = "chm", _
                             "da Ajuda ", "") & "deste Módulo em:"
        flblMsg002.Refresh
        flblMsg003.Caption = fstrArgmto(2)
        flblMsg003.Refresh
        flblMsg005.Caption = "Sua Versão de " & Format(ldatDatLoc, "dd/mm/yyyy") & " às " & _
                                                Format(ldatDatLoc, "hh:mm:ss") & " será substituída pela "
        flblMsg005.Refresh
        flblMsg006.Caption = "mais recente de " & Format(ldatDatRde, "dd/mm/yyyy") & " às " & _
                                                  Format(ldatDatRde, "hh:mm:ss")
        flblMsg006.Refresh
        flblMsg008.Caption = "Tecle F3 para Atualizar sua Versão agora."
        flblMsg008.Refresh
        flblMsg009.Caption = "Após a Atualização, uma outra Execução"
        flblMsg009.Refresh
        flblMsg010.Caption = "do Módulo " & fstrModulo & " será iniciada já com a nova Versão."
        flblMsg010.Refresh
End Sub
Private Sub rlsLimparMensagens()
        flblMsg002.Caption = ""
        flblMsg002.Refresh
        flblMsg003.Caption = ""
        flblMsg003.Refresh
        flblMsg004.Caption = ""
        flblMsg004.Refresh
        flblMsg005.Caption = ""
        flblMsg005.Refresh
        flblMsg006.Caption = ""
        flblMsg006.Refresh
        flblMsg008.Caption = ""
        flblMsg008.Refresh
        flblMsg009.Caption = ""
        flblMsg009.Refresh
        flblMsg010.Caption = ""
        flblMsg010.Refresh
End Sub
Private Sub rlsObterArgumentos()
        Dim lbooNaoLeu As Boolean

        Dim lintIndice As Integer

        Dim lintQtdArg As Integer

        Dim lstrChrCmd As String

        Dim lvarLinCmd As Variant

            lintQtdArg = 0
            lbooNaoLeu = True
            lvarLinCmd = Command()

        For lintIndice = 1 To Len(lvarLinCmd)
            lstrChrCmd = Mid(lvarLinCmd, lintIndice, 1)

        If (lstrChrCmd <> " ") Then
        If (lbooNaoLeu) Then

        If (lintQtdArg = 3) Then Exit Sub

            lintQtdArg = lintQtdArg + 1
            lbooNaoLeu = False
        End If
            fstrArgmto(lintQtdArg) = fstrArgmto(lintQtdArg) & lstrChrCmd
        Else
            lbooNaoLeu = True
        End If
        Next
End Sub
