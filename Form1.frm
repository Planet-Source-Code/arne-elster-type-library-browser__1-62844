VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Type Lib Quick View"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8190
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   284
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   546
   StartUpPosition =   2  'Bildschirmmitte
   Begin MSComDlg.CommonDialog dlg 
      Left            =   2400
      Top             =   75
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3840
      Left            =   3000
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Beides
      TabIndex        =   3
      Top             =   300
      Width           =   5040
   End
   Begin VB.ListBox lstInfo 
      Height          =   3855
      IntegralHeight  =   0   'False
      Left            =   150
      TabIndex        =   0
      Top             =   300
      Width           =   2640
   End
   Begin VB.Label Label2 
      Caption         =   "Content:"
      Height          =   240
      Left            =   3000
      TabIndex        =   2
      Top             =   75
      Width           =   2265
   End
   Begin VB.Label Label1 
      Caption         =   "Type Infos:"
      Height          =   240
      Left            =   150
      TabIndex        =   1
      Top             =   75
      Width           =   1665
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open..."
      End
      Begin VB.Menu mnuS1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private typeinfo    As clsTypeLibInfo

Private Sub Form_Load()
    Set typeinfo = New clsTypeLibInfo
End Sub

Private Sub Form_Resize()
    On Error Resume Next

    Dim lwidth  As Long
    Dim lheight As Long

    lwidth = Me.ScaleWidth - lstInfo.Left * 2 - 6

    lstInfo.Width = lwidth * (2 / 6)
    txt.Width = lwidth * (4 / 6)

    txt.Left = lstInfo.Left + lstInfo.Width + 6
    Label2.Left = txt.Left

    lheight = Me.ScaleHeight - lstInfo.Top * 2 + Label1.Height
    lstInfo.Height = lheight
    txt.Height = lheight
End Sub

Private Sub lstInfo_Click()
    ShowTypeInfo lstInfo.ListIndex
End Sub

Private Sub mnuExit_Click()
    typeinfo.CloseTypeLib
    Unload Me
End Sub

Private Sub mnuOpen_Click()
    On Error Resume Next

    dlg.Filter = "Supported (*.dll;*.exe;*.ocx)|*.dll;*.exe;*.ocx|DLL (*.dll)|*.dll|EXE (*.exe)|*.exe|OCX (*.ocx)|*.ocx|All files (*.*)|*.*"
    dlg.ShowOpen
    If Err Then Exit Sub

    ShowTypes dlg.Filename
End Sub

Private Sub ShowTypes(strFile As String)
    Dim i   As Long

    ClearFields

    If Not typeinfo.OpenTypeLib(strFile) Then
        MsgBox "Couldn't open typelib.", vbExclamation
        Exit Sub
    End If

    Me.Caption = "Type Lib Quick View - " & typeinfo.TypeLibName

    lstInfo.Visible = False
    With typeinfo
        For i = 0 To .TypeInfoCount - 1
            .SelectTypeInfo i
            lstInfo.AddItem Left(.TKind2String(.TypeInfoKind), 1) & " " & .TypeInfoName
        Next
    End With
    lstInfo.Visible = True
End Sub

Private Sub ShowTypeInfo(index As Long)
    On Error Resume Next

    Dim i   As Long
    Dim j   As Long

    txt.Visible = False
    txt.Text = ""

    If Not typeinfo.SelectTypeInfo(index) Then
        MsgBox "Couldn't select type info.", vbExclamation
        Exit Sub
    End If

    With typeinfo
        TApp "Name: " & .TypeInfoName
        TApp "GUID: " & .TypeInfoGUID
        TApp "Prog ID: " & .TypeInfoPrgID
        TApp "Kind: " & .TKind2String(.TypeInfoKind)
        TApp
        TApp "typedef: " & .AliasName
        TApp
        TApp "Functions: " & .TypeInfoFunctions

        For i = 0 To .TypeInfoFunctions - 1
            If .SelectFunction(i) Then
                TApp "(VTable " & Format(.FunctionVTOffset, "00") & ") ", False
                TApp .FunctionReturnType & " ", False
                TApp .FunctionName & " (", False
                For j = 0 To .ParameterCount - 1
                    .SelectParameter j
                    TApp "[" & .ParamFlags2String(.ParameterFlags), False
'                    If .ParameterFlags And PARAMFLAG_FHASDEFAULT Then
'                        TApp ", default=", False
'                        TApp .ParameterDefault, False
'                    End If
                    TApp "] ", False
                    TApp .ParameterName & " As ", False
                    TApp .ParameterType & IIf(j < .ParameterCount - 1, ", ", ""), False
                Next
                TApp ") ", False
                TApp .InvKind2String(.FunctionInvKind, True)
            Else
                TApp ">> no info about function " & i
            End If
        Next

        TApp
        TApp "Variables: " & .VariableCount

        For i = 0 To .VariableCount - 1
            If .SelectVariable(i) Then
                TApp .VarKind2String(.VariableKind) & " ", False
                TApp .VariableName, False
                TApp " As " & .VariableType, False
                TApp " = " & .VariableValue
            Else
                TApp ">> no info about variable " & i
            End If
        Next

        TApp
        TApp "Implements: " & .TypeInfoImplements
        For i = 0 To .TypeInfoImplements - 1
            If .SelectImplement(i) Then
                TApp .ImplementName & " " & .ImplementGUID
            End If
        Next

    End With

    txt.Visible = True
End Sub

Private Sub ClearFields()
    lstInfo.Clear
    txt.Text = ""
End Sub

Private Sub TApp(Optional strText As String, Optional break As Boolean = True)
    txt.Text = txt.Text & strText & IIf(break, vbCrLf, "")
End Sub
