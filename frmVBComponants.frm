VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmVBComponants 
   Caption         =   "VB Componants"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10695
   Icon            =   "frmVBComponants.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6390
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox txtLines 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   7140
      TabIndex        =   6
      Text            =   "0"
      Top             =   60
      Width           =   1035
   End
   Begin vbComponants.MSplitter vSplitter 
      Height          =   2115
      Left            =   3300
      TabIndex        =   5
      Top             =   720
      Width           =   75
      _extentx        =   132
      _extenty        =   3731
      controlname1    =   "tvProject"
      controlname2    =   "lvProject"
      splitterwidth   =   75
      controlvisible1 =   -1  'True
      controlvisible2 =   -1  'True
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6600
      Top             =   3240
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBComponants.frx":08CA
            Key             =   "Reference"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBComponants.frx":11A4
            Key             =   "Designer"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBComponants.frx":14BE
            Key             =   "Form"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBComponants.frx":17D8
            Key             =   "MDIForm"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBComponants.frx":1AF2
            Key             =   "Module"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBComponants.frx":23CC
            Key             =   "ProjectGroup"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBComponants.frx":2CA6
            Key             =   "Project"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBComponants.frx":3580
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmVBComponants.frx":3E5A
            Key             =   "UC"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvProject 
      Height          =   2115
      Left            =   0
      TabIndex        =   4
      Top             =   720
      Width           =   3300
      _ExtentX        =   5821
      _ExtentY        =   3731
      _Version        =   393217
      LabelEdit       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin MSComctlLib.ListView lvProject 
      Height          =   2115
      Left            =   3375
      TabIndex        =   3
      Top             =   720
      Width           =   7320
      _ExtentX        =   12912
      _ExtentY        =   3731
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   0   'False
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   4
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "Name"
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Path"
         Text            =   "Path"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "Type"
         Text            =   "Type"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Key             =   "Lines"
         Text            =   "Lines"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   315
      Left            =   5640
      TabIndex        =   2
      Top             =   60
      Width           =   735
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8940
      Top             =   60
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtProject 
      Height          =   285
      Left            =   840
      TabIndex        =   0
      Top             =   60
      Width           =   4155
   End
   Begin VB.Label Label2 
      Caption         =   "Lines"
      Height          =   195
      Left            =   6600
      TabIndex        =   7
      Top             =   120
      Width           =   495
   End
   Begin VB.Image imgPath 
      Height          =   240
      Left            =   5160
      Picture         =   "frmVBComponants.frx":4734
      Top             =   120
      Width           =   240
   End
   Begin VB.Label Label1 
      Caption         =   "Project"
      Height          =   195
      Left            =   180
      TabIndex        =   1
      Top             =   120
      Width           =   675
   End
End
Attribute VB_Name = "frmVBComponants"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()

On Error GoTo errhandler

Dim iFile           As Integer
Dim sFile           As String
Dim sstr            As String
Dim sType           As String
Dim iAt             As Integer
Dim iAt2            As Integer
Dim listx           As ListItem
Dim Nodex           As Node
Dim sTemp           As String
Dim sName           As String
Dim sPath           As String
Dim aSplit          As Variant
Dim lTotLines       As String
Dim lLines          As Long

    If Dir(txtProject.Text) = "" Then
        Exit Sub
    End If
    
    lvProject.ListItems.Clear
    While tvProject.Nodes("Forms").Children > 0
        tvProject.Nodes.Remove tvProject.Nodes("Forms").Child.Index
    Wend

    While tvProject.Nodes("Modules").Children > 0
        tvProject.Nodes.Remove tvProject.Nodes("Modules").Child.Index
    Wend

    While tvProject.Nodes("Classes").Children > 0
        tvProject.Nodes.Remove tvProject.Nodes("Classes").Child.Index
    Wend
    
    While tvProject.Nodes("References").Children > 0
        tvProject.Nodes.Remove tvProject.Nodes("References").Child.Index
    Wend
    
    While tvProject.Nodes("Designers").Children > 0
        tvProject.Nodes.Remove tvProject.Nodes("Designers").Child.Index
    Wend
    
    lTotLines = 0
    
    lvProject.ListItems.Clear

    iFile = FreeFile
    sFile = Trim(txtProject.Text)
    Open sFile For Input As #iFile
    
    While Not EOF(iFile)
        Line Input #iFile, sstr
        iAt = InStr(1, sstr, "=")
        If iAt > 0 Then
            sType = Left(sstr, iAt - 1)
            Select Case sType
                Case "Form"
                    sTemp = Trim(sstr)
                    sTemp = Mid(sTemp, iAt + 1, Len(sTemp) - iAt)
                    Set Nodex = tvProject.Nodes.Add("Forms", tvwChild, "F" & Filename(sTemp), Filename(sTemp), "Form")
                    Nodex.Tag = sTemp
                    
                    lLines = FileLines(Trim(FilePath(txtProject.Text)) & Trim(sTemp))
                    
                    Set listx = lvProject.ListItems.Add(, "F" & Filename(sTemp), Filename(sTemp), "Form", "Form")
                    listx.ListSubItems.Add , "Path", sTemp
                    listx.ListSubItems.Add , "Type", "Form"
                    listx.ListSubItems.Add , "Lines", Format(lLines, "0000")
                    
                    
                    lTotLines = lTotLines + lLines
                    
                
                Case "Module"
                    sTemp = Trim(sstr)
                    sTemp = Mid(sTemp, iAt + 1, Len(sTemp) - iAt)
                    iAt2 = InStr(1, sTemp, ";")
                    sName = Trim(Left(sTemp, iAt2 - 1))
                    sPath = Trim(Mid(sTemp, iAt2 + 1, Len(sTemp) - iAt2))
                    Set Nodex = tvProject.Nodes.Add("Modules", tvwChild, "M" & sName, sName, "Module")
                    Nodex.Tag = sPath
                    
                    lLines = FileLines(FilePath(Trim(txtProject.Text)) & sPath)
                    
                    Set listx = lvProject.ListItems.Add(, "M" & Filename(sName), Filename(sName), "Module", "Module")
                    listx.ListSubItems.Add , "Path", sPath
                    listx.ListSubItems.Add , "Type", "Module"
                    listx.ListSubItems.Add , "Lines", Format(lLines, "0000")
                    
                    lTotLines = lTotLines + lLines
                    
                Case "Class"
                    sTemp = Trim(sstr)
                    sTemp = Mid(sTemp, iAt + 1, Len(sTemp) - iAt)
                    iAt2 = InStr(1, sTemp, ";")
                    sName = Trim(Left(sTemp, iAt2 - 1))
                    sPath = Mid(sTemp, iAt2 + 1, Len(sTemp) - iAt2)
                    Set Nodex = tvProject.Nodes.Add("Classes", tvwChild, "C" & sName, sName, "Class")
                    Nodex.Tag = sPath
            
            
                    lLines = FileLines(FilePath(Trim(txtProject.Text)) & Trim(sPath))
                
                    Set listx = lvProject.ListItems.Add(, "C" & Filename(sName), Filename(sName), "Class", "Class")
                    listx.ListSubItems.Add , "Path", sPath
                    listx.ListSubItems.Add , "Type", "Class"
                    listx.ListSubItems.Add , "Lines", Format(lLines, "0000")
                    
                    lTotLines = lTotLines + lLines
                
                Case "Designer"
                    sTemp = Trim(sstr)
                    sTemp = Mid(sTemp, iAt + 1, Len(sTemp) - iAt)
                    Set Nodex = tvProject.Nodes.Add("Designers", tvwChild, "D" & sTemp, sTemp, "Designer")
                    Nodex.Tag = sTemp
                
                Case "Reference"
                    sTemp = Trim(sstr)
                    sTemp = Mid(sTemp, iAt + 1, Len(sTemp) - iAt)
                    aSplit = Split(sTemp, "#")
                    Set Nodex = tvProject.Nodes.Add("References", tvwChild, "R" & Trim(aSplit(4)), Trim(aSplit(4)), "Reference")
                    Nodex.Tag = Trim(aSplit(3))
            
                Case "UserControl"
                    sTemp = Trim(sstr)
                    sTemp = Mid(sTemp, iAt + 1, Len(sTemp) - iAt)
                    sName = Filename(Trim(sTemp))
                    sPath = Trim(sTemp)
                    Set Nodex = tvProject.Nodes.Add("UserControls", tvwChild, "UC" & sName, sName, "UC")
                    Nodex.Tag = sPath
            
            
                    lLines = FileLines(FilePath(Trim(txtProject.Text)) & Trim(sPath))
                
                    Set listx = lvProject.ListItems.Add(, "UC" & sName, sName, "UC", "UC")
                    listx.ListSubItems.Add , "Path", sPath
                    listx.ListSubItems.Add , "Type", "User Control"
                    listx.ListSubItems.Add , "Lines", Format(lLines, "0000")
                    
                    lTotLines = lTotLines + lLines
            
            End Select
        End If
    Wend
                    
    Close #iFile

' ***********************************************
' *** Second Pass to get Paths
' ***********************************************
    iFile = FreeFile
    sFile = Trim(txtProject.Text)
    Open sFile For Input As #iFile
    
On Error Resume Next

    While Not EOF(iFile)
        Line Input #iFile, sstr
        iAt = InStr(1, sstr, "=")
        If iAt > 0 Then
            sType = Left(sstr, iAt - 1)
            Select Case sType
                Case "Form"
                    sTemp = Trim(sstr)
                    sTemp = Mid(sTemp, iAt + 1, Len(sTemp) - iAt)
                    If Filename(Trim(sTemp)) <> Trim(sTemp) Then
                        sName = Filename(sTemp)
                        tvProject.Nodes("F" & sName).Tag = sTemp
                    End If
                
'                Case "Module"
'                    sTemp = Trim(sstr)
'                    sTemp = Mid(sTemp, iAt + 1, Len(sTemp) - iAt)
'                    If Filename(Trim(sTemp)) = Trim(sTemp) Then
'                        Set Nodex = tvProject.Nodes.Add("Modules", tvwChild, "M" & sTemp, sTemp, "Module")
'                    End If
'
'                Case "Class"
'                    sTemp = Trim(sstr)
'                    sTemp = Mid(sTemp, iAt + 1, Len(sTemp) - iAt)
'                    iAt2 = InStr(1, sTemp, ";")
'                    sName = Trim(Left(sTemp, iAt2 - 1))
'                    sPath = Mid(sTemp, iAt2 + 1, Len(sTemp) - iAt2)
'                    If Filename(sPath) = sPath Then
'                        Set Nodex = tvProject.Nodes.Add("Classes", tvwChild, "C" & sName, sName, "Class")
'                    End If
'
'                Case "Designer"
'                    sTemp = Trim(sstr)
'                    sTemp = Mid(sTemp, iAt + 1, Len(sTemp) - iAt)
'                    Set Nodex = tvProject.Nodes.Add("Designers", tvwChild, "D" & sTemp, sTemp, "Designer")
'
'                Case "Reference"
'                    sTemp = Trim(sstr)
'                    sTemp = Mid(sTemp, iAt + 1, Len(sTemp) - iAt)
'                    aSplit = Split(sTemp, "#")
'                    Set Nodex = tvProject.Nodes.Add("References", tvwChild, "R" & Trim(aSplit(4)), Trim(aSplit(4)), "Reference")
            
            End Select
        End If
    Wend
                    
    Close #iFile
    
    tvProject.Nodes("Forms").Expanded = False
    tvProject.Nodes("Modules").Expanded = False
    tvProject.Nodes("Classes").Expanded = False
    tvProject.Nodes("Designers").Expanded = False
    tvProject.Nodes("References").Expanded = False

    tvProject.Nodes("Forms").Sorted = True
    tvProject.Nodes("Modules").Sorted = True
    tvProject.Nodes("Classes").Sorted = True
    tvProject.Nodes("Designers").Sorted = True
    tvProject.Nodes("References").Sorted = True
    
    txtLines.Text = Format(lTotLines, "####0")

    Exit Sub
    
errhandler:
    ShowError
    On Error Resume Next
    Close #iFile

    
End Sub

Private Sub Form_Load()
    tvProject.Nodes.Add , , "Project", "Project", "Project"
    tvProject.Nodes.Add "Project", tvwChild, "Forms", "Forms", "Form"
    tvProject.Nodes.Add "Project", tvwChild, "Modules", "Modules", "Module"
    tvProject.Nodes.Add "Project", tvwChild, "Classes", "Classes", "Class"
    tvProject.Nodes.Add "Project", tvwChild, "References", "References", "Reference"
    tvProject.Nodes.Add "Project", tvwChild, "Designers", "Designers", "Designer"
    tvProject.Nodes.Add "Project", tvwChild, "UserControls", "User Controls", "UC"
    
End Sub

Private Sub Form_Resize()
On Error Resume Next
    vSplitter.Height = Me.ScaleHeight - vSplitter.Top
    vSplitter.Resize
    
End Sub

Private Sub imgPath_Click()

    txtProject.Text = GetFileName(CommonDialog1, "Open Project", "VB Projects|*.vbp", "VB Projects|*.vbp", True)
    
End Sub

Private Sub lvProject_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
   lvProject.SortKey = ColumnHeader.Index - 1
   lvProject.Sorted = True
    
End Sub

Private Sub lvProject_DblClick()
Dim slink       As String
    If lvProject.ListItems.Count < 1 Then Exit Sub
    
    slink = FilePath(Trim(txtProject.Text)) & Trim(lvProject.SelectedItem.ListSubItems("Path"))
    
    ShellExecute hWnd, "Open", slink, "", App.Path, 1
    
End Sub

Private Sub tvProject_NodeClick(ByVal Node As MSComctlLib.Node)
On Error Resume Next
    
    If Node.Key <> "Project" Then
        Select Case tvProject.Nodes(Node.Key).Parent.Key
            Case "Forms", "Modules", "Classes", "UserControls"
                lvProject.ListItems(Node.Key).Selected = True
                lvProject.ListItems(Node.Key).EnsureVisible
        End Select
    End If
End Sub
