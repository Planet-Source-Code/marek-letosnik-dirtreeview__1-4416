VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "DirTreeView"
   ClientHeight    =   5730
   ClientLeft      =   1950
   ClientTop       =   1590
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   6315
   Begin VB.CommandButton Command1 
      Caption         =   "Go"
      Height          =   255
      Left            =   5880
      TabIndex        =   3
      Top             =   5420
      Width           =   375
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   20
      TabIndex        =   2
      Text            =   "c:\dokumenty"
      Top             =   5400
      Width           =   5760
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   2520
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   3000
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.ImageList img 
      Left            =   1320
      Top             =   3000
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   128
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form.frx":0000
            Key             =   "unknown"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form.frx":0944
            Key             =   "fixed"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form.frx":30F8
            Key             =   "ram"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form.frx":58AC
            Key             =   "remove"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form.frx":8060
            Key             =   "cd"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form.frx":A814
            Key             =   "folder"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form.frx":CFC8
            Key             =   "open"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form.frx":F77C
            Key             =   "remote"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView DirTree 
      Height          =   5295
      Left            =   20
      TabIndex        =   0
      Top             =   30
      Width           =   6285
      _ExtentX        =   11086
      _ExtentY        =   9340
      _Version        =   393217
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "img"
      Appearance      =   1
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Author: Marek Letosnik
' letosnik@atlas.cz

Private nNode As Node
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long

Private Sub Command1_Click()
Dim Strom As String, Where As String, h As Integer
  Where = Text1
  h = 1
Znova:
  Do
    Strom = LCase(Mid(DirTree.Nodes(h).FullPath, InStr(1, DirTree.Nodes(h).FullPath, ":") - 1, 2) & Mid(DirTree.Nodes(h).FullPath, InStr(1, DirTree.Nodes(h).FullPath, ":") + 2))
    If Left(LCase(Where), Len(Strom)) = LCase(Strom) Then
        DirTree.Nodes(h).Expanded = True
        If DirTree.Nodes(h).Children > 0 Then h = DirTree.Nodes(h).Child.Index Else Exit Do
        GoTo Znova
        Exit Do
    End If
    If h = DirTree.Nodes(h).LastSibling.Index Then Exit Do
    h = DirTree.Nodes(h).Next.Index
  Loop
End Sub

Private Sub DirTree_Expand(ByVal Node As MSComctlLib.Node)
Dim j As Integer
    For j = Node.Child.FirstSibling.Index To Node.Child.LastSibling.Index
        DirTree_NodeClick DirTree.Nodes(j)
    Next j
    DirTree_NodeClick Node
    Node.Selected = True
End Sub

Private Sub DirTree_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then DirTree.Nodes.Clear: LoadTreeView
End Sub

Private Sub Form_Load()
    LoadTreeView
End Sub

Private Sub DirTree_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim Path As String
    If Left(Node.Key, 4) = "root" Then
        On Error Resume Next
        If Node.Children > 0 Then GoTo Skok
        DisplayDir Mid(Node.Text, Len(Node.Text) - 2, 2), Node.Key
    End If
    Path = Mid(Node.FullPath, InStr(1, Node.FullPath, ":") - 1, 2) & Mid(Node.FullPath, InStr(1, Node.FullPath, ":") + 2)
    If Node.Children > 0 Then GoTo Skok
    DisplayDir Path, Node.Index

Skok:
    Path = Mid(Node.FullPath, InStr(1, Node.FullPath, ":") - 1, 2) & Mid(Node.FullPath, InStr(1, Node.FullPath, ":") + 2)
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    Text1 = Path
End Sub

Sub DisplayDir(Pth, Parent)
Dim j As Integer
    On Error Resume Next
    Pth = Pth & "\"
    tmp = Dir(Pth, vbDirectory)
    Do Until tmp = ""
        If tmp <> "." And tmp <> ".." Then
            If GetAttr(Pth & tmp) And vbDirectory Then
                'I use ListBox with property Sorted=True to
                'alphabetize directories. Easy eh? ;-)
                List1.AddItem StrConv(tmp, vbProperCase)
                'StrConv function convert for example
                '"WINDOWS" to "Windows"
            End If
        End If
        tmp = Dir
    Loop
    'Add sorted directory names to TreeView
    For j = 1 To List1.ListCount
        Set nNode = DirTree.Nodes.Add(Parent, tvwChild, , List1.List(j - 1), "folder")
        nNode.ExpandedImage = "open"
    Next j
    List1.Clear
End Sub

Private Sub LoadTreeView()
    Dim DriveNum As String
    Dim DriveType As Long
    DriveNum = 64
    On Error Resume Next
    Do
        DriveNum = DriveNum + 1
        DriveType = GetDriveType(Chr$(DriveNum) & ":\")
        If DriveNum > 90 Then Exit Do
        Select Case DriveType
            Case 0: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "unknown")
                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
            Case 2: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, "(" & Chr$(DriveNum) & ":)", "remove")
            Case 3: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "fixed")
                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
            Case 4: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "remote")
                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
            Case 5: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "cd")
                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
            Case 6: Set nNode = DirTree.Nodes.Add(, , "root" & DriveNum, StrConv(Dir(Chr$(DriveNum) & ":", vbVolume), vbProperCase) & " (" & Chr$(DriveNum) & ":)", "ram")
                    DisplayDir Mid(DirTree.Nodes("root" & DriveNum).Text, Len(DirTree.Nodes("root" & DriveNum).Text) - 2, 2), "root" & DriveNum
        End Select
    Loop
End Sub
