VERSION 5.00
Begin VB.Form ADDItem 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   2520
      Width           =   1455
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3360
      TabIndex        =   7
      Text            =   "1"
      Top             =   1800
      Width           =   615
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   1455
      Begin VB.OptionButton Option2 
         Caption         =   "Movie"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Grocery"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1215
      End
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   360
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Original QTY"
      Height          =   375
      Left            =   1920
      TabIndex        =   8
      Top             =   1800
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Name"
      Height          =   375
      Left            =   360
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Item Number"
      Height          =   375
      Left            =   360
      TabIndex        =   1
      Top             =   360
      Width           =   975
   End
End
Attribute VB_Name = "ADDItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsetcustomers As ADODB.Recordset
Dim rsetheader As ADODB.Recordset
Dim conAVB As ADODB.Connection
Dim header As ADODB.Command



Private Sub Command1_Click()


Dim strsql As String
Dim strch As String

If Option1.value = True Then
strch = "Grocery"
End If
If Option2.value = True Then
strch = "Movie"
End If




Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & _
        "persist security info=False;Data Source=" & App.Path & _
        "\inventory.mdb;Mode = readwrite"
    conAVB.Open
   Set header = New ADODB.Command
    Set header.ActiveConnection = conAVB

 Set rsetheader = New ADODB.Recordset
    rsetheader.Open strch, conAVB, adOpenDynamic, adLockOptimistic, adCmdTable
    
    'Set rsetheader = header.Execute

 
    header.CommandText = "insert into " & strch & "(item_num,Name,QTY,Original_Date) values ('" & Text1.Text & "','" & Text2.Text & "'," & Text3.Text & ",#" & Date & "#)"
    header.Execute
   
    
conAVB.Close

Unload Me

End Sub


Private Sub Form_Load()
If strmov = False Then
Option1.value = True
End If
If strmov = True Then
Option2.value = True
End If





Text1.Text = stritem

'Text2.SetFocus

End Sub


Private Sub Text1_GotFocus()
Text2.SetFocus
End Sub


