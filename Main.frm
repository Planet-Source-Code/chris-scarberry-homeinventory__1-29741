VERSION 5.00
Begin VB.Form Main 
   Caption         =   "Inventory V1.0"
   ClientHeight    =   3630
   ClientLeft      =   165
   ClientTop       =   735
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3630
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Option8 
      Caption         =   "Video"
      Height          =   255
      Left            =   0
      TabIndex        =   17
      Top             =   3360
      Width           =   1095
   End
   Begin VB.OptionButton Option7 
      Caption         =   "Grocery"
      Height          =   255
      Left            =   0
      TabIndex        =   16
      Top             =   3120
      Width           =   1095
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      TabIndex        =   13
      Top             =   2160
      Width           =   2535
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Done"
      Height          =   495
      Left            =   2880
      TabIndex        =   11
      Top             =   3120
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Height          =   495
      Left            =   1200
      TabIndex        =   10
      Top             =   3120
      Width           =   1335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3600
      TabIndex        =   7
      Text            =   "1"
      Top             =   1080
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   1575
      Left            =   120
      TabIndex        =   4
      Top             =   1440
      Width           =   1455
      Begin VB.OptionButton Option6 
         Caption         =   "Returned"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Borrow"
         Height          =   315
         Left            =   120
         TabIndex        =   12
         Top             =   720
         Width           =   1215
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Delete"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   480
         Width           =   1215
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Add"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   480
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   1455
      Begin VB.OptionButton Option2 
         Caption         =   "Keyboard"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Scan(cat)"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Name"
      Height          =   375
      Left            =   2040
      TabIndex        =   14
      Top             =   1800
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "QTY"
      Height          =   375
      Left            =   2760
      TabIndex        =   9
      Top             =   1080
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Item Number"
      Height          =   375
      Left            =   2040
      TabIndex        =   8
      Top             =   120
      Width           =   1215
   End
   Begin VB.Menu mnureports 
      Caption         =   "Reports"
      Begin VB.Menu mnugroc 
         Caption         =   "Grocery List"
      End
      Begin VB.Menu mnubmov 
         Caption         =   "Borrowed Movies"
      End
      Begin VB.Menu mnustock 
         Caption         =   "Groceries in Stock"
      End
      Begin VB.Menu mnumovoh 
         Caption         =   "Movies on Hand"
      End
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsetcustomers As ADODB.Recordset
Dim rsetheader As ADODB.Recordset
Dim conAVB As ADODB.Connection
Dim header As ADODB.Command
Dim addnew As Boolean



Private Sub Command1_Click()
Dim math As String
Dim video As Boolean

video = False
If addnew = True Then
Text2.Text = "0"
End If

If Option3.value = True Then
math = "+"
End If
If Option4.value = True Then
math = "-"
End If
If Option5.value = True Then
video = True
End If
If Option8.value = True Then
video = True
End If
If Option6.value = True Then
video = True
End If


If video = False Then

Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & _
        "persist security info=False;Data Source=" & App.Path & _
        "\inventory.mdb;Mode = readwrite"
    conAVB.Open
   Set header = New ADODB.Command
    Set header.ActiveConnection = conAVB

 Set rsetheader = New ADODB.Recordset
 If Option7.value = True Then
    rsetheader.Open "Grocery", conAVB, adOpenDynamic, adLockOptimistic, adCmdTable
    
    'Set rsetheader = header.Execute

 
    header.CommandText = "UPDATE Grocery set qty=qty" & math & Text2 & ",print1 = no where Item_Num = '" & Text1 & "'"
    header.Execute
   End If
   End If
   
   
   If video = True Then
   Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & _
        "persist security info=False;Data Source=" & App.Path & _
        "\inventory.mdb;Mode = readwrite"
    conAVB.Open
   Set header = New ADODB.Command
    Set header.ActiveConnection = conAVB

 Set rsetheader = New ADODB.Recordset
If Option8.value = True And Option5.value = False And Option6.value = False Then
    rsetheader.Open "Movie", conAVB, adOpenDynamic, adLockOptimistic, adCmdTable
    
    'Set rsetheader = header.Execute
End If
 If Option5.value = True Then
    header.CommandText = "UPDATE Movie set borrowed=yes,bname = '" & Text3 & "' where Item_Num = '" & Text1 & "'"
    header.Execute
   End If
   If Option6.value = True Then
   header.CommandText = "UPDATE Movie set borrowed = no where Item_Num = '" & Text1 & "'"
    header.Execute
    End If
    If Option3.value = True Then
    header.CommandText = "UPDATE Movie set qty=qty" & math & Text2 & " where Item_Num = '" & Text1 & "'"
    header.Execute
   End If
   If Option4.value = True Then
    header.CommandText = "UPDATE Movie set qty=qty" & math & Text2 & " where Item_Num = '" & Text1 & "'"
    header.Execute
   End If
End If
    
conAVB.Close



Text1 = ""
'Text1.SetFocus

End Sub

Private Sub Command2_Click()
End


End Sub

Private Sub Form_Load()
Text3.Visible = False
Label3.Visible = False
Option3.value = True
Option7.value = True

addnew = False


End Sub

Private Sub mnubmov_Click()
Dim dbForReport_file As String
 Dim cn_ForReport As ADODB.Connection
Dim rs_ForReport As ADODB.Recordset
Dim strsql As String


 dbForReport_file = App.Path & "\inventory.mdb"

    ' Open a connection.
    Set cn_ForReport = New ADODB.Connection
    cn_ForReport.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & dbForReport_file & ";" & _
        "Persist Security Info=False"
    borrowed.WindowState = vbMaximized
    cn_ForReport.Open
    ' Open the Recordset.
   strsql = "SELECT Name,bname from movie where borrowed = yes"
    Set rs_ForReport = cn_ForReport.Execute(strsql, , adCmdText)
    ' Connect the Recordset to the DataReport.
    Set borrowed.DataSource = rs_ForReport
   
   
   borrowed.Caption = "                               Grocery List"
    borrowed.WindowState = vbMaximized
    borrowed.Show
End Sub

Private Sub mnugroc_Click()
Dim dbForReport_file As String
 Dim cn_ForReport As ADODB.Connection
Dim rs_ForReport As ADODB.Recordset
Dim strsql As String


 dbForReport_file = App.Path & "\inventory.mdb"

    ' Open a connection.
    Set cn_ForReport = New ADODB.Connection
    cn_ForReport.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & dbForReport_file & ";" & _
        "Persist Security Info=False"
    grocrpt.WindowState = vbMaximized
    cn_ForReport.Open
    ' Open the Recordset.
   strsql = "SELECT Name,qty from Grocery where qty < 3 and print1 = no "
    Set rs_ForReport = cn_ForReport.Execute(strsql, , adCmdText)
    ' Connect the Recordset to the DataReport.
    Set grocrpt.DataSource = rs_ForReport
   
   
   grocrpt.Caption = "                               Grocery List"
    grocrpt.WindowState = vbMaximized
    grocrpt.Show

 Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & _
        "persist security info=False;Data Source=" & App.Path & _
        "\inventory.mdb;Mode = readwrite"
    conAVB.Open
   Set header = New ADODB.Command
    Set header.ActiveConnection = conAVB

 Set rsetheader = New ADODB.Recordset

    rsetheader.Open "Grocery", conAVB, adOpenDynamic, adLockOptimistic, adCmdTable
    
    'Set rsetheader = header.Execute

 
    header.CommandText = "UPDATE Grocery set print1 = yes where qty < 3"
    header.Execute
  
   

End Sub

Private Sub mnumovoh_Click()
Dim dbForReport_file As String
 Dim cn_ForReport As ADODB.Connection
Dim rs_ForReport As ADODB.Recordset
Dim strsql As String


 dbForReport_file = App.Path & "\inventory.mdb"

    ' Open a connection.
    Set cn_ForReport = New ADODB.Connection
    cn_ForReport.ConnectionString = _
        "Provider=Microsoft.Jet.OLEDB.4.0;" & _
        "Data Source=" & dbForReport_file & ";" & _
        "Persist Security Info=False"
    grocrpt.WindowState = vbMaximized
    cn_ForReport.Open
    ' Open the Recordset.
   strsql = "SELECT Name from Movie Where borrowed = no and sold = no order by Name"
    Set rs_ForReport = cn_ForReport.Execute(strsql, , adCmdText)
    ' Connect the Recordset to the DataReport.
    Set moviehand.DataSource = rs_ForReport
   
   
   moviehand.Caption = "                               Movies on Hand"
    moviehand.WindowState = vbMaximized
    moviehand.Show
End Sub


Private Sub Option3_Click()
Text3.Visible = False
Label3.Visible = False

End Sub

Private Sub Option4_Click()
Text3.Visible = False
Label3.Visible = False

End Sub

Private Sub Option5_Click()
Text3.Visible = True
Label3.Visible = True
Option8.value = True

End Sub


Private Sub Option6_Click()
Option8.value = True

End Sub


Private Sub Text1_LostFocus()
If Option7.value = True Then
 strmov = False
 End If
 If Option8.value = True Then
 strmov = True
 End If
If Text1.Text <> "" Then

If Option1.value = True Then
Me.Text1.Text = CueCatDecode(Text1.Text, 3)
End If
 Dim strsql As String
    Dim strsearch As String
    Dim abort As Boolean
    Set conAVB = New ADODB.Connection
    
    conAVB.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;" & _
        "persist security info=False;Data Source=" & App.Path & _
        "\Inventory.mdb;Mode = readwrite"
    conAVB.Open
    
    
      
    '--------------------------------------------------
    Set header = New ADODB.Command
    Set header.ActiveConnection = conAVB
    
    If Option7.value = True Then
    strsearch = "Select * from Grocery "
    strsearch = strsearch & "Where Item_Num = '" & Text1 & "'"
    End If
    If Option8.value = True Then
    strsearch = "Select * from Movie "
    strsearch = strsearch & "Where Item_Num = '" & Text1 & "'"
    
    
    End If
    header.CommandText = strsearch
    Set rsetheader = header.Execute
    abort = False
   With rsetheader
       ' .MoveFirst ' start at the begining
       ' .Find strSearch
       If .EOF Then
                             
              abort = True ' if the Po and Customer doesn't exist then it continues with out
                            ' filling in the text box's
                            
        End If
        
            'MsgBox "A problem exist in your database", vbExclamation
          '  .MoveFirst
        'End If
        If abort = True Then
        stritem = Text1.Text
        'itemnum = Text1.Text
        addnew = True
    
        ADDItem.Show
        Else
        addnew = False
        
         
        End If
    End With
    conAVB.Close
    'Text2.SetFocus
 Else
 End If
 
 
 
End Sub



Public Property Get itemnum() As String

End Property

Public Property Let itemnum(ByVal vNewValue As String)

'item = Text1.Text

vNewValue = item
ADDItem.Show



End Property

Private Sub Text2_LostFocus()
If Option5.value = True Then
    Text3.SetFocus
Else
Command1.SetFocus
End If


End Sub




