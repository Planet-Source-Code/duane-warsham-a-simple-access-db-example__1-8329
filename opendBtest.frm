VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Access test.mdb example."
   ClientHeight    =   3195
   ClientLeft      =   150
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Height          =   315
      Left            =   1733
      TabIndex        =   5
      Text            =   "Text6"
      Top             =   2340
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   315
      Left            =   1733
      TabIndex        =   4
      Text            =   "Text5"
      Top             =   1902
      Width           =   2055
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1733
      TabIndex        =   3
      Text            =   "Text4"
      Top             =   1464
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Height          =   315
      Left            =   1733
      TabIndex        =   2
      Text            =   "Text3"
      Top             =   1026
      Width           =   2055
   End
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   1733
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   588
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      DataField       =   "Name"
      DataSource      =   "rstInfo"
      Height          =   315
      Left            =   1733
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   150
      Width           =   2055
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      Caption         =   "Phone #: "
      Height          =   255
      Left            =   953
      TabIndex        =   12
      Top             =   2400
      Width           =   795
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Zip Code: "
      Height          =   240
      Left            =   893
      TabIndex        =   11
      Top             =   1950
      Width           =   825
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "State: "
      Height          =   255
      Left            =   1223
      TabIndex        =   10
      Top             =   1500
      Width           =   525
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "City: "
      Height          =   255
      Left            =   1283
      TabIndex        =   9
      Top             =   1050
      Width           =   465
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Address: "
      Height          =   270
      Left            =   1013
      TabIndex        =   8
      Top             =   630
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Name: "
      Height          =   255
      Left            =   1163
      TabIndex        =   7
      Top             =   180
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Label1"
      Height          =   255
      Left            =   1463
      TabIndex        =   6
      Top             =   2850
      Width           =   1755
   End
   Begin VB.Menu Clear 
      Caption         =   "Clear"
   End
   Begin VB.Menu Next 
      Caption         =   "Next"
   End
   Begin VB.Menu Previous 
      Caption         =   "Previous"
   End
   Begin VB.Menu Add 
      Caption         =   "Add"
   End
   Begin VB.Menu Delete 
      Caption         =   "Delete"
   End
   Begin VB.Menu Exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Written by Duane Warsham, BKN Computer Services,in VB 6.0.
'This is a simple example of using VB to View a Microsoft Access
'data base on a local machine.
'Be sure to include in your project references the Microsoft
'DAO 3.6 Object Library.
'Be sure to copy the test.mdb to your my documents folder or
'you can change the location path below to its unzipped location.
'The menu editor was used on the form instead of using command buttons.
'Feel free to change and edit as you wish.
'===============================================================
Option Explicit
Public db As Database
Public rstInfo As Recordset
'Run this sub on startup.
Private Sub Form_Load()
Set db = OpenDatabase("c:\my documents\test.mdb")
With db

    Set rstInfo = .OpenRecordset("Info")
    Label1.Caption = rstInfo.RecordCount & " records"
    If rstInfo.RecordCount = 0 Then Exit Sub
    With rstInfo
        .Edit
        Text1.Text = !Name
        Text2.Text = !address
        Text3.Text = !city
        Text4.Text = !state
        Text5.Text = !zip
        Text6.Text = !phone
    End With
End With
End Sub
'Just clear the text boxes.
Private Sub Clear_Click()
        Text1.Text = ""
        Text2.Text = ""
        Text3.Text = ""
        Text4.Text = ""
        Text5.Text = ""
        Text6.Text = ""
End Sub
'This steps forward in the file.  Check for a EOF (end of file)
'status of true, if true, stay at last record, else display it.
Private Sub Next_Click()
With rstInfo
    .MoveNext
    If rstInfo.EOF Then
        MsgBox "End of file.", vbOKOnly, "  Error!"
        .MoveLast
        Exit Sub
    Else
        .Edit
            Text1.Text = !Name
            Text2.Text = !address
            Text3.Text = !city
            Text4.Text = !state
            Text5.Text = !zip
            Text6.Text = !phone
    End If
End With
End Sub
'This steps backward in the file.  Check for a BOF (beginning of file)
'status of true, if true, stay at first record, else display it.
Private Sub Previous_Click()
With rstInfo
    .MovePrevious
    If rstInfo.BOF Then
        MsgBox "Front End of file.", vbOKOnly, "  Error!"
        .MoveFirst
        Exit Sub
    Else
        .Edit
            Text1.Text = !Name
            Text2.Text = !address
            Text3.Text = !city
            Text4.Text = !state
            Text5.Text = !zip
            Text6.Text = !phone
    End If
End With
End Sub
'This will insert a new record the current text box data to
'a new record.
Private Sub Add_Click()
    With rstInfo
        .AddNew
         !Name = Text1.Text
         !address = Text2.Text
         !city = Text3.Text
         !state = Text4.Text
         !zip = Text5.Text
         !phone = Text6.Text
        .Update
    End With
Call Form_Load
End Sub
'This deletes the current record being viewed.
Private Sub Delete_Click()
    With rstInfo
        .Delete
    End With
Call Form_Load
End Sub
'Close database and end the program.
Private Sub Exit_Click()
db.Close
End
End Sub

