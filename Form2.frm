VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form2 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00008000&
   Caption         =   "Form2"
   ClientHeight    =   6465
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5925
   LinkTopic       =   "Form2"
   ScaleHeight     =   6465
   ScaleWidth      =   5925
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option2 
      DownPicture     =   "Form2.frx":0000
      Height          =   975
      Left            =   3720
      Picture         =   "Form2.frx":02BE
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   2040
      Width           =   735
   End
   Begin VB.OptionButton Option1 
      DownPicture     =   "Form2.frx":0590
      Height          =   975
      Left            =   2880
      Picture         =   "Form2.frx":084E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   2040
      Width           =   735
   End
   Begin VB.CheckBox Check1 
      DownPicture     =   "Form2.frx":0B20
      Height          =   975
      Left            =   1560
      Picture         =   "Form2.frx":0DDE
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   2040
      Width           =   735
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   100
      ImageHeight     =   100
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":10B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form2.frx":12DA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      DownPicture     =   "Form2.frx":150F
      Height          =   975
      Left            =   4440
      Picture         =   "Form2.frx":17CD
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DownPicture     =   "Form2.frx":1A9F
      Height          =   975
      Index           =   1
      Left            =   3360
      Picture         =   "Form2.frx":1D5D
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   240
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      DownPicture     =   "Form2.frx":202F
      Height          =   975
      Index           =   0
      Left            =   2520
      Picture         =   "Form2.frx":22ED
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   240
      Width           =   615
   End
   Begin VB.PictureBox picButton 
      AutoRedraw      =   -1  'True
      BackColor       =   &H008080FF&
      BorderStyle     =   0  'None
      ClipControls    =   0   'False
      Height          =   1455
      Left            =   720
      Picture         =   "Form2.frx":25BF
      ScaleHeight     =   1455
      ScaleWidth      =   1455
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      Begin VB.TextBox Text1 
         Alignment       =   2  'Center
         Height          =   735
         Left            =   360
         MultiLine       =   -1  'True
         TabIndex        =   5
         Text            =   "Form2.frx":27D9
         Top             =   728
         Width           =   855
      End
   End
   Begin VB.Label Label2 
      Caption         =   "Two OptionButtons"
      Height          =   255
      Index           =   5
      Left            =   3000
      TabIndex        =   11
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "CheckBox control"
      Height          =   255
      Index           =   4
      Left            =   1200
      TabIndex        =   10
      Top             =   3120
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "CommandButton by itself."
      Height          =   495
      Index           =   2
      Left            =   4080
      TabIndex        =   8
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "CommandButton ctrls. in a ctrl array."
      Height          =   495
      Index           =   1
      Left            =   2520
      TabIndex        =   7
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "PictureBox control"
      Height          =   255
      Index           =   0
      Left            =   720
      TabIndex        =   6
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label1 
      Height          =   2655
      Left            =   1320
      TabIndex        =   1
      Top             =   3600
      Width           =   3255
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ShapeTheControls As clsTransForm 'make a reference to the class

Private Sub Form_Activate()
Static booAlreadyDone As Boolean
'Shaping buttons does not work in Form_Load, so must be here.
'Has something to do with how VB makes CommandButtons, CheckBoxes and OptionButtons.
'Form_Activate happens many times over the life of the form,
'so make sure this code only runs once or weird things happen (try it).

Form2.SetFocus 'if a button has focus the focus ring does not become
'transparent, so set the focus to the form first

If Not booAlreadyDone Then 'if booAlreadyDone = False then
    'It appears there is a glitch if you try to load region data from a file for real command buttons, so
    'you must do them the hard way.  Remove the comment ticks to see what I mean.
    ShapeTheControls.ShapeMe Command1(0), RGB(255, 255, 255) ', True, App.Path & "\Form2RegionData1.dat"
    ShapeTheControls.ShapeMe Command1(1), RGB(255, 255, 255) ', True, App.Path & "\Form2RegionData1.dat"
    ShapeTheControls.ShapeMe Command2, RGB(255, 255, 255) ', True, App.Path & "\Form2RegionData1.dat"
    ShapeTheControls.ShapeMe Check1, RGB(255, 255, 255), True, App.Path & "\Form2RegionData1.dat"
    ShapeTheControls.ShapeMe Option1, RGB(255, 255, 255), True, App.Path & "\Form2RegionData1.dat"
    ShapeTheControls.ShapeMe Option2, RGB(255, 255, 255), True, App.Path & "\Form2RegionData1.dat"
    booAlreadyDone = Not booAlreadyDone 'make it True
End If

End Sub

Private Sub Form_Load()
Set ShapeTheControls = New clsTransForm 'instantiate the object from the class

Label1 = "Notice the parts of the TextBox not on the black square are not visible.  Also, the PictureBox has a slight delay if you click it fast enough.  I even tried using the BitBlt API, but that did not get rid of the problem.  That forced me to figure out how to shape a real button.  The one quirk is that it appears VB adds highlighting to every type of button, thus the white lines when pressed.  If you give your buttons a white border then it shouldn't be noticeable.  The CheckBox and OptionButtons are the same as a CommandButton internally."

'shape the picturebox
ShapeTheControls.ShapeMe picButton, RGB(255, 255, 255), True, App.Path & "\Form2RegionData2.dat"

End Sub

Private Sub Form_Unload(Cancel As Integer)

Set ShapeTheControls = Nothing 'destroy the object
Set Form2 = Nothing 'good practice to free resources VB doesn't normally free when you unload a form!

End Sub

Private Sub picButton_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

Set picButton.Picture = ImageList1.ListImages(2).Picture
Text1 = "In"

End Sub

Private Sub picButton_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Set picButton.Picture = ImageList1.ListImages(1).Picture
Text1 = "Out"

End Sub
