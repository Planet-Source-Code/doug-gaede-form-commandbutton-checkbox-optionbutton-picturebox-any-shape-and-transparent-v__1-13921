Attribute VB_Name = "Module1"
Option Explicit

Sub Main()
'always start your program with a Sub Main as part of good coding practice.

Form1.Show
Form2.Show
DoEvents
Form1.SetFocus
'note that DoEvents is necessary to ensure the form
'has completed drawing if you show another form after this one.
'Otherwise the CommandButtons do not form completely.
'Uncomment the code below and comment the code above, then
'run this with and without the DoEvents to demonstrate this.

'Form2.Show
'DoEvents
'Form1.Show

'DoEvents is a kludge, but subclassing within the class is
'necessary to overcome this and that won't be until version 3.

End Sub
