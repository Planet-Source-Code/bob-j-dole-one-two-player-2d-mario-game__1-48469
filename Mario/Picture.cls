VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Picture"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'local variable(s) to hold property value(s)
Private mvarpicmask As StdPicture 'local copy
Private mvarpicsource As StdPicture 'local copy


Public Property Set picsource(ByVal vData As StdPicture)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.picsource = Form1
    Set mvarpicsource = vData
End Property


Public Property Get picsource() As StdPicture
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.picsource
    Set picsource = mvarpicsource
End Property



Public Property Set picmask(ByVal vData As StdPicture)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.picmask = Form1
    Set mvarpicmask = vData
End Property


Public Property Get picmask() As StdPicture
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.picmask
    Set picmask = mvarpicmask
End Property





Public Sub SetMsk(adress As String)
'Set Badguy(3).picsource = New StdPicture
'Set Badguy(3).picsource = LoadPicture(App.Path + "\Pics\turtsspr.bmp")
'Badguy(3).hsourcedc = CreateCompatibleDC(HTempDC)
'HOldbmp = SelectObject(Badguy(3).hsourcedc, Badguy(3).picsource.Handle)

End Sub
Public Sub SetPicsource(adress As String)
Attribute SetPicsource.VB_Description = "Set the picsource"
End Sub

