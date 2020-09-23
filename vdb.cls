VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "vdb"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'This is the VDB class that will store the properties for the controls
'This does not have much funcionality but is rather intended to keep track of
'all that stuff
Public nm As String 'variable for the control name property
Public frm As String 'variable for the form name property
Public tp As Integer 'variable for the top property
Public lft As Integer 'variable for the left property
Public fntsize As Integer 'variable for the font size property
Public x1 As Integer 'variable for the x1 position of a line
Public x2 As Integer 'variable for the x2 position of a line
Public y1 As Integer 'variable for the y1 position of a line
Public y2 As Integer 'variable for the y2 position of a line
Public w As Integer 'variable for the width property
Public h As Integer 'variable for the height property
Public fw As Integer 'variable for the form width property
Public fh As Integer 'variable for the form height property

