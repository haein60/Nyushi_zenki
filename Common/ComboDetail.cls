VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ComboDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'*************************************************************************************************
'Form Name      :   ComboDetail
'Author         :   Dileep Cherian
'Created On     :   10/8/01
'Description    :   This form makes a provision for master maintenance of tbSTRZipCode Table.
'Reference      :   Functional Specs Of MasterMaintenance Ver 1.0
'**************************************************************************************************

'local variable(s) to hold property value(s)
Private mvarValue As Long 'local copy
Private mvarDescription As String 'local copy
Private mvarGroupId As String  'local copy
Private mvarGroupValue As Long  'local copy

Public Property Let GroupValue(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strErrorMsg = 5
    mvarGroupValue = vData
End Property


Public Property Get GroupValue() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strErrorMsg
    GroupValue = mvarGroupValue
End Property

Public Property Let Value(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strErrorMsg = 5
    mvarValue = vData
End Property


Public Property Get Value() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strErrorMsg
    Value = mvarValue
End Property

Public Property Let Description(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strErrorMsg = 5
    mvarDescription = vData
End Property


Public Property Get Description() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strErrorMsg
    Description = mvarDescription
End Property
Public Property Let GroupId(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strErrorMsg = 5
'    mvarDescription = vData
    mvarGroupId = vData
End Property

Public Property Get GroupId() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strErrorMsg
'    Description = mvarDescription
    GroupId = mvarGroupId
End Property

