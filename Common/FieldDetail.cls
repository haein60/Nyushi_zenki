VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FieldDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'*************************************************************************************************
'Form Name      :   FieldDetail
'Author         :   Dileep Cherian
'Created On     :   10/8/01
'Description    :   This form makes a provision for master maintenance of tbSTRZipCode Table.
'Reference      :   Functional Specs Of MasterMaintenance Ver 1.0
'**************************************************************************************************

'local variable(s) to hold property value(s)
Private mvarDBFieldName As String 'local copy
Private mvarSCRFieldName As String 'local copy
Private mvarGridColPos As Integer 'local copy
Private mvarbMandatory As Boolean 'local copy
Private mvarbPrimaryKey As Boolean 'local copy
Private mvarstrValue As String 'local copy
Private mvarstrDataType As String 'local copy
Private mvarColWidth As Long 'local copy     - added by team
Private mvarstrErrorMsg As String 'local copy
Private mvarstrDataFormat As String
Private mvarDBReadFieldName As String

Public Property Let strErrorMsg(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strErrorMsg = 5
    mvarstrErrorMsg = vData
End Property


Public Property Get strErrorMsg() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strErrorMsg
    strErrorMsg = mvarstrErrorMsg
End Property



Public Property Let strDataType(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strDataType = 5
    mvarstrDataType = vData
End Property


Public Property Get strDataType() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strDataType
    strDataType = mvarstrDataType
End Property

'Public Property Let strDataFormat(ByVal vData As String) ByVal‚¾‚Æ‚Í‚¢‚ç‚È‚¢
Public Property Let strDataFormat(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strDataType = 5
    mvarstrDataFormat = vData
End Property

Public Property Get strDataFormat() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strDataType
    strDataFormat = mvarstrDataFormat
End Property


Public Property Let strValue(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.strValue = 5
    mvarstrValue = vData
End Property


Public Property Get strValue() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.strValue
    strValue = mvarstrValue
End Property



Public Property Let bMandatory(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bMandatory = 5
    mvarbMandatory = vData
End Property


Public Property Get bMandatory() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bMandatory
    bMandatory = mvarbMandatory
End Property

Public Property Let PrimaryKey(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.bMandatory = 5
    mvarbPrimaryKey = vData
End Property


Public Property Get PrimaryKey() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.bMandatory
    PrimaryKey = mvarbPrimaryKey
End Property


Public Property Let GridColPos(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.GridColPos = 5
    mvarGridColPos = vData
End Property


Public Property Get GridColPos() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.GridColPos
    GridColPos = mvarGridColPos
End Property



Public Property Let SCRFieldName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.SCRFieldName = 5
    mvarSCRFieldName = vData
End Property


Public Property Get SCRFieldName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.SCRFieldName
    SCRFieldName = mvarSCRFieldName
End Property

Public Property Let DBFieldName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DBFieldName = 5
    mvarDBFieldName = vData
End Property

Public Property Get DBFieldName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DBFieldName
     DBFieldName = mvarDBFieldName
End Property

Public Property Let DBReadFieldName(vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DBFieldName = 5
    mvarDBReadFieldName = vData
End Property

Public Property Get DBReadFieldName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DBFieldName
     DBReadFieldName = mvarDBReadFieldName
End Property

Public Property Get ColWidth() As Long
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DBFieldName
     ColWidth = mvarColWidth
End Property

Public Property Let ColWidth(ByVal vData As Long)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DBFieldName = 5
    mvarColWidth = vData
End Property

