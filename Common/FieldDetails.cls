VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FieldDetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Collection" ,"FieldDetail"
Attribute VB_Ext_KEY = "Member0" ,"FieldDetail"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
'*************************************************************************************************
'Form Name      :   FieldDetails
'Author         :   Dileep Cherian
'Created On     :   10/8/01
'Description    :   This form makes a provision for master maintenance of tbSTRZipCode Table.
'Reference      :   Functional Specs Of MasterMaintenance Ver 1.0
'**************************************************************************************************
'local variable to hold collection
Private mCol As Collection

Public Function Add(DBFieldName As String, SCRFieldName As String, GridColPos As Integer, bMandatory As Boolean, PrimaryKey As Boolean, strValue As String, strDataType As String, ColWidth As Long, strErrorMsg As String, Optional sKey As String, Optional DBReadFieldName As String, Optional sDataFormat As String) As FieldDetail
    'create a new object
    Dim objNewMember As FieldDetail
    Set objNewMember = New FieldDetail

    'set the properties passed into the method
    objNewMember.DBFieldName = DBFieldName
    objNewMember.SCRFieldName = SCRFieldName
    objNewMember.GridColPos = GridColPos
    objNewMember.bMandatory = bMandatory
    objNewMember.PrimaryKey = PrimaryKey
    objNewMember.strValue = strValue
    objNewMember.strDataType = strDataType
    objNewMember.ColWidth = ColWidth
    objNewMember.strErrorMsg = strErrorMsg
    If Len(DBReadFieldName) = 0 Then
        objNewMember.DBReadFieldName = DBFieldName
    Else
        objNewMember.DBReadFieldName = DBReadFieldName
    End If
    If Len(sDataFormat) = 0 Then
        objNewMember.strDataFormat = ""
    Else
        objNewMember.strDataFormat = sDataFormat
    End If
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If


    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function

Public Property Get Item(vntIndexKey As Variant) As FieldDetail
Attribute Item.VB_UserMemId = 0
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub


Public Property Get NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
Attribute NewEnum.VB_MemberFlags = "40"
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Private Sub Class_Initialize()
    'creates the collection when this class is created
    Set mCol = New Collection
End Sub


Private Sub Class_Terminate()
    'destroys collection when this class is terminated
    Set mCol = Nothing
End Sub

