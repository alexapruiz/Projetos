VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private mCol As Collection

Public Property Get Count() As Long
    'used when retrieving the number of elements in the
    'collection. Syntax: Debug.Print x.Count
    Count = mCol.Count
End Property


Public Property Get Item(vntIndexKey As Variant) As CalculoCheque
    'used when referencing an element in the collection
    'vntIndexKey contains either the Index or Key to the collection,
    'this is why it is declared as a Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property




Public Property Get NewEnum() As IUnknown
    'this property allows you to enumerate
    'this collection with the For...Each syntax
    Set NewEnum = mCol.[_NewEnum]
End Property


Public Sub Remove(vntIndexKey As Variant)
    'used when removing an element from the collection
    'vntIndexKey contains either the Index or Key, which is why
    'it is declared as a Variant
    'Syntax: x.Remove(xyz)


    mCol.Remove vntIndexKey
End Sub



Public Function Add(ByVal pCMC7 As String, Optional sKey As String) As CalculoCheque
    'create a new object
    Dim objNewMember As CalculoCheque
    Set objNewMember = New CalculoCheque

    'set the properties passed into the method
'    objNewMember.ValorDeposito = ValorDeposito
'    objNewMember.Quantidade = Quantidade
'    objNewMember.IdBordero = IdBordero
'    objNewMember.DataProcessamento = DataProcessamento

    objNewMember.CMC7 = pCMC7
    'objNewMember.DadosDivergentes = DadosDivergentes
    If Len(sKey) = 0 Then
        mCol.Add objNewMember
    Else
        mCol.Add objNewMember, sKey
    End If
    
    'return the object created
    Set Add = objNewMember
    Set objNewMember = Nothing


End Function


Private Sub Class_Initialize()

    Set mCol = New Collection

End Sub


Private Sub Class_Terminate()

    Set mCol = Nothing

End Sub


