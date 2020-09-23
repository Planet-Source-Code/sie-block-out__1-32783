Attribute VB_Name = "modEncryption"
Option Explicit
   
Private mStrKey
Private mBytKeyAry(255)
Private mBytCypherAry(255)

Private Sub InitializeCypher()
    
    Dim lBytJump
    Dim lBytIndex
    Dim lBytTemp

    For lBytIndex = 0 To 255
        mBytCypherAry(lBytIndex) = lBytIndex
    Next
    ' Switch values of Cypher arround based off of index and Key value
    lBytJump = 0
    For lBytIndex = 0 To 255

        ' Figure index To switch
        lBytJump = (lBytJump + mBytCypherAry(lBytIndex) + mBytKeyAry(lBytIndex)) Mod 256
    
        ' Do the switch
        lBytTemp = mBytCypherAry(lBytIndex)
        mBytCypherAry(lBytIndex) = mBytCypherAry(lBytJump)
        mBytCypherAry(lBytJump) = lBytTemp
    
    Next
End Sub

Public Function Crypt(ByRef pStrMessage) As String
    Dim lBytIndex
    Dim lBytJump
    Dim lBytTemp
    Dim lBytY
    Dim lLngT
    Dim lLngX
    
    'Validate data
    If Len(pStrMessage) = 0 Then Exit Function
    Call InitializeCypher
    
    lBytIndex = 0
    lBytJump = 0
    For lLngX = 1 To Len(pStrMessage)
        lBytIndex = (lBytIndex + 1) Mod 256 ' wrap index
        lBytJump = (lBytJump + mBytCypherAry(lBytIndex)) Mod 256 ' wrap J+S()
        
            ' Add/Wrap those two
        lLngT = (mBytCypherAry(lBytIndex) + mBytCypherAry(lBytJump)) Mod 256
        
        'Switch
        lBytTemp = mBytCypherAry(lBytIndex)
        mBytCypherAry(lBytIndex) = mBytCypherAry(lBytJump)
        mBytCypherAry(lBytJump) = lBytTemp
        lBytY = mBytCypherAry(lLngT)
            ' Character Encryption ...
        Crypt = Crypt & Chr(Asc(Mid(pStrMessage, lLngX, 1)) Xor lBytY)
    Next
    
End Function

