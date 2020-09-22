Attribute VB_Name = "modMain"
Private BitwiseOperator As New clsBitwiseOperator
Private lngCRC32Table(255) As Long

Public Function InitializeCRC32Table(Optional ByVal lngPolynomial As Long = &HEDB88320, Optional ByVal lngInitializeValue As Long = &HFFFFFFFF) As Long
    Dim lngBytes As Long, lngBits As Long
    Dim lngCRC32 As Long, lngTemporaryCRC32 As Long

    '// cycle through every byte available and generate a lookup table
    For lngBytes = 0 To 255

        lngCRC32 = lngBytes

        For lngBits = 0 To 7
            lngTemporaryCRC32 = BitwiseOperator.ShiftRight(lngCRC32, 1)

            If (lngCRC32 And &H1) Then
                lngCRC32 = lngTemporaryCRC32 Xor lngPolynomial
            Else
                lngCRC32 = lngTemporaryCRC32
            End If
        Next lngBits

        lngCRC32Table(lngBytes) = lngCRC32
    Next lngBytes

    InitializeCRC32Table = lngInitializeValue
End Function

Public Function GenerateCRC32(ByVal strBuffer As String, ByVal lngInitializeValue As Long) As Long
    Dim lngBytes As Long, bytCurrentByte As Byte
    Dim lngActualValue As Long
    Dim lngTableValue As Long, lngIndex As Long
    
    For lngBytes = 1 To Len(strBuffer)
        
        bytCurrentByte = Asc(Mid$(strBuffer, lngBytes, 1))
        
        lngActualValue = BitwiseOperator.ShiftRight(lngInitializeValue, 8)
        
        lngIndex = lngInitializeValue And &HFF
        lngIndex = lngIndex Xor bytCurrentByte
        lngTableValue = lngCRC32Table(lngIndex)
        
        lngInitializeValue = lngActualValue Xor lngTableValue
    Next lngBytes
    
    GenerateCRC32 = lngInitializeValue Xor &HFFFFFFFF
End Function
