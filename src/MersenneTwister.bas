Attribute VB_Name = "MersenneTwister"
'http://www.math.sci.hiroshima-u.ac.jp/~m-mat/MT/VERSIONS/BASIC/visualbasicMT.txt
'
' Visual Basic Mersenne-Twister
' Author: Carmine Arturo Sangiovanni
'         carmine @ daygo.com.br
'
'         Aug 13,2004
'
'         based on C++ code
'



Option Explicit

Const N = 624
Const M = 397

Global mt(0 To N) As Currency
Global mti As Currency

Dim MATRIX_A As Currency
Dim UPPER_MASK As Currency
Dim LOWER_MASK As Currency
Dim FULL_MASK As Currency
Dim TEMPERING_MASK_B As Currency
Dim TEMPERING_MASK_C As Currency

Function tempering_shift_u(ty As Currency)
'>> 11'  => 2^11 => 2048
    tempering_shift_u = f_and(Int(ty / 2048@), FULL_MASK)
End Function

Function tempering_shift_s(ty As Currency)
'<<  7   => 2^7 => 128
    tempering_shift_s = and_ffffffff(ty * 128@)
End Function

Function tempering_shift_t(ty As Currency)
'<<  15   => 2^15 => 32768
    tempering_shift_t = and_ffffffff(ty * 32768@)
End Function

Function tempering_shift_l(ty As Currency)
'>> 18'  => 2^18 =>262144
    tempering_shift_l = f_and(Int(ty / 262144@), FULL_MASK)
End Function

Function f_and(p1 As Currency, p2 As Currency)
    Dim v As Currency
    Dim i As Integer
    
    
    If (p1 < UPPER_MASK) Then
      
      If (p2 < UPPER_MASK) Then
         f_and = p1 And p2
      Else
         f_and = p1 And (p2 - UPPER_MASK)
      End If
      
    Else
    
      If (p2 < UPPER_MASK) Then
         f_and = (p1 - UPPER_MASK) And p2
      Else
         f_and = (p1 - UPPER_MASK) And (p2 - UPPER_MASK)
         f_and = f_and + UPPER_MASK
      End If
   End If
End Function

Function f_or(p1 As Currency, p2 As Currency)
    Dim v As Currency
    Dim i As Integer
    Dim f As Boolean
    
    
    If (p1 < UPPER_MASK) Then
      
      If (p2 < UPPER_MASK) Then
          f_or = p1 Or p2
      Else
          f_or = p1 Or (p2 - UPPER_MASK)
          f_or = f_or + UPPER_MASK
      End If
      
    Else
    
      If (p2 < UPPER_MASK) Then
          f_or = (p1 - UPPER_MASK) Or p2
          f_or = f_or + UPPER_MASK
      Else
          f_or = (p1 - UPPER_MASK) Or (p2 - UPPER_MASK)
          f_or = f_or + UPPER_MASK
      
      End If
      
    End If
End Function

Function LongUnSignedToCurrency(value As Currency) As Currency
   If value < 0 Then
      Const Val_0x100000000 = 4294967296#
      LongUnSignedToCurrency = value + Val_0x100000000
   End If
End Function


Function f_xor(p1 As Currency, p2 As Currency)
    Dim v As Currency
    Dim i As Integer

    If (p1 < UPPER_MASK) Then
      
      If (p2 < UPPER_MASK) Then
         f_xor = p1 Xor p2
      Else
         f_xor = p1 Xor (p2 - UPPER_MASK)
         f_xor = f_xor + UPPER_MASK
      End If
      
    Else
    
      If (p2 < UPPER_MASK) Then
         f_xor = (p1 - UPPER_MASK) Xor p2
         f_xor = f_xor + UPPER_MASK
      Else
         f_xor = (p1 - UPPER_MASK) Xor (p2 - UPPER_MASK)
      End If
      
    End If
    
    
End Function

'Bug 3: ByVal required to avoid unwanted sideeffects
Function f_lower(ByVal p1 As Currency)
    Do
        If p1 < UPPER_MASK Then
            f_lower = p1
            Exit Do
        Else
            p1 = p1 - UPPER_MASK
        End If
    Loop
End Function

Function f_upper(p1 As Currency)
    If p1 > LOWER_MASK Then
        f_upper = UPPER_MASK
    Else
        f_upper = 0
    End If
End Function

Function f_xor3(p1 As Currency, p2 As Currency, p3 As Currency) As Currency
    Dim v As Currency
    Dim tmp As Currency
    Dim i As Integer
    Dim f As Integer
    
    
    
    If (p1 < UPPER_MASK) Then
      
      If (p2 < UPPER_MASK) Then
         tmp = p1 Xor p2
      Else
         tmp = p1 Xor (p2 - UPPER_MASK)
         tmp = tmp + UPPER_MASK
      End If
      
    Else
    
      If (p2 < UPPER_MASK) Then
         tmp = (p1 - UPPER_MASK) Xor p2
         tmp = tmp + UPPER_MASK
      Else
         tmp = (p1 - UPPER_MASK) Xor (p2 - UPPER_MASK)
      End If
      
    End If
    
    
    
    
    If (tmp < UPPER_MASK) And (p3 < UPPER_MASK) Then
        f_xor3 = tmp Xor p3
    End If
    If (tmp < UPPER_MASK) And (p3 >= UPPER_MASK) Then
        f_xor3 = tmp Xor (p3 - UPPER_MASK)
        f_xor3 = f_xor3 + UPPER_MASK
    End If
    If (tmp >= UPPER_MASK) And (p3 < UPPER_MASK) Then
        f_xor3 = (tmp - UPPER_MASK) Xor p3
        f_xor3 = f_xor3 + UPPER_MASK
    End If
    If (tmp >= UPPER_MASK) And (p3 >= UPPER_MASK) Then
        f_xor3 = (tmp - UPPER_MASK) Xor (p3 - UPPER_MASK)
    End If
    
End Function

Function and_ffffffff(c As Currency)
    Dim e As Currency
    Dim i As Integer
    'Delete each bit before 32Bit (=>2^15)
    i = 32
    Do
        e = 2 ^ (i + 16)
        Do While c >= e
            c = c - e
        Loop
        i = i - 1
    Loop While i > 15
    
    and_ffffffff = c
End Function

'Sub random_init(key As Long)
'   Dim seed As Currency
'   If key < 0 Then
'      seed = key + 4294967296# '0x100000000
'   Else
'      seed = key
'   End If
'
'
'    mt(0) = and_ffffffff(seed)
'    For mti = 1 To N - 1
'        mt(mti) = and_ffffffff(69069 * mt(mti - 1))
'    Next mti
'End Sub
Sub random_init(seed As Long)

  'int with  16Bit Values
   Dim tmp As New StringReader
   With tmp
      .Data = Space(&H9C0)

      Dim newKey&
      newKey = seed
      .EOS = False
      Do Until .EOS

'         Dim i
        
        'Store HighPart
         Dim Key32Bit_HighPart&
         Key32Bit_HighPart = shr16(newKey, &H10)

'Debug.Print H32(Key32Bit_HighPart), i, H16(i * 2)
'Debug.Assert i <> 470 * 2
         .int16 = Key32Bit_HighPart


         newKey = MulInt32(newKey, 69069) + 1
         
'         Inc i
         
      Loop


   'Write 32Bit Values
   .EOS = False
    For mti = 0 To N - 1
    
       Dim L@, h@
       h = .int16
       L = .int16
       mt(mti) = h * CCur(&H10000) + L
    Next mti

   End With



End Sub


Function Mersenne_twister_random(Optional BitMask As Long = &HFFFFFFFF, Optional XorKey As Currency = &H0) As Long

'    XorKey = LongUnSignedToCurrency(XorKey)
   
   'f_Xor will not come along with negative values
    If XorKey < 0 Then
      Static alreadyShow As Boolean
      If alreadyShow = False Then
         log "-> WARNING: Stupid Bug - Mersenne_twister_random() does not come along with a XorKey negative values! Decryption will fail!"
         alreadyShow = True
      End If
      
'      Exit Function
      
    End If
    
    Dim kk As Integer
    
    Dim ty1 As Currency
    Dim ty2 As Currency
    Dim y As Currency
    
    Dim mag01(0 To 1) As Currency
    
    MATRIX_A = 2567483615@              '&H9908b0df
    UPPER_MASK = 2147483648@            '&H80000000
    LOWER_MASK = 2147483647@            '&H7fffffff
    FULL_MASK = LOWER_MASK + UPPER_MASK '&Hffffffff
    TEMPERING_MASK_B = 2636928640@      '&H9D2C5680 (<= &HFF3A58AD << 0x7 )
    TEMPERING_MASK_C = 4022730752@      '&HEFC60000 (<= &HFFFFDF8C << 0xF )

    
    mag01(0) = 0@
    mag01(1) = MATRIX_A
    
    If mti >= N Then
        'Init already called?
        If mti = N + 1 Then
            random_init 4537
        End If
        
'         009F1D80   2616263785  1485109129  1222459022  1470865537
'         009F1D90   2221574427    34070531  2495128680  3752208502
'         009F1DA0   3221593492  3182452335  2992174400  1056224458
'         009F1DB0    736937993  2469744774  1683846344  2704511074
'         009F1DC0   1318321291  2944046138  2842695760  1475108897
'         009F1DD0   2302328966  1376761661  3318325827  1612024421
'
        
        For kk = 0 To (N - M) - 1
'        Debug.Assert kk <> 73
            y = f_or(f_upper(mt(kk)), f_lower(mt(kk + 1)))
            '3632592777=D884F789
            
            mt(kk) = f_xor3(mt(kk + M), Int(y / 2@), mag01(f_and(y, 1)))
            'mt(3)? 473F343F 1195324479 bad
            '        73F343F  121582655 good
            'mt(74)?  1373583696 Good
            '         1372600656 Bad


        Next kk
        
        For kk = kk To (N - 1) - 1
            y = f_or(f_upper(mt(kk)), f_lower(mt(kk + 1)))
            
            mt(kk) = f_xor3(mt(kk + (M - N)), Int(y / 2@), mag01(f_and(y, 1)))
        Next kk
        
        y = f_or(f_upper(mt(N - 1)), f_lower(mt(0)))
        
        mt(N - 1) = f_xor3(mt(M - 1), Int(y / 2@), mag01(f_and(y, 1)))
        mti = 0
    End If
    
    '---------------------------------------------------
    y = mt(mti): mti = mti + 1
    
    '---------------------------------------------------
    '
' IONCUBE Specific added line!
    y = f_xor(y, XorKey)
    'y 8B3D956F -1958898321
    'XorKey 3555331
    '-1834523451
    
    
    y = f_xor(y, tempering_shift_u(y))
    
    ty1 = f_and(tempering_shift_s(y), TEMPERING_MASK_B)
    y = f_xor(y, ty1)
    
    ty1 = f_and(tempering_shift_t(y), TEMPERING_MASK_C)
    y = f_xor(y, ty1)
    
    y = f_xor(y, tempering_shift_l(y))
    
    '---------------------------------------------------
  
  'Convert to long
   If y >= UPPER_MASK Then
      Mersenne_twister_random = -(4294967296@ - y) '(2 * UPPER_MASK) - y
   Else
      Mersenne_twister_random = y
   End If
  
  'Apply Bitmask
   Mersenne_twister_random = Mersenne_twister_random And BitMask
 
End Function

