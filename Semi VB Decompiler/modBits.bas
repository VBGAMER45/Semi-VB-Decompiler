Attribute VB_Name = "HiLo"
' *********************************************************************
'  Copyright ©1994-2000 Karl E. Peterson, All Rights Reserved
'  http://www.mvps.org/vb
' *********************************************************************
'  You are free to use this code within your own applications, but you
'  are expressly forbidden from selling or otherwise distributing this
'  source code without prior written consent.
' *********************************************************************
'  Most of the routines in this module were written to be self-
'  contained. If you routinely drop the entire module into projects,
'  you may want to replace (2 ^ Bit) references with calls to the
'  TwoToThe() function, as VB's compiler doesn't optimize
'  exponentiation as well as it could.
' *********************************************************************
Option Explicit

' *********************************************************************
'  Public Functions: Bits (within Bytes)
' *********************************************************************
Public Function BitSetB(ByVal ByteIn As Byte, ByVal Bit As Integer) As Byte
   If Bit >= 0 And Bit <= 7 Then
      ' Set the Nth Bit to 1.
      BitSetB = (ByteIn Or 2 ^ Bit)
   Else
      ' Could raise an error, if more appropriate?
      BitSetB = ByteIn
   End If
End Function

Public Function BitClearB(ByVal ByteIn As Byte, ByVal Bit As Integer) As Byte
   If Bit >= 0 And Bit <= 7 Then
      ' Clear the Nth Bit to 0.
      BitClearB = (ByteIn And Not 2 ^ Bit)
   Else
      ' Could raise an error, if more appropriate?
      BitClearB = ByteIn
   End If
End Function

Public Function BitToggleB(ByVal ByteIn As Byte, ByVal Bit As Integer) As Byte
   If Bit >= 0 And Bit <= 7 Then
      ' Return Nth power bit as true/false
      BitToggleB = (ByteIn Xor (2 ^ Bit))
   Else
      ' Could raise an error, if more appropriate?
      BitToggleB = ByteIn
   End If
End Function

Public Function BitValueB(ByVal ByteIn As Byte, ByVal Bit As Integer) As Boolean
   If Bit >= 0 And Bit <= 7 Then
      ' Return Nth power bit as true/false
      BitValueB = ((ByteIn And (2 ^ Bit)) > 0)
   Else
      ' Could raise an error, if more appropriate?
      BitValueB = False
   End If
End Function

' *********************************************************************
'  Public Functions: Bits (within Words)
' *********************************************************************
Public Function BitClearI(ByVal WordIn As Integer, ByVal Bit As Integer) As Integer
   ' Prevent overflow by using an integer constant for (2^15).
   If Bit < 15 Then
      BitClearI = WordIn And Not (2 ^ Bit)
   ElseIf Bit = 15 Then
      BitClearI = WordIn And Not &H8000
   Else
      ' Could raise an error, if more appropriate?
      BitClearI = WordIn
   End If
End Function

Public Function BitSetI(ByVal WordIn As Integer, ByVal Bit As Integer) As Integer
   ' Prevent overflow by using an integer constant for (2^15).
   If Bit < 15 Then
      BitSetI = WordIn Or (2 ^ Bit)
   ElseIf Bit = 15 Then
      BitSetI = WordIn Or &H8000
   Else
      ' Could raise an error, if more appropriate?
      BitSetI = WordIn
   End If
End Function

Public Function BitToggleI(ByVal WordIn As Integer, ByVal Bit As Integer) As Integer
   ' Prevent overflow by using an integer constant for (2^15).
   If Bit < 15 Then
      BitToggleI = WordIn Xor (2 ^ Bit)
   ElseIf Bit = 15 Then
      BitToggleI = WordIn Xor &H8000
   Else
      ' Could raise an error, if more appropriate?
      BitToggleI = WordIn
   End If
End Function

Public Function BitValueI(ByVal WordIn As Integer, ByVal Bit As Integer) As Boolean
   If Bit >= 0 And Bit <= 15 Then
      ' Return Nth power bit as true/false
      BitValueI = ((WordIn And (2 ^ Bit)) > 0)
   Else
      ' Could raise an error, if more appropriate?
      BitValueI = False
   End If
End Function

' *********************************************************************
'  Public Functions: Bytes
' *********************************************************************
Public Function ByteShiftL(ByVal InVal As Byte, Optional ByVal Bits As Long = 1) As Byte
   ' Shifting more than 7 bits would
   ' clear entire value.
   If Bits >= 0 And Bits <= 7 Then
      ' Shift left requested number of bits.
      ' Intermediate result can overflow a
      ' byte, so need to strip the high byte
      ' of initial multiplication.
      ByteShiftL = (InVal * (2 ^ Bits)) And &HFF
   End If
End Function

Public Function ByteShiftR(ByVal InVal As Byte, Optional ByVal Bits As Long = 1) As Byte
   ' Shifting more than 7 bits would
   ' clear entire value.
   If Bits >= 0 And Bits <= 7 Then
      ' Shift right requested number of bits.
      ByteShiftR = InVal \ (2 ^ Bits)
   End If
End Function

Public Function ByteRotateL(ByVal InVal As Byte, Optional ByVal Bits As Long = 1) As Byte
   ' Leverage power in other rotate routine.
   ByteRotateL = ByteRotateR(InVal, 8 - (Bits Mod 8))
End Function

Public Function ByteRotateR(ByVal InVal As Byte, Optional ByVal Bits As Long = 1) As Byte
   Dim nRet As Byte
   Dim nMask As Byte
   
   ' Might as well allow going in circles
   ' as many times as desired, eh?
   Bits = Bits Mod 8
   
   If Bits > 0 Then
      ' Shift right requested number of bits.
      nRet = (InVal \ 2 ^ (Bits)) And &HFF
      ' Promote low N bits to high N bits.
      nMask = (InVal * 2 ^ (8 - Bits)) And &HFF
      ' Combine with original shifted bits.
      ByteRotateR = nRet Or nMask
   Else
      ByteRotateR = InVal
   End If
End Function

' *********************************************************************
'  Public Functions: Words
' *********************************************************************
Public Function ByteHi(ByVal WordIn As Integer) As Byte
   ' Lop off low byte with divide. If less than
   ' zero, then account for sign bit (adding &h10000
   ' implicitly converts to Long before divide).
   If WordIn < 0 Then
      ByteHi = (WordIn + &H10000) \ &H100
   Else
      ByteHi = WordIn \ &H100
   End If
End Function

Public Function ByteLo(ByVal WordIn As Integer) As Byte
   ' Mask off high byte and return low.
   ByteLo = WordIn And &HFF
End Function

Public Function ByteSwap(ByVal WordIn As Integer) As Integer
   Dim ByteHi As Integer
   Dim ByteLo As Integer
   Dim NewHi As Long
   
   ' Separate bytes using same strategy as in
   ' ByteHi and ByteLo functions. Faster to do
   ' it inline than to make function calls.
   If WordIn < 0 Then
      ByteHi = (WordIn + &H10000) \ &H100
   Else
      ByteHi = WordIn \ &H100
   End If
   ByteLo = WordIn And &HFF
   
   ' Shift low byte left by 8
   NewHi = ByteLo * &H100&
   
   ' Account for sign-bit
   If NewHi > &H7FFF Then
      ByteLo = NewHi - &H10000
   Else
      ByteLo = NewHi
   End If
   
   ' Place high byte in low position
   ByteSwap = ByteLo Or ByteHi
End Function

Public Function WordShiftL(ByVal InVal As Integer, Optional ByVal Bits As Long = 1) As Integer
   ' Shifting more than 15 bits would
   ' clear entire value.
   If Bits >= 0 And Bits <= 15 Then
      ' Shift left requested number of bits.
      ' Intermediate result of multiplication
      ' can overflow, so just return the low
      ' word.
      WordShiftL = WordLo(InVal * (2 ^ Bits))
   End If
End Function

Public Function WordShiftR(ByVal InVal As Integer, Optional ByVal Bits As Long = 1) As Integer
   Dim nRet As Long
   ' Shifting more than 15 bits would
   ' clear entire value.
   If Bits >= 0 And Bits <= 15 Then
      ' Coercion avoids problems with sign
      ' bit. Clear high word of temp.
      nRet = CLng(InVal) And Not &HFFFF0000
      ' Shift right requested number of bits.
      ' Just return the low word.
      WordShiftR = WordLo(nRet \ (2 ^ Bits))
   End If
End Function

Public Function WordRotateL(ByVal InVal As Integer, Optional ByVal Bits As Long = 1) As Integer
   ' Leverage power in other rotate routine.
   WordRotateL = WordRotateR(InVal, 16 - (Bits Mod 16))
End Function

Public Function WordRotateR(ByVal InVal As Integer, Optional ByVal Bits As Long = 1) As Integer
   Dim nRet As Long
   Dim nMask As Long
   ' Might as well allow going in circles
   ' as many times as desired, eh?
   Bits = Bits Mod 16
   If Bits > 0 Then
      ' Shift right requested number of bits.
      nRet = WordShiftR(InVal, Bits)
      ' Tack rightmost bits onto left end by
      ' masking them off, shifting left, and
      ' recombining.
      nMask = (InVal And Not IntMaskHiN(16 - Bits))
      nMask = WordShiftL(nMask, 16 - Bits)
      WordRotateR = nRet Or nMask
   Else
      WordRotateR = InVal
   End If
End Function

' *********************************************************************
'  Public Functions: DWords
' *********************************************************************
Public Function DWordShiftL(ByVal InVal As Long, Optional ByVal Bits As Long = 1) As Long
   Dim nRet As Currency

   ' Shifting more than 31 bits would
   ' clear entire value.
   If Bits = 0 Then
      ' Return value as is.
      nRet = InVal

   ElseIf Bits >= 1 And Bits <= 31 Then
      ' Clear top N bits, as they'll disappear
      ' regardless.
      nRet = InVal And Not LongMaskHiN(Bits)

      ' Shift left requested number of bits.
      nRet = nRet * (2 ^ Bits)

      ' Account for sign-bit
      Const SignBitToggle = 4294967296# '(2 ^ 32)
      If nRet > &H7FFFFFFF Then
         nRet = nRet - SignBitToggle
      End If
   End If

   ' Return results
   DWordShiftL = nRet
End Function

Public Function DWordShiftR(ByVal InVal As Long, Optional ByVal Bits As Long = 1) As Long
   Dim nRet As Long
   Dim ResetShiftedSign As Boolean
   Const SignBit As Long = &H80000000
   
   ' Shifting more than 31 bits would
   ' clear entire value.
   If Bits = 0 Then
      ' Return value as is.
      nRet = InVal
      
   ElseIf Bits >= 1 And Bits <= 30 Then
      ' Clear sign bit
      If InVal And SignBit Then
         nRet = InVal And Not SignBit
         ResetShiftedSign = True
      Else
         nRet = InVal
      End If
      
      ' Shift right requested number of bits.
      nRet = nRet \ (2 ^ Bits)
      
      ' Reset sign bit in new position
      If ResetShiftedSign Then
         nRet = nRet Or (2 ^ (31 - Bits))
      End If
      
   ElseIf Bits = 31 Then
      ' Just turn on the sign bit if needed
      If InVal And SignBit Then
         nRet = 1
      End If
   End If

   ' Return results
   DWordShiftR = nRet
End Function
Public Function INT64ShiftR(ByVal InVal As Long, Optional ByVal Bits As Long = 1) As Currency
   Dim nRet As Currency
   Dim ResetShiftedSign As Boolean
   Const SignBit As Long = &H80000000
   
   ' Shifting more than 31 bits would
   ' clear entire value.
   If Bits = 0 Then
      ' Return value as is.
      nRet = InVal
      
   ElseIf Bits >= 1 And Bits <= 62 Then
      ' Clear sign bit
      If InVal And SignBit Then
         nRet = InVal And Not SignBit
         ResetShiftedSign = True
      Else
         nRet = InVal
      End If
      
      ' Shift right requested number of bits.
      nRet = nRet \ (2 ^ 1)
      
      ' Reset sign bit in new position
      If ResetShiftedSign Then
         nRet = nRet Or (2 ^ (63 - Bits))
      End If
      
   ElseIf Bits = 63 Then
      ' Just turn on the sign bit if needed
      If InVal And SignBit Then
         nRet = 1
      End If
   End If

   ' Return results
  INT64ShiftR = nRet
End Function
Public Function DWordRotateL(ByVal InVal As Long, Optional ByVal Bits As Long = 1) As Long
   ' Leverage power in other rotate routine.
   DWordRotateL = DWordRotateR(InVal, 32 - (Bits Mod 32))
End Function

Public Function DWordRotateR(ByVal InVal As Long, Optional ByVal Bits As Long = 1) As Long
   Dim nRet As Long
   Dim nMask As Long
   
   ' Might as well allow going in circles
   ' as many times as desired, eh?
   Bits = Bits Mod 32
   nRet = InVal
   
   If Bits > 0 Then
      ' Shift right requested number of bits.
      nRet = DWordShiftR(InVal, Bits)
      
      ' Tack rightmost bits onto left end by
      ' masking them off, shifting left, and
      ' recombining.
      nMask = (InVal And Not LongMaskHiN(32 - Bits))
      nMask = DWordShiftL(nMask, 32 - Bits)
      nRet = nRet Or nMask
   End If
      
   ' Return results
   DWordRotateR = nRet
End Function

Public Function WordHi(ByVal LongIn As Long) As Integer
   ' Mask off low word then do integer divide to
   ' shift right by 16.
   WordHi = (LongIn And &HFFFF0000) \ &H10000
End Function

Public Function WordLo(ByVal LongIn As Long) As Integer
   ' Low word retrieved by masking off high word.
   ' If low word is too large, twiddle sign bit.
   If (LongIn And &HFFFF&) > &H7FFF Then
      WordLo = (LongIn And &HFFFF&) - &H10000
   Else
      WordLo = LongIn And &HFFFF&
   End If
End Function

Public Function WordSwap(ByVal LongIn As Long) As Long
   Dim WordLo As Variant
   Dim WordHi As Integer
   
   ' Use the same logic as the WordLo and WordHi functions.
   If (LongIn And &HFFFF&) > &H7FFF Then
      WordLo = (LongIn And &HFFFF&) - &H10000
   Else
      WordLo = LongIn And &HFFFF&
   End If
   WordHi = (LongIn And &HFFFF0000) \ &H10000
   
   ' Ditto the MakeLong function. Note that WordLo needs
   ' to be a Variant to avoid overflow when it's equal
   ' to &h8000.
   WordSwap = (WordLo * &H10000) Or (WordHi And &HFFFF&)
End Function

' *********************************************************************
'  Public Functions: Combinatorial
' *********************************************************************
Public Function MakeInt64(ByVal LongHi As Long, ByVal LongLo As Long) As Double
   ' High word is coerced to Currency to allow it to
   ' overflow limits of multiplication which shifts
   ' it left.
   MakeInt64 = (CDbl(LongHi) * &H100000) Or (LongLo And &HFFFFF)
End Function

Public Function MakeLong(ByVal WordHi As Integer, ByVal WordLo As Integer) As Long
   ' High word is coerced to Currency to allow it to
   ' overflow limits of multiplication which shifts
   ' it left.
   MakeLong = (CLng(WordHi) * &H10000) Or (WordLo And &HFFFF&)
End Function

Public Function MakeWord(ByVal ByteHi As Byte, ByVal ByteLo As Byte) As Integer
   ' If the high byte would push the final result out of the
   ' signed integer range, it must be slid back.
   If ByteHi > &H7F Then
      MakeWord = ((ByteHi * &H100&) Or ByteLo) - &H10000
   Else
      MakeWord = (ByteHi * &H100&) Or ByteLo
   End If
End Function

' *********************************************************************
'  Public Functions: Formatting
' *********************************************************************
Public Function FmtBin(ByVal InVal As Variant) As String
   Dim tmpRet As String
   Dim i As Integer
   Dim pos As Integer
   Dim length As Integer
   
   ' Determine proper output length, based on vartype
   Select Case VarType(InVal)
      Case vbByte
         length = 8
      Case vbInteger
         length = 16
      Case vbLong
         ' Function only designed to handle Integers, as
         ' (InVal And (2 ^ i)) will overflow when a Long
         ' variable's sign bit is set.  Need to recurse.
         FmtBin = FmtBin(WordHi(InVal)) & " " & FmtBin(WordLo(InVal))
         Exit Function
      Case Else
         ' Function not designed for anything else
         Debug.Print TypeName(InVal)
         err.Raise 13, "HiLo.FmtBin", "Type mismatch"
         Exit Function
   End Select
   
   ' Pad output string with all 0's
   tmpRet = String$(length, "0")
   
   ' Cycle through each position inserting a 1 in
   ' return string if that bit is set.
   For i = (length - 1) To 0 Step -1
      pos = pos + 1
      If InVal And (2 ^ i) Then
         Mid$(tmpRet, pos, 1) = "1"
      End If
   Next i
   
   ' Add a space between bytes
   Select Case length
      Case 8
         FmtBin = tmpRet
      Case 16
         FmtBin = Left$(tmpRet, 8) & " " & Right$(tmpRet, 8)
   End Select
End Function

Public Function FmtHex(ByVal InVal As Long, ByVal OutLen As Integer) As String
   ' Left pad with zeros to OutLen.
   FmtHex = Right$(String$(OutLen, "0") & Hex$(InVal), OutLen)
End Function

' *********************************************************************
'  Private Functions: Lookup tables
' *********************************************************************
Private Function TwoToThe(ByVal Power As Long) As Long
   Static BeenHere As Boolean
   Static Results(0 To 31) As Long
   Dim i As Long
   
   ' Build lookup table, first time through.
   ' Results hold powers of two from 0-31.
   If Not BeenHere Then
      For i = 0 To 30
         Results(i) = 2 ^ i
      Next i
      Results(31) = &H80000000
      BeenHere = True
   End If
   
   ' Return requested result
   If Power >= 0 And Power <= 31 Then
      TwoToThe = Results(Power)
   End If
End Function

Private Function LongMaskHiN(ByVal Bits As Long) As Long
   Static BeenHere As Boolean
   Static Masks(1 To 32) As Long
   Dim i As Long
   
   ' Build lookup table, first time through.
   ' Masks fill high N bits with 1.
   If Not BeenHere Then
      Masks(32) = &HFFFFFFFF
      Masks(1) = &H80000000
      For i = 31 To 2 Step -1
         Masks(i) = Masks(i + 1) - TwoToThe(31 - i)
      Next i
      BeenHere = True
   End If
   
   ' Return requested mask
   If Bits >= 1 And Bits <= 32 Then
      LongMaskHiN = Masks(Bits)
   End If
End Function

Private Function IntMaskHiN(ByVal Bits As Long) As Integer
   Static BeenHere As Boolean
   Static Masks(1 To 16) As Integer
   Dim i As Long
   
   ' Build lookup table, first time through.
   ' Masks fill high N bits with 1.
   If Not BeenHere Then
      Masks(16) = &HFFFF
      Masks(1) = &H8000
      For i = 15 To 2 Step -1
         Masks(i) = Masks(i + 1) - TwoToThe(15 - i)
      Next i
      BeenHere = True
   End If
   
   ' Return requested mask
   If Bits >= 1 And Bits <= 16 Then
      IntMaskHiN = Masks(Bits)
   End If
End Function




