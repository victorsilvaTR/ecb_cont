Attribute VB_Name = "modRedondeo"
''   AsymDown      Asymmetrically rounds numbers down - similar to Int().
''                 Negative numbers get more negative.
''
''   SymDown       Symmetrically rounds numbers down - similar to Fix().
''                 Truncates all numbers toward 0.
''                 Same as AsymDown for positive numbers.
''
''   AsymUp        Asymmetrically rounds numbers fractions up.
''                 Same as SymDown for negative numbers.
''                 Similar to Ceiling.
''
''   SymUp         Symmetrically rounds fractions up - that is, away from 0.
''                 Same as AsymUp for positive numbers.
''                 Same as AsymDown for negative numbers.
''
''   AsymArith     Asymmetric arithmetic rounding - rounds .5 up always.
''                 Similar to Java worksheet Round function.
''
''   SymArith      Symmetric arithmetic rounding - rounds .5 away from 0.
''                 Same as AsymArith for positive numbers.
''                 Similar to Excel Worksheet Round function.
''
''   BRound Banker       's rounding.
''                 Rounds .5 up or down to achieve an even number.
''                 Symmetrical by definition.
''
''   RandRound     Random rounding.
''                 Rounds .5 up or down in a random fashion.
''
''   AltRound      Alternating rounding.
''                 Alternates between rounding .5 up or down.
''
''   ATruncDigits  Same as AsyncTrunc but takes different arguments.
                   
   Function AsymDown(ByVal X As Double, _
            Optional ByVal Factor As Double = 1) As Double
     AsymDown = Int(X * Factor) / Factor
   End Function

   Function SymDown(ByVal X As Double, _
            Optional ByVal Factor As Double = 1) As Double
     SymDown = Fix(X * Factor) / Factor
   '  Alternately:
   '  SymDown = AsymDown(Abs(X), Factor) * Sgn(X)
   End Function

   Function AsymUp(ByVal X As Double, _
            Optional ByVal Factor As Double = 1) As Double
   Dim Temp As Double
     Temp = Int(X * Factor)
     AsymUp = (Temp + IIf(X = Temp, 0, 1)) / Factor
   End Function

   Function SymUp(ByVal X As Double, _
            Optional ByVal Factor As Double = 1) As Double
   Dim Temp As Double
     Temp = Fix(X * Factor)
     SymUp = (Temp + IIf(X = Temp, 0, Sgn(X))) / Factor
   End Function

   Function AsymArith(ByVal X As Double, _
            Optional ByVal Factor As Double = 1) As Double
     AsymArith = Int(X * Factor + 0.5) / Factor
   End Function

   Function SymArith(ByVal X As Double, _
            Optional ByVal Factor As Double = 1) As Double
     SymArith = Fix(X * Factor + 0.5 * Sgn(X)) / Factor
   '  Alternately:
   '  SymArith = Abs(AsymArith(X, Factor)) * Sgn(X)
   End Function

   Function BRound(ByVal X As Double, _
            Optional ByVal Factor As Double = 1) As Double
   '  For smaller numbers:
   '  BRound = CLng(X * Factor) / Factor
   Dim Temp As Double, FixTemp As Double
     Temp = X * Factor
     FixTemp = Fix(Temp + 0.5 * Sgn(X))
     ' Handle rounding of .5 in a special manner
     If Temp - Int(Temp) = 0.5 Then
       If FixTemp / 2 <> Int(FixTemp / 2) Then ' Is Temp odd
         ' Reduce Magnitude by 1 to make even
         FixTemp = FixTemp - Sgn(X)
       End If
     End If
     BRound = FixTemp / Factor
   End Function

   Function RandRound(ByVal X As Double, _
            Optional ByVal Factor As Double = 1) As Double
   ' Should Execute Randomize statement somewhere prior to calling.
   Dim Temp As Double, FixTemp As Double
     Temp = X * Factor
     FixTemp = Fix(Temp + 0.5 * Sgn(X))
     ' Handle rounding of .5 in a special manner.
     If Temp - Int(Temp) = 0.5 Then
       ' Reduce Magnitude by 1 in half the cases.
       FixTemp = FixTemp - Int(Rnd * 2) * Sgn(X)
     End If
     RandRound = FixTemp / Factor
   End Function

   Function AltRound(ByVal X As Double, _
            Optional ByVal Factor As Double = 1) As Double
   Static fReduce As Boolean
   Dim Temp As Double, FixTemp As Double
     Temp = X * Factor
     FixTemp = Fix(Temp + 0.5 * Sgn(X))
     ' Handle rounding of .5 in a special manner.
     If Temp - Int(Temp) = 0.5 Then
       ' Alternate between rounding .5 down (negative) and up (positive).
       If (fReduce And Sgn(X) = 1) Or (Not fReduce And Sgn(X) = -1) Then
       ' Or, replace the previous If statement with the following to
       ' alternate between rounding .5 to reduce magnitude and increase
       ' magnitude.
       ' If fReduce Then
         FixTemp = FixTemp - Sgn(X)
       End If
       fReduce = Not fReduce
     End If
     AltRound = FixTemp / Factor
   End Function

   Function ADownDigits(ByVal X As Double, _
            Optional ByVal Digits As Integer = 0) As Double
     ADownDigits = AsymDown(X, 10 ^ Digits)
   End Function

