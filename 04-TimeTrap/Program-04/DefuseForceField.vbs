Option Explicit

Sub Main
  ' 10  REM   DEFUSE THE FORCE FIELD
  ' 20  C = 0
  Dim C
  C = 0
  ' 30  S = 38
  Dim S
  S = 38

  Dim R
  R = 0

  Dim Exploded, Diffused
  Exploded = False
  Diffused = False
  ''' 190 loops back here
  Do While Not (Exploded Or Diffused)
    ' 40  X1 = INT(RND (1) * (S - 1)) + 1
    Dim X1
    X1 = Int(Rnd(1) * (S - 1)) + 1
    ' 50  X2 = 99
    Dim X2
    X2 = 99
    ' 60  PRINT
    WScript.StdOut.WriteLine()

    ''' 150 loops back here
    Do While True
      ' 70  GOSUB 250
      Sub250 X1, X2, S
      ' 80  PRINT "WHAT NUMBER(1..";S;")";
      WScript.StdOut.Write("WHAT NUMBER(1.." & S & ") ")
      ' 90  INPUT X2
      X2 = CInt(WScript.StdIn.ReadLine())
      ' 100 GOSUB 250
      Sub250 X1, X2, S
      ' 110 IF X1 = X2 THEN 160
      If X1 = X2 Then Exit Do
      ' 120 PRINT "WRONG"
      WScript.StdOut.WriteLine "WRONG"
      ' 130 W = W + 1
      Dim W
      W = W + 1
      ' 140 IF W = 5 THEN 200
      If W = 5 Then
        Exploded = True
        Exit Do
      End If
    ' 150 GOTO 70
    Loop
    ' 160 R = R + 1
    R = R + 1
    ' 170 PRINT R;" NUMBERS RIGHT!"
    WScript.StdOut.WriteLine(R & " NUMBERS RIGHT!")
    ' 180 IF R = 3 THEN 220
    If R = 3 Then
      Diffused = True
    End If
  ' 190 GOTO 40
  Loop

  If Exploded Then
    ' 200 PRINT "FORCE FIELD EXPLODES -- YOU DIE!"
    WScript.StdOut.WriteLine("FORCE FIELD EXPLODES -- YOU DIE!")
    ' 210 END
  Else
    ' 220 PRINT "***** FORCE FIELD"
    WScript.StdOut.WriteLine("***** FORCE FIELD")
    ' 230 PRINT "      DEACTIVATED"
    WScript.StdOut.WriteLine("      DEACTIVATED")
    ' 240 END
  End If
End Sub

Sub Sub250(X1, X2, S)
  ' 250 PRINT
  WScript.StdOut.WriteLine()
  ' 260 PRINT "!";
  WScript.StdOut.Write "!"
  ' 270 FOR I = 1 TO S
  Dim I
  For I = 1 To S
    ' 280 IF X1 = I THEN 320
    ' 290 IF X2 = I THEN 370
    ' 300 PRINT " ";
    ' 310 GOTO 380
    ' 320 IF X1 = X2 THEN 350
    ' 330 PRINT "*";
    ' 340 GOTO 380
    ' 350 PRINT "#";
    ' 360 GOTO 380
    ' 370 PRINT "^";

    If X1 = I Then
      If X1 = X2 Then
        WScript.StdOut.Write "#"
      Else
        WScript.StdOut.Write "*"
      End If
    ElseIf X2 = I Then
      WScript.StdOut.Write "^"
    Else
      WScript.StdOut.Write " "
    End If
  ' 380 NEXT I
  Next
  ' 390 PRINT "!"
  WScript.StdOut.WriteLine "!"
  ' 400 RETURN
End Sub

''' ASSERT FUNCTIONS -----------------------------------------------------------

Sub Assert(X, Msg)
  If Not X Then
    Err.Raise 1, Msg, Msg       
  End If
End Sub

Sub AssertEqual(Expected, Actual)
  Dim Message
  Message = "Values are not equal. EXPECTED: " & Expected & "; ACTUAL: " & Actual

  Assert (Expected = Actual), Message
End Sub

''' TEST FUNCTIONS -------------------------------------------------------------

Sub RunTests
    

  WScript.Echo "ALL TESTS PASSED!"
End Sub

''' ENTRY POINT ----------------------------------------------------------------

If WScript.Arguments.Count = 0 Then
  Main()
ElseIf LCase(WScript.Arguments(0)) = "test" Then
  RunTests()
End If
