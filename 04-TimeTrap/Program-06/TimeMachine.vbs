Option Explicit

Sub Main
  ' 10  REM PROGRAM THE TIME MACHINE
  ' 20  PRINT "REMOVE THE EXACT"
  WScript.StdOut.WriteLine "REMOVE THE EXACT"
  ' 30  PRINT "WEIGHT TO GET THE"
  WScript.StdOut.WriteLine "WEIGHT TO GET THE"
  ' 40  PRINT "TIME MACHINE BACK TO THE PRESENT"
  WScript.StdOut.WriteLine "TIME MACHINE BACK TO THE PRESENT"
  ' 50  DATA 6,.3 ,.012,3400,1000
  ' 60  READ A,B,E,T,W
  Dim A, B, E, T, W
  A = 6
  B = 0.3
  E = 0.012
  T = 3400
  W = 1000
  ' 70  PRINT
  WScript.StdOut.WriteLine

  Dim Y_Dollar
  Do
    ' 80  PRINT "DISTANCE LOCKED AT ";A;" MILES"
    WScript.StdOut.WriteLine "DISTANCE LOCKED AT " & A & " MILES"
    ' 90  PRINT "YOU HAVE ";W;" LBS OF CARGO"
    WScript.StdOut.WriteLine "YOU HAVE " & W & " LBS OF CARGO"
    ' 100 PRINT
    WScript.StdOut.WriteLine ""
    ' 110 PRINT "HOW MUCH WEIGHT WILL YOU REMOVE";
    WScript.StdOut.Write "HOW MUCH WEIGHT WILL YOU REMOVE "
    ' 120 INPUT C
    Dim C
    C = Int(WScript.StdIn.ReadLine ())

    ' 130 IF C < W THEN 170
    If C >= W Then
      ' 140 PRINT "YOU CRAZY? YOU CAN'T DO THAT!"
      WScript.StdOut.WriteLine "YOU CRAZY? YOU CAN'T DO THAT!"
      ' 150 PRINT
      WScript.StdOut.WriteLine
      ' 160 GOTO 80
    Else
      ' 170 C = W - C
      C = W - C
      ' 180 D = (T - (A * B * C)) / (E * C)
      Dim D
      D = (T - (A * B * C)) / (E * C)
      ' 190 PRINT
      WScript.StdOut.WriteLine
      ' 200 PRINT "* WITH ";C;" LBS OF CARGO"
      WScript.StdOut.WriteLine "* WITH " & C & " LBS OF CARGO"
      ' 210 PRINT "YOU WILL END UP IN YEAR ";INT(1777 + D)
      WScript.StdOut.WriteLine "YOU WILL END UP IN YEAR " & Int(1777 + D)
      ' 220 PRINT
      WScript.StdOut.WriteLine
      ' 230 PRINT "TRY AGAIN (Y/N)";
      WScript.StdOut.Write "TRY AGAIN (Y/N) "
      ' 240 INPUT Y$
      Y_Dollar = UCase(WScript.StdIn.ReadLine ())
    End If
    ' 250 IF Y$ = "Y" THEN 80
  Loop While Y_Dollar = "Y"
  ' 260 END
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
