Option Explicit

Sub PromptForSecretNumber
    WScript.StdOut.Write "What is the secret input number? "
End Sub

Sub PromptForSecretMessage
    WScript.StdOut.Write "Secret message-> "
End Sub

Function IsTerminationKeyword(text)
    IsTerminationKeyword = (text = "STOP")
End Function

Function IsValidCharacterCode(charCode)
    IsValidCharacterCode = ((charCode >= Asc("A")) And (charCode <= Asc("Z")))
End Function

Sub Main
    ' Decoder
    PromptForSecretNumber()

    Dim SecretNumber
    SecretNumber = WScript.StdIn.ReadLine()
    WScript.StdOut.WriteLine("SecretNumber : " & SecretNumber)

    WScript.StdOut.WriteLine "Type 'STOP' when asked for the secret message to end the program"

    Dim SecretMessage

    PromptForSecretMessage()
    SecretMessage = WScript.StdIn.ReadLine()

    Do Until (IsTerminationKeyword(SecretMessage))
        WScript.StdOut.WriteLine(DecodeMessage(SecretNumber, SecretMessage))

        PromptForSecretMessage()
        SecretMessage = WScript.StdIn.ReadLine()
    Loop

    WScript.StdOut.WriteLine("***END OF DECODING***")
End Sub

Function DecodeMessage(SecretNumber, SecretMessage)
    Dim DecodedMessage
    DecodedMessage = ""

    Dim i, EncodedCharacterCode
    For i = 1 To Len(SecretMessage)
        EncodedCharacterCode = Asc(Mid(SecretMessage, i, 1))

        Dim DecodedCharacter
        If IsValidCharacterCode(EncodedCharacterCode) Then
            EncodedCharacterCode = Asc(Mid(SecretMessage, i, 1)) - Asc("A") + 1
            DecodedCharacter = _
                Chr((EncodedCharacterCode + SecretNumber) _
                    - Int((EncodedCharacterCode + SecretNumber) / 26) _
                    * 26 _
                    + Asc("A"))
        Else
            DecodedCharacter = Chr(EncodedCharacterCode)
        End If

        DecodedMessage = DecodedMessage & DecodedCharacter
    Next

    DecodeMessage = DecodedMessage
End Function

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

Sub DecodeMessage_GivenSecretCodeAndMessage_CorrectlyDecodesMessage
    Dim SecretNumber, SecretMessage, Actual, Expected
    SecretNumber = 17
    SecretMessage = "KWLM ZML. VIBQWVIT MUMZOMVKG. JZCBM AKQMVBQAB WV BPM TWWAM. EQBP BQUM UIKPQVM IVL VCKTMIZ LMDQKM. PQA WJRMKBQDM QA BW KPIVOM BPM WCBKWUM WN BPM IUMZQKIV ZMDWTCBQWV."
    Expected = "CODE RED. NATIONAL EMERGENCY. BRUTE SCIENTIST ON THE LOOSE. WITH TIME MACHINE AND NUCLEAR DEVICE. HIS OBJECTIVE IS TO CHANGE THE OUTCOME OF THE AMERICAN REVOLUTION."

    Actual = DecodeMessage(SecretNumber, SecretMessage)

    AssertEqual Expected, Actual
End Sub

Sub IsValidCharacterCode_GivenACharacterBetweenCapitalAAndCapitalZ_ReturnsTrue
    Assert IsValidCharacterCode(Asc("A")), "'A' is not a valid character?"
    Assert IsValidCharacterCode(Asc("B")), "'B' is not a valid character?"
    Assert IsValidCharacterCode(Asc("C")), "'C' is not a valid character?"
    Assert IsValidCharacterCode(Asc("D")), "'D' is not a valid character?"
    Assert IsValidCharacterCode(Asc("E")), "'E' is not a valid character?"
    Assert IsValidCharacterCode(Asc("F")), "'F' is not a valid character?"
    Assert IsValidCharacterCode(Asc("G")), "'G' is not a valid character?"
    Assert IsValidCharacterCode(Asc("H")), "'H' is not a valid character?"
    Assert IsValidCharacterCode(Asc("I")), "'I' is not a valid character?"
    Assert IsValidCharacterCode(Asc("J")), "'J' is not a valid character?"
    Assert IsValidCharacterCode(Asc("K")), "'K' is not a valid character?"
    Assert IsValidCharacterCode(Asc("L")), "'L' is not a valid character?"
    Assert IsValidCharacterCode(Asc("M")), "'M' is not a valid character?"
    Assert IsValidCharacterCode(Asc("N")), "'N' is not a valid character?"
    Assert IsValidCharacterCode(Asc("O")), "'O' is not a valid character?"
    Assert IsValidCharacterCode(Asc("P")), "'P' is not a valid character?"
    Assert IsValidCharacterCode(Asc("Q")), "'Q' is not a valid character?"
    Assert IsValidCharacterCode(Asc("R")), "'R' is not a valid character?"
    Assert IsValidCharacterCode(Asc("S")), "'S' is not a valid character?"
    Assert IsValidCharacterCode(Asc("T")), "'T' is not a valid character?"
    Assert IsValidCharacterCode(Asc("U")), "'U' is not a valid character?"
    Assert IsValidCharacterCode(Asc("V")), "'V' is not a valid character?"
    Assert IsValidCharacterCode(Asc("W")), "'W' is not a valid character?"
    Assert IsValidCharacterCode(Asc("X")), "'X' is not a valid character?"
    Assert IsValidCharacterCode(Asc("Y")), "'Y' is not a valid character?"
    Assert IsValidCharacterCode(Asc("Z")), "'Z' is not a valid character?"

    Assert Not IsValidCharacterCode(Asc("a")), "'a' is a valid character?"
    Assert Not IsValidCharacterCode(Asc("z")), "'z' is a valid character?"

End Sub

Sub RunTests
    IsValidCharacterCode_GivenACharacterBetweenCapitalAAndCapitalZ_ReturnsTrue()
    DecodeMessage_GivenSecretCodeAndMessage_CorrectlyDecodesMessage()

    WScript.Echo "ALL TESTS PASSED!"
End Sub

If WScript.Arguments.Count = 0 Then
    Main()
ElseIf LCase(WScript.Arguments(0)) = "test" Then
    RunTests()
End If
