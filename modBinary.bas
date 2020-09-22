Attribute VB_Name = "modBinary"
'Please refer to HELP (modHelp.bas) for documentation.
'-----------------------------------------------------


Public Function RandomBin(Length As Long) As String
'generate random binary data (length of binary data)

'E.G. RandomBin(100) would produce a random binary number
'of length 100

    Dim Position, StringLen As Long
    Dim rndString, Chars As String

    Chars = "10" 'either 1 or 0...that is binary, right?
    StringLen = 0

    Randomize

    Do Until StringLen = Length
        Position = Int((Len(Chars) * Rnd) + 1)
            rndString = rndString & Mid(Chars, Position, 1)
        StringLen = StringLen + 1
    Loop

    RandomBin = rndString

End Function
