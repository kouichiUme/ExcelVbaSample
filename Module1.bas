Attribute VB_Name = "Module1"
Option Explicit

Sub sample001()

    'PêZðwè
    Range("A1").Value = "ZA1"
    'ZÍÍðwè
    Range("A3:A5").Value = "ZA3:A5"

    'ZC1ÆZD10ðÍÞZÍÍiZC1:D10jðwè
    Range(Range("C1"), Range("D10")).Value = "ÍÞÍÍ"

    'PsÚEUñÚÌZðwè
    Cells(1, 6).Value = "PsEUñ"
    'RsÚEFñÚÌZðwè
    Cells(3, "F").Value = "RsEFñ"
    Cells(3, 6).Value = "RsEUñ"


End Sub
