Attribute VB_Name = "Module1"
Option Explicit

Sub sample001()

    '単一セルを指定
    Range("A1").Value = "セルA1"
    'セル範囲を指定
    Range("A3:A5").Value = "セルA3:A5"

    'セルC1とセルD10を囲むセル範囲（セルC1:D10）を指定
    Range(Range("C1"), Range("D10")).Value = "囲む範囲"

    '１行目・６列目のセルを指定
    Cells(1, 6).Value = "１行・６列"
    '３行目・F列目のセルを指定
    Cells(3, "F").Value = "３行・F列"
    Cells(3, 6).Value = "３行・６列"


End Sub
