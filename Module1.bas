Attribute VB_Name = "Module1"
Option Explicit

Sub sample001()

    '�P��Z�����w��
    Range("A1").Value = "�Z��A1"
    '�Z���͈͂��w��
    Range("A3:A5").Value = "�Z��A3:A5"

    '�Z��C1�ƃZ��D10���͂ރZ���͈́i�Z��C1:D10�j���w��
    Range(Range("C1"), Range("D10")).Value = "�͂ޔ͈�"

    '�P�s�ځE�U��ڂ̃Z�����w��
    Cells(1, 6).Value = "�P�s�E�U��"
    '�R�s�ځEF��ڂ̃Z�����w��
    Cells(3, "F").Value = "�R�s�EF��"
    Cells(3, 6).Value = "�R�s�E�U��"


End Sub
