Attribute VB_Name = "libdata"
Option Explicit

Public Feriados As Object
Private Const ABA_FERIADO As String = "Feriado"

Public Function quantidadeDCEntre(ByVal DataInicial As Date, ByVal DataFinal As Date, _
    Optional ByVal excluiInicial As Boolean = False, Optional ByVal excluiFinal As Boolean = True) As Long
    
    DataFinal = DateAdd("d", 1, DataFinal)

    If excluiInicial = True Then
        DataInicial = DateAdd("d", 1, DataInicial)
    End If
    If excluiFinal = True Then
        DataFinal = DateAdd("d", -1, DataFinal)
    End If
    
    Dim Quantidade As Long
    Quantidade = DateDiff("d", DataInicial, DataFinal)

    quantidadeDCEntre = Quantidade

End Function


Public Function quantidadeDUEntre(ByVal DataInicial As Date, ByVal DataFinal As Date, _
    Optional ByVal excluiInicial As Boolean = False, Optional ByVal excluiFinal As Boolean = True) As Long

    If DataFinal < DataInicial Then
        Dim TempData As Date
        TempData = DataInicial
        DataInicial = DataFinal
        DataFinal = TempData
    End If
    
    Dim Data As Date
    If excluiInicial = True Then
        DataInicial = DateAdd("d", 1, DataInicial)
    End If
    If excluiFinal = True Then
        DataFinal = DateAdd("d", -1, DataFinal)
    End If
    
    Dim Quantidade As Long
    Quantidade = 0
    For Data = DataInicial To DataFinal
        If diaUtil(Data) Then
            Quantidade = Quantidade + 1
        End If
    Next Data

    quantidadeDUEntre = Quantidade

End Function

Public Function diaUtilPosterior(ByVal dateVal As Date, Optional ByVal incluirData As Boolean = False) As Date
    
    If Feriados Is Nothing Then Set Feriados = pegarFeriados
    
    If incluirData = False Then
        dateVal = DateAdd("d", 1, dateVal)
    End If
    
    Do While diaUtil(dateVal) = False
        dateVal = DateAdd("d", 1, dateVal)
    Loop
    
    diaUtilPosterior = dateVal

End Function

Public Function diaUtilAnterior(ByVal dateVal As Date, Optional ByVal incluirData As Boolean = False) As Date

    If Feriados Is Nothing Then Set Feriados = pegarFeriados
    
    If incluirData = False Then
        dateVal = DateAdd("d", -1, dateVal)
    End If
    
    Do While diaUtil(dateVal) = False
        dateVal = DateAdd("d", -1, dateVal)
    Loop
    
    diaUtilAnterior = dateVal

End Function

Public Function diaUtil(ByVal dateVal As Date) As Boolean

    If Feriados Is Nothing Then Set Feriados = pegarFeriados
    
    If Weekday(dateVal, vbMonday) > 5 Or Feriados.Exists(dateVal) Then
        diaUtil = False
    Else
        diaUtil = True
    End If

End Function

Public Function pegarFeriados() As Object

    Dim i As Long
    i = 2
    Set pegarFeriados = CreateObject("Scripting.Dictionary")
    Do While Not IsEmpty(ThisWorkbook.Sheets(ABA_FERIADO).Cells(i, 1))
            pegarFeriados.add ThisWorkbook.Sheets(ABA_FERIADO).Cells(i, 1).value, ThisWorkbook.Sheets(ABA_FERIADO).Cells(i, 1).value
        i = i + 1
    Loop

End Function
