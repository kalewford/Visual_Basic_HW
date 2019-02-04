Sub WallStreetMultipleYearStock()



Dim LastRowNum, FirstTickerRowNum, NextTickerRowNum As Long



Dim TickerYearClose, TickerYearOpen As Double



Dim TickerNum As Long



Dim TickerName, NextTickerName  As String



Dim Volume As Double



Dim TotalTabNum As Integer



TotalTabNum = ThisWorkbook.Sheets.Count



For j = 1 To TotalTabNum

  Sheets(j).Activate





  TickerNum = 1

  FirstTickerRowNum = 2



  Range("M1").Value = "Ticker"

  Range("N1").Value = "Yearly Change"

  Range("O1").Value = "Percent Change"

  Range("P1").Value = "Total Stock Volume"

  

  LastRowNum = Cells(Rows.Count, 1).End(xlUp).Row

  

  Volume = Cells(2, 7).Value

  For i = 2 To LastRowNum

    TickerName = Cells(i, 1).Value

    NextTickerName = Cells(i + 1, 1).Value

    If TickerName = NextTickerName Then

      Volume = Volume + Cells(i + 1, 7).Value

    Else

       TickerNum = TickerNum + 1

       NextTickerRowNum = i + 1

       Cells(TickerNum, 13).Value = TickerName

       Cells(TickerNum, 16).Value = Volume

       TickerYearClose = Cells(NextTickerRowNum - 1, 6).Value

       TickerYearOpen = Cells(FirstTickerRowNum, 3).Value

      

       While TickerYearOpen = 0

             FirstTickerRowNum = FirstTickerRowNum + 1

             TickerYearOpen = Cells(FirstTickerRowNum, 3).Value

       Wend



       Cells(TickerNum, 14).Value = TickerYearClose - TickerYearOpen



       

       

       Cells(TickerNum, 15).Value = Cells(TickerNum, 14).Value / TickerYearOpen

       

       Cells(TickerNum, 15).NumberFormat = "0.00%"



       If Cells(TickerNum, 14).Value > 0 Then

            Cells(TickerNum, 14).Interior.ColorIndex = 4

        Else

            Cells(TickerNum, 14).Interior.ColorIndex = 3

        End If

     

      FirstTickerRowNum = NextTickerRowNum

      Volume = Cells(FirstTickerRowNum, 7).Value

      

    End If

 Next i

 Range("M1:P1").Columns.AutoFit

Next j



End Sub



Sub QuickSort(vArray As Variant, arrLbound As Double, arrUbound As Double)

'Smallest to largest



Dim pivotVal As Variant

Dim vSwap    As Variant

Dim tmpLow   As Double

Dim tmpHi    As Double



 

tmpLow = arrLbound

tmpHi = arrUbound

pivotVal = vArray((arrLbound + arrUbound) \ 2)

 

While (tmpLow <= tmpHi) 'divide

   While (vArray(tmpLow) < pivotVal And tmpLow < arrUbound)

      tmpLow = tmpLow + 1

   Wend

  

   While (pivotVal < vArray(tmpHi) And tmpHi > arrLbound)

      tmpHi = tmpHi - 1

   Wend

 

   If (tmpLow <= tmpHi) Then

      vSwap = vArray(tmpLow)

      vArray(tmpLow) = vArray(tmpHi)

      vArray(tmpHi) = vSwap

      tmpLow = tmpLow + 1

      tmpHi = tmpHi - 1

   End If

Wend

 

  If (arrLbound < tmpHi) Then QuickSort vArray, arrLbound, tmpHi

  If (tmpLow < arrUbound) Then QuickSort vArray, tmpLow, arrUbound

  

  

End Sub



Sub SortingStockData()



Dim TotalTabNum As Integer



Dim LastRowNum As Long



Dim PercentChangeData() As Variant

Dim VolumeData() As Variant



TotalTabNum = ThisWorkbook.Sheets.Count



For j = 1 To TotalTabNum

  Sheets(j).Activate

  Range("S1").Value = "Ticker"

  Range("T1").Value = "Value"

  Range("R2").Value = "Greatest % Increase"

  Range("R3").Value = "Greatest % Decrease"

  Range("R4").Value = "Greatest Total Volume"

  

  Range("R:T").Columns.AutoFit



  'Return Row Number

  LastRowNum = Cells(Rows.Count, 15).End(xlUp).Row



  ReDim PercentChangeData(1 To LastRowNum - 1)

  ReDim VolumeData(1 To LastRowNum - 1)



  For i = 2 To LastRowNum

    PercentChangeData(i - 1) = Cells(i, 15).Value

    VolumeData(i - 1) = Cells(i, 16).Value

  Next i



  Call QuickSort(PercentChangeData(), LBound(PercentChangeData), UBound(PercentChangeData))

  Call QuickSort(VolumeData(), LBound(VolumeData), UBound(VolumeData))



  

  For i = 2 To LastRowNum

    If Cells(i, 15).Value = PercentChangeData(1) Then

        Range("S3").Value = Cells(i, 13).Value

        Range("T3").Value = Cells(i, 15).Value

        Range("T3").NumberFormat = "0.00%"

    End If



    If Cells(i, 15).Value = PercentChangeData(LastRowNum - 1) Then

        Range("S2").Value = Cells(i, 13).Value

        Range("T2").Value = Cells(i, 15).Value

        Range("T2").NumberFormat = "0.00%"

    End If



    If Cells(i, 16).Value = VolumeData(LastRowNum - 1) Then

        Range("S4").Value = Cells(i, 13).Value

        Range("T4").Value = Cells(i, 16).Value

        Range("T4").NumberFormat = "0"

    End If

  Next i

  

 Next j

 

End Sub







