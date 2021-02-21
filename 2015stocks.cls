VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub stocks()

    Dim Ticker As String
    Dim Vol As Double
    Dim Year_Open As Double
    Dim Year_Close As Double
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Summary_Table_Row As Integer
    Dim colorInd As Integer

    Cells(1, 9).Value = "Ticker"
    Cells(1, 10).Value = "Yearly_Change"
    Cells(1, 11).Value = "Percent_Change"
    Cells(1, 12).Value = "Total_Stock_Vol"
    
    
    Summary_Table_Row = 2
    Vol = 0
    
    For i = 2 To 760192
        
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
        
    
    Ticker = Cells(i, 1).Value
    Vol = Cells(i, 7).Value
    
    Year_Open = Cells(i, 3).Value
    Year_Close = Cells(i, 6).Value
    
    Yearly_Change = Year_Close - Year_Open
    Percent_Change = (Year_Close - Year_Open) / Year_Open
    
    
        Cells(Summary_Table_Row, 9).Value = Ticker
        Cells(Summary_Table_Row, 10).Value = Yearly_Change
        Cells(Summary_Table_Row, 11).Value = Percent_Change
        Cells(Summary_Table_Row, 12).Value = Vol
        Summary_Table_Row = Summary_Table_Row + 1
  
  
     End If
     
     Next i
     For j = 2 To 3005
    
        If Cells(j, 10) > 0 Then
                Cells(j, 10).Interior.Color = vbGreen
        Else
                Cells(j, 10).Interior.Color = vbRed
        End If
    Next j
   
   
      
End Sub

