# Lee-Carter Model based on MLE with different jump-off dates using VBA

Basic useful feature list:

 * Ctrl+S / Cmd+S to save the file
 * Ctrl+Shift+S / Cmd+Shift+S to choose to save as Markdown or HTML
 * Drag and drop a file into here to load it
 * File contents are saved in the URL so you can share files


Reference List:
Ronald Lee and Timothy Miller (2001), EVALUATING THE PERFORMANCE OF THE LEECARTER METHOD FOR FORECASTING MORTALITY, Demography

A.E. Renshaw, S. Haberman, A cohort-based extension to the Lee–Carter model for
mortality reduction factors, Insurance: Mathematics and Economics 38 (2006) 556–570

Ciolo Miguel C. Calma, Monica E. Revadulla (2014), FORECASTING PHILIPPINE
MORTALITY RATES USING THE LEE-CARTER MODEL, Academia, available at
<http://www.academia.edu/11982901/Forecasting_Mortality_Rates_using_the_Lee-
Carter_model> assessed at 1st October 2016


And here's my code! Any comments and modifications are preferred!

```javascript

Sub CompletelyRefit()
Dim a As Double, b As Double, c As Double, e As Double, f As Double, kk As Double, d As Double
Dim Alpha As Variant, Beta As Variant, Kt As Variant, updating As Variant
Dim index As Double, OrignalMxt As Variant, Alphanew As Variant
Dim index1 As Double, index2 As Double, index3 As Double, index4 As Double, index5 As Double
Dim sum1 As Double, sum2 As Double, kkk As Double, Array1 As Variant, Array2 As Variant

Worksheets("Raw Data").Activate
PartialMxt = Application.Evaluate("CR2:DL91")

Worksheets("Refitted Models").Activate
Matrix1 = Range("BA1:BA90").Value

For a = 1 To 21

b = [A1].End(xlToRight).Column
Range(Cells(1, b + 1), Cells(90, b + 1)) = Application.index(PartialMxt, , a)
b = [A1].End(xlToRight).Column

OrignalMxt = Range(Cells(1, 1), Cells(90, b)).Value
Alpha = Range("AR1:AR90").Value
Alphanew = Range("AR1:AR90").Value
Beta = Range("AU1:AU90").Value
Betanew = Range("AU1:AU90").Value

For c = 1 To b
Cells(c, 50) = 0
Next c

Kt = Range(Cells(1, 50), Cells(b, 50)).Value
Ktnew = Range(Cells(1, 50), Cells(b, 50)).Value

Array1 = Application.MMult(Alpha, WorksheetFunction.Transpose(Matrix1))
Array2 = Application.MMult(Beta, WorksheetFunction.Transpose(Kt))

For e = 1 To 90
   For f = 1 To b
      Cells(92 + e, f).Value = Array1(e, f) + Array2(e, f)
   Next f
Next e

updating = Range(Cells(93, 1), Cells(182, b)).Value

'update alpha coefficients

For kkk = 1 To 30
For d = 1 To 90
    index = 0
    For c = 1 To b
        index = index + OrignalMxt(d, c) - updating(d, c)
    Next c
    Alphanew(d, 1) = Alphanew(d, 1) + index / b
Next d

Array1 = Application.MMult(Alphanew, WorksheetFunction.Transpose(Matrix1))
Array2 = Application.MMult(Betanew, WorksheetFunction.Transpose(Ktnew))

For e = 1 To 90
   For f = 1 To b
      Cells(92 + e, f).Value = Array1(e, f) + Array2(e, f)
   Next f
Next e

updating = Range(Cells(93, 1), Cells(182, b)).Value

'update kt coefficients

For c = 1 To b
    index1 = 0
    index2 = 0
    For d = 1 To 90
        index1 = index1 + (OrignalMxt(d, c) - updating(d, c)) * Betanew(d, 1)
        index2 = index2 + Betanew(d, 1) ^ 2
    Next d
    Ktnew(c, 1) = Ktnew(c, 1) + index1 / index2
Next c

'make sure sum of kt coefficients is zero

sum1 = 0
For c = 1 To b
    sum1 = sum1 + Ktnew(c, 1)
Next c

For c = 1 To b
    Ktnew(c, 1) = Ktnew(c, 1) - sum1 / b
Next c

Array1 = Application.MMult(Alphanew, WorksheetFunction.Transpose(Matrix1))
Array2 = Application.MMult(Betanew, WorksheetFunction.Transpose(Ktnew))

For e = 1 To 90
   For f = 1 To b
      Cells(92 + e, f).Value = Array1(e, f) + Array2(e, f)
   Next f
Next e

updating = Range(Cells(93, 1), Cells(182, b)).Value

'update beta coefficients

For d = 1 To 90
    index3 = 0
    index4 = 0
    For c = 1 To b
        index3 = index3 + (OrignalMxt(d, c) - updating(d, c)) * Ktnew(c, 1)
        index4 = index4 + Ktnew(c, 1) ^ 2
    Next c
    Betanew(d, 1) = Betanew(d, 1) + index3 / index4
Next d

'make sure sum of beta coefficients is 1

sum2 = 0
For d = 1 To 90
    sum2 = sum2 + Betanew(d, 1)
Next d

For d = 1 To 90
    Betanew(d, 1) = Betanew(d, 1) / sum2
Next d

Array1 = Application.MMult(Alphanew, WorksheetFunction.Transpose(Matrix1))
Array2 = Application.MMult(Betanew, WorksheetFunction.Transpose(Ktnew))

For e = 1 To 90
   For f = 1 To b
      Cells(92 + e, f).Value = Array1(e, f) + Array2(e, f)
   Next f
Next e

updating = Range(Cells(93, 1), Cells(182, b)).Value

'calculate the deviation

index5 = 0
For d = 1 To 90
    For c = 1 To b
    index5 = index5 + (OrignalMxt(d, c) - updating(d, c)) ^ 2
    Next c
Next d

Next kkk

'write the alpha, beta and kt

Range(Cells(1, 36 + b), Cells(90, 36 + b)).Value = Alphanew
Range(Cells(1, 60 + b), Cells(90, 60 + b)).Value = Betanew
Range(Cells(1, 84 + b), Cells(b, 84 + b)).Value = Ktnew

'write the converged deviation of each Lee-Carter model
 
Range("DX1").Select
ActiveCell.End(xlDown).Select
ActiveCell.Offset(1).Select
ActiveCell.Value = index5

Next a
End Sub


Sub GetLifeExpectancy()
Dim a As Double
Worksheets("Refitted Life Expectancy").Activate
ProjectedMxt2011 = Range("DW2:EQ91").Value

For a = 1 To 21
    Range("EW2:EW91") = Application.index(ProjectedMxt2011, , a)
    Range("FF1").Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(1).Select
    ActiveCell.Value = Range("FC2").Value
Next a

End Sub


Sub TrueLifeExpectancy()
Dim a As Double
Worksheets("Raw Data").Activate
TrueMxt = Range("B2:AP91").Value

Worksheets("Life Expectancy").Activate
For a = 1 To 41
    Range("D2:D91") = Application.index(TrueMxt, , a)
    Range("L1").Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(1).Select
    ActiveCell.Value = Range("J2").Value
Next a

End Sub


Sub ProjectedLifeExpectancy()
Dim a As Double
Worksheets("Refitted Models").Activate
NewMxt = Range("A185:AO274").Value

Worksheets("Life Expectancy").Activate
For a = 1 To 41
    Range("D2:D91") = Application.index(NewMxt, , a)
    Range("M1").Select
    ActiveCell.End(xlDown).Select
    ActiveCell.Offset(1).Select
    ActiveCell.Value = Range("J2").Value
Next a

End Sub

```

