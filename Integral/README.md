**Calculation of area under a curve**

Numerical methods for integration can be used to integrate functions, whether you have your equation or a data table that describes it (Nakamura, 1998). In some cases it is even possible that the numerical solution can be faster than the analytical solution (finding the antiderivative of the function), in case you are only interested in the numeric value of the integral.
In this section, three methods for numerical integration will be described: rectangular method (or Riemann sum on the left and on the right), trapezoidal method and Simpson's rule.
The first thing to do is create a window like the one shown in figure 13.1.

![MainWindow](https://user-images.githubusercontent.com/11558301/85922859-a69c8900-b85c-11ea-88a8-90a30e8ceb21.png)

**Figure 13.1**

This main window was given the name of Integral (remember that this is the name of the form, that is, the value of the _Name_ property).
Figure 13.1 also shows the names that were given to each of the controls. 
The equation text box is the one that receives the function to which the numerical integral will be calculated in the interval between _LowLim_ and _UpLim_. The sample text box receives the number of divisions that will be used for the calculation.
Let's now program the _Calculate_ button. To do this we double-click on the button and in the code window we will create the subroutine that will contain the code. Remember that by default, the subroutine for buttons is created with the _Click_ event. That is, at run time, the subroutine will be executed by clicking on the button. Subroutine 13.1 shows the code that will be executed when you press this button. This event can be changed (respond to the _double-click_ event for example).

**Subroutine 13.1.**

1 | Private   Sub Calc_Click()
-- | --
2 | Dim i,   n As Integer
3 | Dim ws   As Worksheet
4 | Dim x,   y As String
5 | Dim ll,   ul, sumIn, deltaX As Double
6 | Dim R   As Range
7 | n =   CInt(Integral.samples)
8 | ll =   CDbl(Integral.LowLim.Text)
9 | ul =   CDbl(Integral.UpLim.Text)
10 | deltaX   = (ul - ll) / n
11 | sumIn =   0#
12 | Set ws   = Worksheets("function")
13 | ws.Range("A:B").Clear
14 | ws.Range("A1")   = "x"
15 | ws.Range("B1")   = "f(x)"
16 | For i =   0 To n
17 | x = VBA.Format(ll + i * deltaX,   "0.0000000")
18 | ws.Range("A" & i + 2).Value   = CDbl(x)
19 | ws.Range("B" & i + 2).Value   = VBA.Format(Application.WorksheetFunction.Substitute _
20 | (Integral.equation.Text, "x",   "A" & i + 2), "0.00")
21 | Next i
22 | If   rectangular.Value = True Then
23 | LeftSum.Caption =   VBA.Format(LeftsumIn(ll, ul, n), "0.0000000")
24 | RightSum.Caption =   VBA.Format(RightsumIn(ll, ul, n), "0.0000000")
25 | End If
26 | If   trapezoidal.Value = True Then
27 | RTrapezoidal.Caption =   VBA.Format(trapezoid(ll, ul, n), "0.0000000")
28 | End If
29 | If   simpson.Value = True And (n / 2) = Int(n / 2) Then
30 | RSimpson.Caption =   VBA.Format(fsimpson(ll, ul, n), "0.0000000")
31 | Else
32 | RSimpson.Caption = "The number of   samples must be an even number"
33 | End If
34 | Set R =   ws.Range("A1:B" & n + 1)
35 | Call   ChartPaint(R, n)
36 | End Sub

It is important to take into account that if you are going to enter a function that contains the letter x, such as EXP, it must be done using capital letters, to differentiate it from the variable "x". Additionally, the equation must begin with the "=" sign.
Now we will proceed to explain the most relevant aspects of the code.
In line 3 the variable "ws" has been declared as a _Worksheet_ type. This has some advantages that I find very useful, since the Intellisense algorithm of VBA-Excel recognizes the variable as an object and it is activated when writing the name of the variable, showing its properties and methods as shown in the figure 13.2. This statement is not necessary, but it has some incredible advantages, starting with the fact that it is not necessary to write _Worksheets("function")_ every time. Of course, at this point you already know that you can use a **With** block, but the Intellisense function will not be activated when using this alternative.

![image](https://user-images.githubusercontent.com/11558301/85922928-19a5ff80-b85d-11ea-87fd-2550acb4b7e3.png)

**Figure 13.2**

When textboxes are used to capture user information, it is necessary to take into account that VBA-Excel treats its contents as text strings, regardless of whether they are numbers. For this reason, in line 7, the **CInt** function was used to convert the numeric text input into an **Integer** number.
In lines 8 and 9 the **CDbl** function has been used to convert the entries of the text boxes into **Double** type variables
In line 11 we see that the variable sum is equal to 0#. This is the same as 0.0. The change is done by VBA-Excel automatically.
In line 12 we make the variable "ws" equal to the Worksheets("function") object. When working with objects it is necessary to use the **Set** instruction.
In line 13, we clear columns A and B so that the values of the function to be plotted can be written to them. The values of the X axis will be between _ll_ and _ul_ and will be spaced a distance equal to _deltaX_.
In line 19 is where the application trick is. We are used to using the letter "x" as a variable in our equations. So what the **Substitute** instruction does is to change the letter "x" that the user enters in the equation text box (figure 13.1) by A2, A3, A4 and so on ("A" & i + 2, with the variable i varying within a loop **For**) and write it in the cells of column B. These values correspond therefore to those of the function f(x).
Between lines 22 and 25 the values of the Riemann sum are calculated from the right and from the left. For each of them, a function was written. For the sum on the left, the function is shown in subroutine 13.2 and for the sum on the right, the function is shown in subroutine 13.3.

**Subroutine 13.2.**

1 | Function   LeftsumIn(ByVal lowl As Double, ByVal upl As Double, ByVal s As Integer) As   Double
-- | --
2 | Dim i   As Integer
3 | Dim   deltaX As Double
4 | LeftsumIn   = 0
5 | deltaX   = Abs((upl - lowl) / s)
6 | For i =   0 To s - 1
7 | LeftsumIn = LeftsumIn +   Worksheets("function").Range("B" & i + 2) * deltaX
8 | Next i
9 | End   Function

**Subroutine 13.3.**

1 | Function   RightsumIn(ByVal lowl As Double, ByVal upl As Double, ByVal s As Integer) As   Double
-- | --
2 | Dim i   As Integer
3 | Dim deltaX   As Double
4 | RightsumIn   = 0
5 | deltaX   = Abs((upl - lowl) / s)
6 | For i =   1 To s
7 | RightsumIn = RightsumIn +   Worksheets("function").Range("B" & i + 2) * deltaX
8 | Next i
9 | End   Function

Between lines 26 and 28, the calculations are made according to the trapezoidal rule. The function that does this is shown in subroutine 13.4. 

**Subroutine 13.4.**

1 | Function   trapezoid(ByVal lowl As Double, ByVal upl As Double, ByVal s As Integer) As   Double
-- | --
2 | Dim i   As Integer
3 | Dim   deltaX As Double
4 | trapezoid   = 0
5 | deltaX   = Abs((upl - lowl) / s)
6 | For i =   0 To s
7 | trapezoid = trapezoid +   (Worksheets("function").Range("B" & i + 2) + _
8 | Worksheets("function").Range("B" & i + 3)) *   deltaX * 0.5
9 | Next i
10 | End   Function

Between lines 29 and 33, the calculations are made using Simpson's rule. The function that does this is shown in the subroutine 13.5.

**Subroutine 13.5.**

1 | Function   fsimpson(ByVal lowl As Double, ByVal upl As Double, ByVal s As Integer) As   Double
-- | --
2 | Dim i   As Integer
3 | Dim factor,   f2, f4 As Double
4 | fsimpson   = 0
5 | f2 = 0
6 | f4 = 0
7 | factor   = Abs((upl - lowl) / (3 * s))
8 | If s =   2 Then
9 | fsimpson   = (Worksheets("function").Range("B2") + _
10 | Worksheets("function").Range("B" & s + 2)) *   factor
11 | End If
12 | If s   <> 2 Then
13 | For i = 1 To s / 2
14 | f4 = f4 + 4 *   Worksheets("function").Range("B" & 2 * i + 1)
15 | Next i
16 | For j = 1 To (s / 2) - 1
17 | f2 = f2 + 2 *   Worksheets("function").Range("B" & 2 * j + 2)
18 | Next j
19 | End If
20 | fsimpson   = factor * (f2 + f4 + Worksheets("function").Range("B2")   + _
21 | Worksheets("function").Range("B" & s + 2))
22 | End   Function

In line 35, the subroutine which is responsible for constructing the graph of the curve and the area below it according to the specified interval is invoked. This subroutine is the one shown below.

**Subroutine 13.6.**

1 | Sub   ChartPaint(ByRef R As Range, ByVal s As Integer)
-- | --
2 | Dim n   As Integer
3 | Dim   xlabels As Range
4 | Set   xlabels = Worksheets("function").Range("B2:B" & s +   1)
5 | n = Worksheets("function").ChartObjects.Count
6 | If n   <> 0 Then
7 | Worksheets("function").ChartObjects.Delete
8 | End If
9 | With   Worksheets("function").ChartObjects.Add _
10 | (Left:=200, Width:=375, Top:=60,   Height:=225)
11 | .Chart.SetSourceData Source:=R
12 | .Chart.SeriesCollection(1).Delete
13 | .Chart.SeriesCollection(1).XValues =   xlabels
14 | .Chart.ChartType = xlArea
15 | .Chart.HasTitle = True
16 | .Chart.ChartTitle.Text = "f(x)"   & Integral.equation
17 | .Chart.Parent.Name = "function"
18 | End   With
19 | End Sub

The code for the Cancel/Close button is shown below.

**Subroutine 13.7.**

1 | Private   Sub CancelIt_Click()
-- | --
2 | Unload   Me
3 | End Sub

To show the main window of our application, we have added a command button in the spreadsheet "function", as shown in figure 13.3. 

![image](https://user-images.githubusercontent.com/11558301/85923011-f2036700-b85d-11ea-9329-7dcc14a40f2c.png)

**Figure 13.3**

To add this button you must go to the "DEVELOPER" tab and there press the "Insert" button. When doing so, the window shown in figure 13.4 appears. The command button is inserted with the control enclosed in the black rectangle.

![image](https://user-images.githubusercontent.com/11558301/85923019-ffb8ec80-b85d-11ea-8376-3aa06e798056.png)

**Figure 13.4**

In order to edit the properties of the button, we must press the "Design Mode" button, which is next to the "Insert" button. Once this is done, click on the button with the right mouse button to see the Properties window. If you cannot do it with the right mouse button, you can also press the "Properties" button next to the "Design Mode" button. The Properties window of the button in this example is shown in figure 13.5.

![image](https://user-images.githubusercontent.com/11558301/85923025-0e070880-b85e-11ea-964d-3caad1750618.png)

**Figure 13.5**

The code that will be executed when you press this button is the one shown below. To access the code window, in edit mode, double-click on it.

**Subroutine 13.8.**

1 | Private   Sub Go_Click()
-- | --
2 | Integral.Show
3 | End Sub

Notice that “Integral” is what we have called the form shown in figure 13.1.
Remember also that since we created functions to do the calculations, these can also be used from any spreadsheet in our Excel book. Figure 13.6 shows the appearance of the spreadsheet with each of the functions used.

![image](https://user-images.githubusercontent.com/11558301/85923047-2d059a80-b85e-11ea-9ee0-53023d668d3a.png)

**Figure 13.6**

In cells K5 and L5 we place the extremes of the interval in which we want to calculate the integral. In cell M5, we place the following instruction:
=COUNTA(A:A)-1
This function will allow to count the quantity of values that are available for the calculation of the integral.
In cell L6 the following instruction is placed:
=RightsumIn(K5;L5;M5)
In cell L7 the following instruction is placed:
=LeftsumIn(K5;L5;M5)
In the cell L8 the following instruction is placed:
=trapezoid(K5;L5;M5)
In cell L9 the following instruction is placed:
=fsimpson(K5;L5;M5)
