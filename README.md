# formula-parser
A library to parse and execute formulas in your application

Parse and execute formulas written in Visual Basic syntax.

# How it works
formula-parser is a recursive descent expression parser. The parser constructs an abstract syntax tree, which is then executed. 

For security, only white-listed functions from the .NET Framework can be invoked by the expression being executed.

# Sample expressions
Arithmetic, string, boolean and date expressions are supported.
```
(latestquote - previousquote) / previousquote
iif(isonlineorder, "Online", "Retail")
FirstName & " " & LastName
iif(IsNothing(Amount), "Amount missing",  Format(Amount, "c2"))
"Q" & ((OrderDate.Month - 1) \ 3 + 1)
MonthName(DatePart(DateInterval.Month, Today), false)
DateAdd(DateInterval.Day, 1, Today)
```
