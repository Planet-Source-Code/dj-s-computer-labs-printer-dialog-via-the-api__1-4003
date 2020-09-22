Attribute VB_Name = "modStart"
Option Explicit

Sub main()
Dim P As New clsPrintDialog
  P.Flags = cdlPDPageNums + cdlPDDisablePrintToFile + cdlPDNoSelection
  P.Min = 1
  P.FromPage = 3
  P.ToPage = 5
  P.Max = 100
  P.ShowPrinter
  Debug.Print Printer.DeviceName
  Debug.Print Printer.Copies
  Debug.Print P.FromPage
  Debug.Print P.ToPage
  ' Write here Print Code with Printer Object...
End Sub
