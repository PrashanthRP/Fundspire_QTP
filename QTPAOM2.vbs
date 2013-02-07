testSet = Array( "TC_01", "TC_02" )
Set app = CreateObject("QuickTest.Application")


If app.Launched <> True then  ' If QuickTest is not yet open 
   app.Launch ' Start QuickTest (with the correct add-ins loaded) 
End If 


app.Visible = True ' Make the QuickTest application visible 

For i = LBound( testSet ) To UBound( testSet ) 
   testPath = "D:\Fundspire_QTP\Test Scripts\" & testSet( i )
   if testPath = "" Then
      Reporter.ReportEvent micFail, "Executing Set", "Test " & testSet( i ) & " not found"
      Exit For
   End If


   app.Open testpath ' Open the test 
   app.Test.Settings.Run.IterationMode = "oneIteration" ' Run only one iteration
   app.Test.Settings.Run.OnError = "Stop"
   Set qtResultsOpt = CreateObject("QuickTest.RunResultsOptions") ' Create the Run Results Options object 
   qtResultsOpt.ResultsLocation = "D:\Fundspire_QTP\Test Results" ' Set the results location 
   app.Test.Run qtResultsOpt, True ' Run the test and wait for return
   if StrComp( app.Test.LastRunResults.Status, "Passed" ) <> 0 Then
      Reporter.ReportEvent micFail, testSet( i ), app.Test.LastRunResults.LastError
     End If
Next
Set app = Nothing ' Release the Application object