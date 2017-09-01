''****************************************************************************************************************************
'      /**
' * Copyright (c) 2007 Juniper Networks, Inc.
' * All Rights Reserved
' *
' * JUNIPER PROPRIETARY/CONFIDENTIAL. Use is subject to license terms.
' *
' */
''*****************************************************************************************************************************

Dim Test_path

Test_path="C:\Jenkins\workspace\ITQA_FT_UFT_SAP\FrameworkDriver"

'QuickTestApplication 'object variable decalaration

Dim qtApp

'Create the Application object

Set qtApp = CreateObject("QuickTest.Application")

'Launch QTP
 
qtApp.Launch 

'Check Application is visible

qtApp.Visible = True

' Set QuickTest run options

'qtApp.Options.Run.RunMode = "Fast"

qtApp.Options.Run.ViewResults = True


' Open the test in read-only mode

qtApp.Open Test_path, True 

'set run settings for the test

Set qtTest = qtApp.Test

' Run the test

qtTest.Run 

' Close the test

qtTest.Close 

qtApp.quit

Set qtTest = Nothing ' Release the Test object
Set qtApp = Nothing ' Release the Application object
