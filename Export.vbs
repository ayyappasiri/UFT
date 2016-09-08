Option Explicit
Public QCConnection, sSubject, sDomain, sPrj, sSubj, WriteFile
'==========================================================================
'
' Quality Center Test Case Exporter
'
' COMMENT:
' Exports test cases with design steps to Excel
'
'==========================================================================
Public Sub DriverTestSets()
Dim sEID, sPW, sList, sItem As String
Dim Message, Title
    'Get the login info
    Title = "Get Login Info"    ' Set title.
    ' Display message, title
    Message = "Enter your EID"    ' Set prompt.
    sEID = InputBox(Message, Title)
    Message = "Enter your Password"    ' Set prompt.
    sPW = InputBox(Message, Title)
'Return the TDConnection object.
Set QCConnection = CreateObject("TDApiOle80.TDConnection")

QCConnection.InitConnectionEx "http://roc-hpqc/qcbin/" '<-- Change me.
QCConnection.Login sEID, sPW
If (QCConnection.LoggedIn <> True) Then
    MsgBox "QC User Authentication Failed"
    Quit
End If
    'Get the lroject name and root folder
    Title = "Get Project Name"    ' Set title.
    ' Display message, title
    Message = "Enter QC Project Name"    ' Set prompt.
    sPrj = InputBox(Message, Title)
'Dim sDomain, sPrj
sDomain = <<"doman name">>   '<-- Change me.
'sPrj = <<"project name">> '<-- Change me.
QCConnection.Connect sDomain, sPrj
If (QCConnection.Connected <> True) Then
    MsgBox "QC Project Failed to Connect to " & sPrj
    Quit
End If
Next_Export:
    ' Display message, title
    Message = "Enter Folder Name"    ' Set prompt.
    sSubj = InputBox(Message, Title)
sSubject = "Subject\" & sSubj
Call ExportTestCases(sSubject)
 Dim Msg, Style, Help, Ctxt, Response, MyString
   'All done or more pre-reads?
    Msg = WriteFile & " ihas been exported.  " _
          & "Click OK to generate another export, or Cancel to Exit " 'Define message.
    Style = vbOKCancel    ' Define buttons.
    Title = sSubj & " Export Complete"                   ' Define title.
            ' Display message.
    Response = MsgBox(Msg, Style, Title, Help, Ctxt)
    If Response = vbOK Then    ' User chose Yes.
        GoTo Next_Export
    Else    ' User chose to exit
        GoTo Exit_Sub
    End If
   
Exit_Sub:
QCConnection.Disconnect
QCConnection.Logout
QCConnection.ReleaseConnection
MsgBox "All Done"
End Sub
 
'Export test cases for the Test Lab node.
'
'@param:    strNodeByPath   String for the node path in Test Lab.
'
'@return:                   No return value.
Function ExportTestCases(strNodeByPath)
    Dim Excel, Sheet
    Set Excel = CreateObject("Excel.Application") 'Open Excel
    Excel.Workbooks.Add        '() 'Add a new workbook
    'Get the first worksheet.
    Set Sheet = Excel.ActiveSheet
   
    Sheet.Name = "Tests"
   
    With Sheet.Range("A1:U1")
        .Font.Name = "Arial"
        .Font.FontStyle = "Bold"
        .Font.Size = 14
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Interior.ColorIndex = 34 'Light Turquoise
    End With
'-------------------------------------------------------------------------------
'  List of field names that export to Excel
'  Change the names to your project's field names 
'-------------------------------------------------------------------------------    Sheet.Cells(1, 1) = "Service Line"
    Sheet.Cells(1, 2) = "Service"
    Sheet.Cells(1, 3) = "Process"
    Sheet.Cells(1, 4) = "Sub-Process"
    Sheet.Cells(1, 5) = "Activity"
    Sheet.Cells(1, 6) = "Test ID"
    Sheet.Cells(1, 7) = "Test Name"
    Sheet.Cells(1, 8) = "Type"
    Sheet.Cells(1, 9) = "Destription"
    Sheet.Cells(1, 10) = "Designer (Owner)"
    Sheet.Cells(1, 11) = "Template"
    Sheet.Cells(1, 12) = "Subject (Folder Name)"
    Sheet.Cells(1, 13) = "Attachment"
    Sheet.Cells(1, 14) = "Step Name"
    Sheet.Cells(1, 15) = "Step Description"
    Sheet.Cells(1, 16) = "Expected Result"
    Sheet.Cells(1, 17) = "Action"
    Sheet.Cells(1, 18) = "Object Name"
    Sheet.Cells(1, 19) = "Object Type"
    Sheet.Cells(1, 20) = "Value"
    Sheet.Cells(1, 21) = "Additional Info"
       
    Dim TreeMgr, TestTree, TestFactory, TestList
    Set TreeMgr = QCConnection.TreeManager
    'Specify the folder path in TestPlan, all the tests under that folder will be exported.
    Set TestTree = TreeMgr.NodeByPath(strNodeByPath)
    Set TestFactory = TestTree.TestFactory
    Set TestList = TestFactory.NewList("") 'Get a list of all from node.
    'Specify Array to contain all nodes of subject tree.
    Dim NodesList()
    ReDim Preserve NodesList(0)
    'Assign root node of subject tree as NodeByPath node.
    NodesList(0) = TestTree.Path
   
    'Gets subnodes and return list in array NodesList
    Call GetNodesList(TestTree, NodesList)
    Dim Row, Node, TestCase
    Row = 2
    For Each Node In NodesList
        Set TestTree = TreeMgr.NodeByPath(Node)
        Set TestFactory = TestTree.TestFactory
        Set TestList = TestFactory.NewList("") 'Get a list of all from node.
        'Iterate through all the tests.
        For Each TestCase In TestList
            Dim DesignStepFactory, DesignStep, DesignStepList
            Set DesignStepFactory = TestCase.DesignStepFactory
            Set DesignStepList = DesignStepFactory.NewList("")
           
            'Debug.Print DesignStepList.Fields.Count
            'Dim ctr As Integer
            'For ctr = 1 To DesignStepList.Fields.Count
            '    Debug.Print DesignStepList.Fields(ctr), Name
            'Next ctr
'-------------------------------------------------------------------------------
'  Change the field names to your project's field names 
'-------------------------------------------------------------------------------               
            If DesignStepList.Count = 0 Then
                Sheet.Cells(Row, 1).Value = Trim(TestCase.Field("TS_USER_09"))
                Sheet.Cells(Row, 2).Value = Trim(TestCase.Field("TS_USER_03"))
                Sheet.Cells(Row, 3).Value = Trim(TestCase.Field("TS_USER_07"))
                Sheet.Cells(Row, 4).Value = Trim(TestCase.Field("TS_USER_04"))
                Sheet.Cells(Row, 5).Value = Trim(TestCase.Field("TS_USER_05"))
                Sheet.Cells(Row, 6).Value = Trim(TestCase.Field("TS_TEST_ID"))
                Sheet.Cells(Row, 7).Value = Trim(TestCase.Field("TS_NAME"))
                Sheet.Cells(Row, 8).Value = Trim(TestCase.Field("TS_TYPE"))
                Sheet.Cells(Row, 9).Value = Trim(TestCase.Field("TS_DESCRIPTION"))
                Sheet.Cells(Row, 10).Value = Trim(TestCase.Field("TS_RESPONSIBLE"))
                Sheet.Cells(Row, 11).Value = Trim(TestCase.Field("TS_TEMPLATE"))
                Sheet.Cells(Row, 12).Value = Trim(TestCase.Field("TS_SUBJECT").Path)
                Row = Row + 1
            Else
                For Each DesignStep In DesignStepList
                    'Save a specified set of fields.
                Sheet.Cells(Row, 1).Value = Trim(TestCase.Field("TS_USER_09"))
                Sheet.Cells(Row, 2).Value = Trim(TestCase.Field("TS_USER_03"))
                Sheet.Cells(Row, 3).Value = Trim(TestCase.Field("TS_USER_07"))
                Sheet.Cells(Row, 4).Value = Trim(TestCase.Field("TS_USER_04"))
                Sheet.Cells(Row, 5).Value = Trim(TestCase.Field("TS_USER_05"))
                Sheet.Cells(Row, 6).Value = Trim(TestCase.Field("TS_TEST_ID"))
                Sheet.Cells(Row, 7).Value = Trim(TestCase.Field("TS_NAME"))
                Sheet.Cells(Row, 8).Value = Trim(TestCase.Field("TS_TYPE"))
                Sheet.Cells(Row, 9).Value = Trim(TestCase.Field("TS_DESCRIPTION"))
                Sheet.Cells(Row, 10).Value = Trim(TestCase.Field("TS_RESPONSIBLE"))
                Sheet.Cells(Row, 11).Value = Trim(TestCase.Field("TS_TEMPLATE"))
                Sheet.Cells(Row, 12).Value = Trim(TestCase.Field("TS_SUBJECT").Path)
               
                    'Save the specified design steps.
                    Sheet.Cells(Row, 13).Value = Trim(DesignStep.Field("DS_ATTACHMENT"))
                    Sheet.Cells(Row, 14).Value = Trim(DesignStep.Field("DS_STEP_NAME"))
                    Sheet.Cells(Row, 15).Value = Trim(DesignStep.Field("DS_DESCRIPTION"))
                    Sheet.Cells(Row, 16).Value = Trim(DesignStep.Field("DS_EXPECTED"))
                    Sheet.Cells(Row, 17).Value = Trim(DesignStep.Field("DS_USER_01"))
                    Sheet.Cells(Row, 18).Value = Trim(DesignStep.Field("DS_USER_02"))
                    Sheet.Cells(Row, 19).Value = Trim(DesignStep.Field("DS_USER_03"))
                    Sheet.Cells(Row, 20).Value = Trim(DesignStep.Field("DS_USER_04"))
                    Sheet.Cells(Row, 21).Value = Trim(DesignStep.Field("DS_USER_07"))
                    Row = Row + 1
                Next
            End If
        Next
    Next
   
    'Excel.Columns.AutoFit
    Excel.Columns("A:U").ColumnWidth = 12
   
    'Set Auto Filter mode.
    If Not Sheet.AutoFilterMode Then
        Sheet.Range("A1").AutoFilter
    End If
   
    'Freeze first row.
    Sheet.Range("A2").Select
    Excel.ActiveWindow.FreezePanes = True
   
    'sSubj = Right(sSubj, 32)
'-------------------------------------------------------------------------------
'  Change the folder path/filename to suit your file system
'-------------------------------------------------------------------------------       
    WriteFile = ""C:\Users\" & Environ$("User") & "\Desktop\Tests"_"& sPrj & "_" & sSubj & "_TestCases.xls"
   
    'Save the newly created workbook and close Excel.
    Excel.ActiveWorkbook.SaveAs (WriteFile)
    Excel.Quit
   
    Set Excel = Nothing
    Set DesignStepList = Nothing
    Set DesignStepFactory = Nothing
    Set TestList = Nothing
    Set TestFactory = Nothing
    Set TestTree = Nothing
    Set TreeMgr = Nothing
End Function

''
'Returns a NodesList array for all children of a given node of a tree.
'
'@param:    Node        Node in a Test Lab tree.
'
'@param:    NodesList   Array to store all children of a given node of a tree.
'
'@return:   No explicit return value.
Function GetNodesList(ByVal Node, ByRef NodesList)
    Dim i
    'Run on all children nodes
    For i = 1 To Node.Count
        Dim NewUpper
        'Add more space to dynamic array
        NewUpper = UBound(NodesList) + 1
        ReDim Preserve NodesList(NewUpper)
       
        'Add node path to array
        NodesList(NewUpper) = Node.Child(i).Path
       
        'If current node has a child then get path on child nodes too.
        If Node.Child(i).Count >= 1 Then
            Call GetNodesList(Node.Child(i), NodesList)
        End If
    Next
End Function