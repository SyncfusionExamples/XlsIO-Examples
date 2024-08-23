using Syncfusion.Office;
using Syncfusion.XlsIO;
using System.Data.Common;

namespace Mulitple_depenedent_dropdown
{
    class Program
    {
        public static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream(Path.GetFullPath(@"Data/InputTemplate.xlsx"), FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Set name range
                IName name1 = workbook.Names.Add("Country");
                name1.RefersToRange = worksheet.Range["A2:A5"];

                IName name2 = workbook.Names.Add("India");
                name2.RefersToRange = worksheet.Range["B2:B6"];

                IName name3 = workbook.Names.Add("Brazil");
                name3.RefersToRange = worksheet.Range["B7:B10"];

                IName name4 = workbook.Names.Add("Australia");
                name4.RefersToRange = worksheet.Range["B11:B13"];

                IName name5 = workbook.Names.Add("USA");
                name5.RefersToRange = worksheet.Range["B14:B16"];

                worksheet.Range["E1"].Text = "Country";
                worksheet.Range["E1"].CellStyle.Font.Bold = true;

                //Data validation in E2 
                IDataValidation validation = worksheet.Range["E2"].DataValidation;
                validation.AllowType = ExcelDataType.User;
                validation.FirstFormula = "=Country";

                //Shows the error message
                validation.ErrorBoxText = "Enter the valid country";
                validation.ErrorBoxTitle = "ERROR";
                validation.PromptBoxText = "Enter the country";
                validation.ShowPromptBox = true;

                worksheet.Range["F1"].Text = "States";
                worksheet.Range["F1"].CellStyle.Font.Bold = true;

                //Data validation in F2
                IDataValidation validation2 = worksheet.Range["F2"].DataValidation;
                validation2.AllowType = ExcelDataType.User;
                validation2.FirstFormula = "=Indirect(E2)";

                //Shows the error message
                validation2.ErrorBoxText = "Enter the valid states";
                validation2.ErrorBoxTitle = "ERROR";
                validation2.PromptBoxText = "Enter the state";
                validation2.ShowPromptBox = true;

                //Creating Vba project
                IVbaProject project = workbook.VbaProject;

                //Accessing vba modules collection
                IVbaModules vbaModules = project.Modules;

                //Accessing sheet module
                IVbaModule vbaModule = vbaModules[worksheet.CodeName];

                //Adding vba code to the module
                vbaModule.Code = @"Private Sub Worksheet_Change(ByVal Target As Range)
                    Dim rngDropdown As Range
                    Dim oldValue As String
                    Dim newValue As String
                    Dim DelimiterType As String
                    DelimiterType = "", "" ' Assuming you want to use a comma followed by a space as the delimiter
                    
                    If Target.Count > 1 Then Exit Sub
                    On Error Resume Next
                    If Target.Column <> 5 And Target.Column <> 6 Then GoTo exitError
                    Set rngDropdown = Cells.SpecialCells(xlCellTypeAllValidation)
                    On Error GoTo exitError
                    If rngDropdown Is Nothing Then GoTo exitError
                    If Intersect(Target, rngDropdown) Is Nothing Then
                        ' Do nothing for non-dropdown cells
                    Else
                        Application.EnableEvents = False
                        newValue = Target.Value
                        Application.Undo
                        oldValue = Target.Value
                        Target.Value = newValue
                        
                        ' Check if the change was in E2 and clear F2
                        If Target.Address = ""$E$2"" Then
                            Range(""F2"").ClearContents
                        ElseIf Target.Address = ""$F$2"" Then
                            ' Append new value to F2 with the delimiter if F2 is not empty
                            If oldValue <> """" Then
                                If newValue <> """" Then
                                    If oldValue = newValue Or _
                                       InStr(1, oldValue, DelimiterType & newValue) > 0 Or _
                                       InStr(1, oldValue, newValue & DelimiterType) > 0 Then
                                        ' Do nothing if the new value is the same or already exists
                                    Else
                                        ' Append the new value with the delimiter
                                        Target.Value = oldValue & DelimiterType & newValue
                                    End If
                                End If
                            Else
                                ' Set F2 directly to the new value if it's empty
                                Target.Value = newValue
                            End If
                        End If
                        
                        ' Re-enable events after handling the change
                        Application.EnableEvents = True
                    End If
                exitError:
                    Application.EnableEvents = True
                End Sub
                
                Private Sub Worksheet_SelectionChange(ByVal Target As Range)
                End Sub";

                //Saving the workbook as stream
                FileStream stream = new FileStream("Output.xlsm", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(stream);
            }
        }
    }
}
