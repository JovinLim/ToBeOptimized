using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Reflection;
using System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using IronXL;
using X = Microsoft.Office.Interop.Excel;

namespace TestParam
{ 
    [Transaction(TransactionMode.Manual)]
    [Regeneration(RegenerationOption.Manual)]
    public class ExcelGeneration : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            //Get application and documnet objects
            UIApplication uiapp = commandData.Application;
            Document doc = uiapp.ActiveUIDocument.Document;
            Autodesk.Revit.DB.View active = doc.ActiveView;
            string cDir = Directory.GetCurrentDirectory();
            string tDir = @"c:\temp";
            string tPath = @"c:\temp\SOA";

            // Creating temporary directory to store excel file
            try
            {
                if (!Directory.Exists(tDir))
                {
                    Directory.CreateDirectory(tDir);
                }
            }

            catch (Exception ex)
            {
                message = ex.Message;
                TaskDialog.Show("failed", message);
                return Result.Failed;
            }

            // Creating lists to store information
            List<string> clinicType = new List<string>
            {
                "Orthopaedic"
            };

            //DICTIONARY FOR ROOM TYPES
            Dictionary<string, double> roomType = new Dictionary<string, double>
            {
                {"RECEPTION", 50 },
                {"WAITING", 20 },
                {"CONSULTATION/EXAMINATION ROOM" , 15 },
                {"TOILET", 8.6 },
                {"STAFF OFFICE", 26 },
                {"INTERVIEW ROOM", 13 },
            };

            TaskDialog dialog = new TaskDialog("Apply Schedule of Accommodations");
            dialog.MainContent = "Please input your schedule of accommodations in the following excel sheet that will be opened up. Click Yes to open up the excel sheet";
            dialog.CommonButtons = TaskDialogCommonButtons.Yes | TaskDialogCommonButtons.Cancel;
            dialog.DefaultButton = TaskDialogResult.Yes; // optional, will default to first button in box if unassigned
            TaskDialogResult response = dialog.Show();

            try
            {
                if (response == TaskDialogResult.Yes)
                {
                    try
                    {
                        X.Application excel = new X.Application();
                        excel.DisplayAlerts = false;
                        X.Workbook SOA = excel.Workbooks.Add();
                        X.Worksheet main = (X.Worksheet)excel.Worksheets.Add();
                        main.Name = "READ ME";
                        // Establish column headings in cells A1 and B1.
                        main.Cells[1, "A"] = "Instructions";
                        main.Cells[1, "B"] = "Clinic Types";
                        main.Cells[2, "A"] = "test";

                        // Creating sheet for each type of clinic
                        for (int i = 0; i < clinicType.Count; i++)
                        {
                            try
                            {
                                X.Worksheet addSheet = (X.Worksheet)excel.Worksheets.Add(After: main);
                                addSheet.Name = clinicType[i];
                                // Header
                                addSheet.Cells[1, "A"] = "Department";
                                addSheet.Cells[1, "B"] = clinicType[i];

                                //Room Types
                                addSheet.Cells[2, "A"] = "Room Type";
                                int room_count = 3;
                                foreach (KeyValuePair<string, double> ele in roomType)
                                {
                                    addSheet.Cells[room_count, "A"] = ele.Key;
                                    addSheet.Cells[room_count, "B"] = ele.Value;
                                    int roomType_rowNum = room_count;
                                    addSheet.Cells[room_count, "D"].Formula = string.Format("=B{0}*C{0}", roomType_rowNum.ToString());
                                    room_count++;
                                };


                                // Unit Area
                                addSheet.Cells[2, "B"] = "Unit Area/sqm";

                                // Quantity
                                addSheet.Cells[2, "C"] = "Quantity";

                                // Total Area + Formula
                                addSheet.Cells[2, "D"] = "Total Area/sqm";
                                addSheet.Cells[15, "D"] = "=SUM(D3:D14)";
                            }

                            catch (Exception ex)
                            {
                                message = ex.Message;
                                TaskDialog.Show("failed", message + "Failed at " + i.ToString());
                                return Result.Failed;
                            }

                        }

                        SOA.SaveAs(tPath + ".xlsx");
                        excel.Workbooks.Open(Filename: tPath + ".xlsx", UpdateLinks: true, ReadOnly: false, Editable: true, Local: true);
                        SOA.Worksheets["Sheet1"].Delete();
                        excel.Visible = true;
                    }
                    catch (Exception ex)
                    {
                        message = ex.Message;
                        TaskDialog.Show("failed", message);
                        return Result.Failed;
                    }
                }

                //Returning value of result for Execute command
                return Result.Succeeded;
            }

            //If the user right-clicks or presses Esc, handle the exception
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
