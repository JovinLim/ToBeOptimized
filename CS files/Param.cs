using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using X = Microsoft.Office.Interop.Excel;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace TBO_Plugin
{
    public partial class Param : System.Windows.Forms.Form
    {
        Document Doc;
        string tDir = @"c:\temp";
        string tPath = @"c:\temp\param.xlsx";

        //Parameter List
        List<string> paramList = new List<string>
            {
            "SCDF Fire Code - 2018",
            "BCA Accessibility Code - 2019",
            "TR42 (Acute General Hospitals) - 2015",
            "TR59 (Community Hospitals) - 2017",
            "TR65 (Polyclinics) - 2018",
            };

        //Objectives List
        List<string> objList = new List<string>
            {
            "Minimize waiting area",
            "Shortest average distance between each C/E room to a waiting room",
            "Maximize number of consultation/examination rooms",
            };

        //Objectives Unit List
        List<string> uList = new List<string>
            {
            "sqm",
            "m",
            "None",
            };

        public Param(Document doc)
        {
            InitializeComponent();
            Doc = doc;
            List<string> strings = new List<string>();
            string cDir = Directory.GetCurrentDirectory();

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
                string message = ex.Message;
                TaskDialog.Show("failed", message);
            }

            // Checking if excel file exists
            try
            {
                if (!File.Exists(tPath))
                {
                    X.Application excel = new X.Application();
                    excel.DisplayAlerts = false;
                    X.Workbook paramWb = excel.Workbooks.Add();
                    X.Worksheet param = (X.Worksheet)paramWb.Worksheets.Add();
                    param.Name = "Parameters";
                    // Creating Excel Worksheet for Parameters
                    param.Cells[1, "A"] = "Param_Name";
                    param.Cells[1, "B"] = "Y/N";
                    for (int p = 0; p < paramList.Count; p++)
                    {
                        param.Cells[p + 2, "A"] = paramList[p];
                        this.checkedListBox1.Items.Add(paramList[p], false);
                    }

                    // Creating Excel Worksheet for Objectives
                    X.Worksheet objectives = (X.Worksheet)paramWb.Worksheets.Add();
                    objectives.Name = "Objectives";
                    objectives.Cells[1, "A"] = "Obj_Name";
                    objectives.Cells[1, "B"] = "Y/N";
                    objectives.Cells[1, "C"] = "Units";
                    for (int o = 0; o < objList.Count; o++)
                    {
                        objectives.Cells[o + 2, "A"] = objList[o];
                        objectives.Cells[o + 2, "C"] = uList[o];
                        this.checkedListBox2.Items.Add(objList[o], false);
                    }

                    paramWb.SaveAs(tPath);
                    paramWb.Close(0);
                    excel.Quit();
                }

                else
                {
                    X.Application excel = new X.Application();
                    excel.DisplayAlerts = false;
                    X.Workbook paramWb = excel.Workbooks.Open(tPath);
                    X._Worksheet param = (X._Worksheet)paramWb.Sheets["Parameters"];
                    X._Worksheet objectives = (X._Worksheet)paramWb.Sheets["Objectives"];

                    // Creating Excel Worksheet for Parameters
                    for (int p = 0; p < paramList.Count; p++)
                    {
                        param.Cells[p+2, "A"] = paramList[p];
                        this.checkedListBox1.Items.Add(paramList[p], false);
                    }

                    // Creating Excel Worksheet for Objectives
                    for (int o = 0; o < objList.Count; o++)
                    {
                        objectives.Cells[o+2, "A"] = objList[o];
                        objectives.Cells[o + 2, "C"] = uList[o];
                        this.checkedListBox2.Items.Add(objList[o], false);
                    }

                    paramWb.SaveAs(tPath);
                    paramWb.Close(0);
                    excel.Quit();
                }
            }

            catch (Exception ex)
            {
                string message = ex.Message;
                TaskDialog.Show("failed", message);
            }
        }

        private void button1_Click(object sender, EventArgs e) //Apply button
        {
            try
            {
                X.Application excel = new X.Application();
                excel.DisplayAlerts = false;
                X.Workbook paramWb = excel.Workbooks.Open(tPath);
                X._Worksheet param = (X._Worksheet)paramWb.Sheets["Parameters"];
                X._Worksheet objectives = (X._Worksheet)paramWb.Sheets["Objectives"];
                int objCount = 0;

                for (int i = 0; i < checkedListBox2.Items.Count; i++)
                {
                    if (checkedListBox2.GetItemChecked(i) == true)
                    {
                        objCount++;
                        objectives.Cells[i + 2, "B"] = "Y";
                    }
                    
                    else if (checkedListBox2.GetItemChecked(i) == false)
                    {
                        objectives.Cells[i + 2, "B"] = "N";
                    }
                }

                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    if (checkedListBox1.GetItemChecked(i) == true)
                    {
                        param.Cells[i + 2, "B"] = "Y";
                    }

                    else if (checkedListBox1.GetItemChecked(i) == false)
                    {
                        param.Cells[i + 2, "B"] = "N";
                    }
                }

                if (objCount == 2)
                {
                    paramWb.SaveAs(tPath);
                    paramWb.Close(0);
                    TaskDialog.Show("Warning", "Please draw your model lines to indicate corridor before selecting boundary input!");
                    excel.Quit();
                    this.Close();
                }

                else
                {
                    TaskDialog.Show("Select Objectives", "Please select two objectives before proceeding.");
                    paramWb.Close(0);
                    excel.Quit();
                }


            }

            catch (Exception ex)
            {
                string message = ex.Message;
                TaskDialog.Show("failed", message);
            }

        }

        private void button2_Click(object sender, EventArgs e) //Cancel button
        {
            TaskDialog.Show("Cancelled", "User cancelled operation");
            this.Close();
        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }
        private void param_Load(object sender, EventArgs e)
        {

        }


        private void checkedListBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void checkedListBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        
    }
}
