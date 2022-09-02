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
    public partial class LayoutGencs3 : System.Windows.Forms.Form
    {
        private RequestHandler2 m_Handler;
        private ExternalEvent m_ExEvent;
        Document Doc;

        string tDir = @"c:\temp";
        string pPath = @"c:\temp\param.xlsx";
        string iPath = @"c:\temp\Output.xlsx";

        public LayoutGencs3(Document doc, ExternalEvent exEvent, RequestHandler2 handler)
        {
            /*  MessageBox.Show("layoutgen");*/
            InitializeComponent();
            LoadExcelSheet(@"c:\temp", 1);
            Doc = doc;
            /* label1.Text = "test";*/
            m_Handler = handler;
            m_ExEvent = exEvent;
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
                if (!File.Exists(pPath))
                {
                    TaskDialog.Show("Error", "No Parameters / Objectives Selected.");
                    this.Close();
                }

                if (!File.Exists(iPath))
                {
                    TaskDialog.Show("Error", "No SOA Input.");
                    this.Close();
                }

                X.Application excel = new X.Application();
                excel.DisplayAlerts = false;
                X.Workbook paramWb = excel.Workbooks.Open(pPath);
                X._Worksheet param = (X._Worksheet)paramWb.Sheets["Parameters"];
                X._Worksheet objectives = (X._Worksheet)paramWb.Sheets["Objectives"];
                Microsoft.Office.Interop.Excel.Range paramRange = (Microsoft.Office.Interop.Excel.Range)param.Range["B1", "B10"];
                Microsoft.Office.Interop.Excel.Range objRangeVal = (Microsoft.Office.Interop.Excel.Range)objectives.Range["B1", "B10"]; //Y or N
                Microsoft.Office.Interop.Excel.Range objRangeName = (Microsoft.Office.Interop.Excel.Range)objectives.Range["A1", "A10"]; //Name of objective
                Microsoft.Office.Interop.Excel.Range objRangeUnit = (Microsoft.Office.Interop.Excel.Range)objectives.Range["C1", "C10"]; //Objective units

                for (int i = 0; i < objRangeVal.Count; i++)
                {
                    if (objRangeVal[i+1].Value2 == "Y")
                    {
                        if (this.label3.Text == "none")
                        {
                            
                            if (objRangeUnit[i + 1].Value2.ToString() == "None")
                            {
                                this.label3.Text = objRangeName[i + 1].Value2.ToString();
                                this.label6.Text = "Objective 1";
                            }

                            else
                            {
                                this.label3.Text = objRangeName[i + 1].Value2.ToString() + " / " + objRangeUnit[i + 1].Value2.ToString();
                                this.label6.Text = "Objective 1 / " + objRangeUnit[i + 1].Value2.ToString();
                            }
                            
                        }
                        
                        else
                        {
                            
                            if (objRangeUnit[i + 1].Value2.ToString() == "None")
                            {
                                this.label4.Text = objRangeName[i + 1].Value2.ToString();
                                this.label7.Text = "Objective 2";
                            }
                            else
                            {
                                this.label4.Text = objRangeName[i + 1].Value2.ToString() + " / " + objRangeUnit[i + 1].Value2.ToString();
                                this.label7.Text = "Objective 2 / " + objRangeUnit[i + 1].Value2.ToString();
                            }
                            
                        }
                    }
                    
                }

                paramWb.Close(0);
                excel.Quit();

            }

            catch (Exception ex)
            {
                string message = ex.Message;
                TaskDialog.Show("Failed", message);
            }
        }

        private void splitContainer1_Panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }

        private void label4_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            MakeRequest(RequestId.Apply);
            //this.Close();
        }

        private void label5_Click(object sender, EventArgs e)
        {

        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void label6_Click(object sender, EventArgs e)
        {

        }

        private void label7_Click(object sender, EventArgs e)
        {

        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

            string sel_index = "Iteration_" + (listBox1.SelectedIndex + 1).ToString();

            if (Enum.IsDefined(typeof(RequestId), sel_index))
            {
                RequestId req_id = (RequestId)Enum.Parse(typeof(RequestId), sel_index, true);
                MakeRequest(req_id);
            }
            else
            {
                MessageBox.Show("Error, RequestId does not exist");
            }
        }

        private void listBox2_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBox3_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBox4_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void LoadExcelSheet(string path, int sheet)
        {
            int row = 0;
            X.Application excel = new X.Application();
            X.Workbook wb = excel.Workbooks.Open(Filename: @"c:\temp\Output.xlsx", UpdateLinks: true, ReadOnly: false, Editable: true, Local: true);
            /*   excel.Visible = true;*/
            X._Worksheet ws = (X._Worksheet)wb.Sheets[1];

            for (row = 2; row < 14; row++)
            {
                string a = "";
                string b = "";
                string c = "";
                string d = "";
                a += ws.Cells[row, 1].Value2 + " ";
                b += ws.Cells[row, 2].Value2 + " ";
                c += ws.Cells[row, 3].Value2 + " ";
                d += ws.Cells[row, 4].Value2 + " ";

                listBox1.Items.Add(a);
                listBox2.Items.Add(b);
                listBox3.Items.Add(c);
            }
            wb.Close();
            excel.Quit();

        }


        /// Form closed event handler
    
        protected override void OnFormClosed(FormClosedEventArgs e)
        {
            // Both the event and the handler
            // Dispose it before we are closed
            m_ExEvent.Dispose();
            m_ExEvent = null;
            m_Handler = null;

            // Call the base class
            base.OnFormClosed(e);
        }

        private void MakeRequest(RequestId request)
        {
            m_Handler.Request.Make(request);
            m_ExEvent.Raise();
            //DozeOff();
        }

  
       

        private void button2_Click(object sender, EventArgs e)
        {
            MakeRequest(RequestId.Cancel);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            checkBox1.ThreeState = false;
            if (checkBox1.Checked == true)
            {
                MakeRequest(RequestId.FireEgress);
                
            }
            
            else
            {
                MakeRequest(RequestId.HideFE);
            }
        }
    }
}
