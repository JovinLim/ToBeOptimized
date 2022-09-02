using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.DB.Analysis;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using X = Microsoft.Office.Interop.Excel;

namespace TBO_Plugin 
{
    public class RequestHandler2 : IExternalEventHandler
    {
        private Request m_request = new Request();
        private List<ElementId> elemId_list = new List<ElementId>();
        private List<Line> fireEgress_list = new List<Line>();
        private List<Line> fireEgressl_list = new List<Line>();
        List<List<XYZ>> m_xyz = null;
        List<GroupType> m_groups = null;
        dynamic genInfo = new List<dynamic>();
        private int count = 0;
        private int iteration = 0;

        public Request Request
        {
            get { return m_request; }
        }

        public String GetName()
        {
            return "R2022 External Event Handler";
        }

        public void Execute(UIApplication uiapp)
        {

            //MessageBox.Show(count.ToString());
            if (count == 0)
            {
                //MessageBox.Show("The Count is Currently " + count.ToString());
                m_xyz = matrixGen();
                //MessageBox.Show("In Execute " + m_xyz.Count.ToString());
                //MessageBox.Show("Running GetModelGroups()");
                m_groups = GetModelGroups(uiapp);
                //MessageBox.Show(m_groups.Count.ToString());
            }
            try
            {
                switch (Request.Take())
                {
                    case RequestId.None:
                        {

                            return;  // no request at this time -> we can leave immediately
                        }
                    case RequestId.Iteration_1:
                        {
                            //MessageBox.Show("Option 1 - Deleting Model Groups");
                            elemId_list = DeleteModelGroups(uiapp, elemId_list);
                            //MessageBox.Show("Option 1 - Getting Group Indexes");
                            Tuple<List<List<float>>, List<Tuple<List<float>, List<float>>>> groupIndex = getGroupIndex(1);
                            //MessageBox.Show("Option 1 - Getting Cluster Types");
                            Tuple<List<GroupType>, List<string>> clusterTypes = GetClusterTypes(1, m_groups);
                            //MessageBox.Show("Option 1 - Getting Cluster Transforms");
                            List<int> transform = GetTransform(1);
                            //MessageBox.Show("Option 1 - Door Orientation");
                            List<int> door_orient = DoorOrientation(1);
                            //MessageBox.Show("Option 1 - Placing Model Groups");
                            elemId_list = loFiGroups(uiapp, m_xyz, clusterTypes.Item1, elemId_list, groupIndex, transform, door_orient);
                            genInfo = new List<dynamic> { groupIndex, clusterTypes.Item2, transform, door_orient, elemId_list };
                            //elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, groupIndex, transform);
                            iteration = 1;
                            break;
                        }
                    case RequestId.Iteration_2:
                        {
                            //MessageBox.Show("Option 2");
                            elemId_list = DeleteModelGroups(uiapp, elemId_list);
                            Tuple<List<List<float>>, List<Tuple<List<float>, List<float>>>> groupIndex = getGroupIndex(2);
                            Tuple<List<GroupType>, List<string>> clusterTypes = GetClusterTypes(2, m_groups);
                            List<int> transform = GetTransform(2);
                            List<int> door_orient = DoorOrientation(2);
                            elemId_list = loFiGroups(uiapp, m_xyz, clusterTypes.Item1, elemId_list, groupIndex, transform, door_orient);
                            genInfo = new List<dynamic> { groupIndex, clusterTypes.Item2, transform, door_orient, elemId_list };
                            //elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, groupIndex, transform);
                            iteration = 2;
                            break;
                        }
                    case RequestId.Iteration_3:
                        {
                            //MessageBox.Show("Option 3");
                            elemId_list = DeleteModelGroups(uiapp, elemId_list);
                            Tuple<List<List<float>>, List<Tuple<List<float>, List<float>>>> groupIndex = getGroupIndex(3);
                            Tuple<List<GroupType>, List<string>> clusterTypes = GetClusterTypes(3, m_groups);
                            List<int> transform = GetTransform(3);
                            List<int> door_orient = DoorOrientation(3);
                            elemId_list = loFiGroups(uiapp, m_xyz, clusterTypes.Item1, elemId_list, groupIndex, transform, door_orient);
                            genInfo = new List<dynamic> { groupIndex, clusterTypes.Item2, transform, door_orient, elemId_list };
                            //elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, groupIndex, transform);
                            iteration = 3;
                            break;
                        }
                    case RequestId.Iteration_4:
                        {
                            //MessageBox.Show("Option 4");
                            elemId_list = DeleteModelGroups(uiapp, elemId_list);
                            Tuple<List<List<float>>, List<Tuple<List<float>, List<float>>>> groupIndex = getGroupIndex(4);
                            Tuple<List<GroupType>, List<string>> clusterTypes = GetClusterTypes(4, m_groups);
                            List<int> transform = GetTransform(4);
                            List<int> door_orient = DoorOrientation(4);
                            elemId_list = loFiGroups(uiapp, m_xyz, clusterTypes.Item1, elemId_list, groupIndex, transform, door_orient);
                            genInfo = new List<dynamic> { groupIndex, clusterTypes.Item2, transform, door_orient, elemId_list };
                            //elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, groupIndex, transform);
                            iteration = 4;
                            break;
                        }

                    case RequestId.Iteration_5:
                        {
                            //MessageBox.Show("Option 5");
                            elemId_list = DeleteModelGroups(uiapp, elemId_list);
                            Tuple<List<List<float>>, List<Tuple<List<float>, List<float>>>> groupIndex = getGroupIndex(5);
                            Tuple<List<GroupType>, List<string>> clusterTypes = GetClusterTypes(5, m_groups);
                            List<int> transform = GetTransform(5);
                            List<int> door_orient = DoorOrientation(5);
                            elemId_list = loFiGroups(uiapp, m_xyz, clusterTypes.Item1, elemId_list, groupIndex, transform, door_orient);
                            genInfo = new List<dynamic> { groupIndex, clusterTypes.Item2, transform, door_orient, elemId_list };
                            //elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, groupIndex, transform);
                            iteration = 5;
                            break;
                        }
                    case RequestId.Iteration_6:
                        {
                            //MessageBox.Show("Option 6");
                            elemId_list = DeleteModelGroups(uiapp, elemId_list);
                            Tuple<List<List<float>>, List<Tuple<List<float>, List<float>>>> groupIndex = getGroupIndex(6);
                            Tuple<List<GroupType>, List<string>> clusterTypes = GetClusterTypes(6, m_groups);
                            List<int> transform = GetTransform(6);
                            List<int> door_orient = DoorOrientation(6);
                            elemId_list = loFiGroups(uiapp, m_xyz, clusterTypes.Item1, elemId_list, groupIndex, transform, door_orient);
                            genInfo = new List<dynamic> { groupIndex, clusterTypes.Item2, transform, door_orient, elemId_list };
                            //elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, groupIndex, transform);
                            iteration = 6;
                            break;
                        }
                    case RequestId.Iteration_7:
                        {
                            //MessageBox.Show("Option 7");
                            elemId_list = DeleteModelGroups(uiapp, elemId_list);
                            Tuple<List<List<float>>, List<Tuple<List<float>, List<float>>>> groupIndex = getGroupIndex(7);
                            Tuple<List<GroupType>, List<string>> clusterTypes = GetClusterTypes(7, m_groups);
                            List<int> transform = GetTransform(7);
                            List<int> door_orient = DoorOrientation(7);
                            elemId_list = loFiGroups(uiapp, m_xyz, clusterTypes.Item1, elemId_list, groupIndex, transform, door_orient);
                            genInfo = new List<dynamic> { groupIndex, clusterTypes.Item2, transform, door_orient, elemId_list };
                            //elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, groupIndex, transform);
                            iteration = 7;
                            break;
                        }
                    case RequestId.Iteration_8:
                        {
                            //MessageBox.Show("Option 8");
                            elemId_list = DeleteModelGroups(uiapp, elemId_list);
                            Tuple<List<List<float>>, List<Tuple<List<float>, List<float>>>> groupIndex = getGroupIndex(8);
                            Tuple<List<GroupType>, List<string>> clusterTypes = GetClusterTypes(8, m_groups);
                            List<int> transform = GetTransform(8);
                            List<int> door_orient = DoorOrientation(8);
                            elemId_list = loFiGroups(uiapp, m_xyz, clusterTypes.Item1, elemId_list, groupIndex, transform, door_orient);
                            genInfo = new List<dynamic> { groupIndex, clusterTypes.Item2, transform, door_orient, elemId_list };
                            //elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, groupIndex, transform);
                            iteration = 8;
                            break;
                        }
                    case RequestId.Iteration_9:
                        {
                            //MessageBox.Show("Option 9");
                            elemId_list = DeleteModelGroups(uiapp, elemId_list);
                            Tuple<List<List<float>>, List<Tuple<List<float>, List<float>>>> groupIndex = getGroupIndex(9);
                            Tuple<List<GroupType>, List<string>> clusterTypes = GetClusterTypes(9, m_groups);
                            List<int> transform = GetTransform(9);
                            List<int> door_orient = DoorOrientation(9);
                            elemId_list = loFiGroups(uiapp, m_xyz, clusterTypes.Item1, elemId_list, groupIndex, transform, door_orient);
                            genInfo = new List<dynamic> { groupIndex, clusterTypes.Item2, transform, door_orient, elemId_list };
                            //elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, groupIndex, transform);
                            iteration = 9;
                            break;
                        }
                    case RequestId.Iteration_10:
                        {
                            //MessageBox.Show("Option 10");
                            elemId_list = DeleteModelGroups(uiapp, elemId_list);
                            Tuple<List<List<float>>, List<Tuple<List<float>, List<float>>>> groupIndex = getGroupIndex(10);
                            Tuple<List<GroupType>, List<string>> clusterTypes = GetClusterTypes(10, m_groups);
                            List<int> transform = GetTransform(10);
                            List<int> door_orient = DoorOrientation(10);
                            elemId_list = loFiGroups(uiapp, m_xyz, clusterTypes.Item1, elemId_list, groupIndex, transform, door_orient);
                            genInfo = new List<dynamic> { groupIndex, clusterTypes.Item2, transform, door_orient, elemId_list };
                            //elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, groupIndex, transform);
                            iteration = 10;
                            break;
                        }
                    case RequestId.Iteration_11:
                        {
                            //MessageBox.Show("Option 11");
                            elemId_list = DeleteModelGroups(uiapp, elemId_list);
                            Tuple<List<List<float>>, List<Tuple<List<float>, List<float>>>> groupIndex = getGroupIndex(11);
                            Tuple<List<GroupType>, List<string>> clusterTypes = GetClusterTypes(11, m_groups);
                            List<int> transform = GetTransform(11);
                            List<int> door_orient = DoorOrientation(11);
                            elemId_list = loFiGroups(uiapp, m_xyz, clusterTypes.Item1, elemId_list, groupIndex, transform, door_orient);
                            genInfo = new List<dynamic> { groupIndex, clusterTypes.Item2, transform, door_orient, elemId_list };
                            //elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, groupIndex, transform);
                            iteration = 11;
                            break;
                        }
                    case RequestId.Iteration_12:
                        {
                            //MessageBox.Show("Option 12");
                            elemId_list = DeleteModelGroups(uiapp, elemId_list);
                            Tuple<List<List<float>>, List<Tuple<List<float>, List<float>>>> groupIndex = getGroupIndex(12);
                            Tuple<List<GroupType>, List<string>> clusterTypes = GetClusterTypes(12, m_groups);
                            List<int> transform = GetTransform(12);
                            List<int> door_orient = DoorOrientation(12);
                            elemId_list = loFiGroups(uiapp, m_xyz, clusterTypes.Item1, elemId_list, groupIndex, transform, door_orient);
                            genInfo = new List<dynamic> { groupIndex, clusterTypes.Item2, transform, door_orient, elemId_list };
                            //elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, groupIndex, transform);
                            iteration = 12;
                            break;
                        }
                    case RequestId.Apply:
                        {
                            //MessageBox.Show("Apply");
                            elemId_list = DeleteModelGroups(uiapp, elemId_list);
                            //MessageBox.Show("Placing hi def model");
                            elemId_list = PlaceModelGroup(uiapp, m_xyz, genInfo, m_groups);
                            break;
                        }
                    case RequestId.Cancel:
                        {
                            elemId_list = DeleteModelGroups(uiapp, elemId_list);
                            break;
                        }

                    case RequestId.FireEgress:
                        {
                            MessageBox.Show("Showing Fire Egress");
                            
                            fireEgress_list = FireEgressPaths(uiapp, m_xyz, elemId_list);
                            break;
                        }

                    case RequestId.HideFE:
                        {
                            MessageBox.Show("Hiding Fire Egress");
                            DeleteLines(uiapp);
                            break;
                        }

                    default:
                        {
                            break;
                        }
                }
            }
            finally
            {
                count++;
            }
            return;
        }

        private void DeleteLines(UIApplication uiapp)
        {
            Document doc = uiapp.ActiveUIDocument.Document;
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> elem_groups = collector.OfCategory(BuiltInCategory.OST_PathOfTravelLines).ToElements();
            List<ElementId> ids = new List<ElementId>();
            Transaction trans = new Transaction(doc);
            try
            {
                trans.Start("Delete Lines");
                foreach (Element ele in elem_groups)
                {
                    doc.Delete(ele.Id);
                }
                trans.Commit();
            }

            catch (Exception ex)
            {
                string message = ex.Message;
                MessageBox.Show(message);
            }
        }

        private List<ElementId> DeleteModelGroups(UIApplication uiapp, List<ElementId> list)
        {
            Document doc = uiapp.ActiveUIDocument.Document;
            Transaction trans = new Transaction(doc);
            trans.Start("Delete Model Groups");
            if (list.Count != 0)
            {
                try
                {
                    foreach (ElementId elemid in list)
                    {
                        doc.Delete(elemid);
                    }
                }
                catch (Exception ex)
                {
                    string message = ex.Message;
                    MessageBox.Show(message);
                }

            }
            trans.Commit();
            list.Clear();
            return list;
        }
        private List<GroupType> GetModelGroups(UIApplication uiapp)
        {
            List<string> clusterNames = new List<string> {
                "CE1L",
                "CDL",
                "CE2L",
                "IH",
                "IL",
                "RH",
                "RL",
                "SH",
                "SL",
                "TH",
                "TL",
                "VH",
                "VL",
                "WBL",
                "WSL",
                "CDH",
                "CE1H",
                "CE2H",
                "WBH",
                "WSH",
                "WSXL",
                "WSXH",
            };


            Document doc = uiapp.ActiveUIDocument.Document;
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            ICollection<Element> elem_groups = collector.OfCategory(BuiltInCategory.OST_IOSModelGroups).ToElements();
            List<GroupType> all_groups = new List<GroupType>();
            List<GroupType> m_groups = null;
            try
            {
                foreach (Element grouptype in elem_groups)
                {
                    if (!all_groups.Contains(grouptype))
                    {
                        if (grouptype.Category.Name == "Model Groups")
                        {
                            if (clusterNames.Contains(grouptype.Name))
                            {
                                GroupType gt = grouptype as GroupType;
                                all_groups.Add(gt);
                            }
                        }
                    }


                }
                m_groups = all_groups.Distinct().ToList();
                for (int i = 0; i < m_groups.Count; i++)
                {
                    if (m_groups[i] == null)
                    {
                        m_groups.RemoveAt(i);
                    }
                }
                //MessageBox.Show(m_groups.Count.ToString());
            }

            catch (Exception ex)
            {
                string message = ex.Message;
                MessageBox.Show(message);
            }
            return m_groups;
        }


        private List<int> GetTransform(int iteration_num)
        {
            string oPath = @"C:\temp\Output.xlsx";
            X.Application excel_python = new X.Application();
            X.Workbook wb_python = excel_python.Workbooks.Open(oPath);
            X._Worksheet ws_python = (X._Worksheet)wb_python.Sheets["Sheet1"];
            Microsoft.Office.Interop.Excel.Range strRange = (Microsoft.Office.Interop.Excel.Range)ws_python.Cells[iteration_num + 1, "G"];
            string full = strRange.Value2;
            string parse_1 = full.Replace("[", "");
            string parse_2 = parse_1.Replace("]", "");
            string[] allDigits = parse_2.Split(',');
            List<int> transform_int = new List<int>();
            foreach (string digit in allDigits)
            {
                transform_int.Add(0);
                //transform_int.Add(Convert.ToInt32(digit));
            }
            wb_python.Close(0);
            excel_python.Quit();
            return transform_int;

        }

        private List<int> DoorOrientation(int iteration_num)
        {
            string oPath = @"C:\temp\Output.xlsx";
            X.Application excel_python = new X.Application();
            X.Workbook wb_python = excel_python.Workbooks.Open(oPath);
            X._Worksheet ws_python = (X._Worksheet)wb_python.Sheets["Sheet1"];
            Microsoft.Office.Interop.Excel.Range strRange = (Microsoft.Office.Interop.Excel.Range)ws_python.Cells[iteration_num + 1, "I"];
            string full = strRange.Value2;
            string parse_1 = full.Replace("[", "");
            string parse_2 = parse_1.Replace("]", "");
            string[] allDigits = parse_2.Split(',');
            List<int> transform_int = new List<int>();
            foreach (string digit in allDigits)
            {
                transform_int.Add(Convert.ToInt32(digit));
            }
            wb_python.Close(0);
            excel_python.Quit();
            return transform_int;

        }
        private Tuple<List<List<float>>, List<Tuple<List<float>, List<float>>>> getGroupIndex(int iteration_num)
        {
            string oPath = @"C:\temp\Output.xlsx";
            X.Application excel_python = new X.Application();
            X.Workbook wb_python = excel_python.Workbooks.Open(oPath);
            X._Worksheet ws_python = (X._Worksheet)wb_python.Sheets["Sheet1"];
            Microsoft.Office.Interop.Excel.Range strRange = (Microsoft.Office.Interop.Excel.Range)ws_python.Cells[iteration_num + 1, "D"];
            string full = strRange.Value2;
            string parse1 = full.Remove(full.Length - 1);
            string[] allDigits = parse1.Split('|');
            List<string> singleDigits = new List<string>();
            List<List<float>> groupIndex = new List<List<float>>();
            try
            {
                foreach (string s in allDigits)
                {
                    List<float> sIndex = new List<float>();
                    string[] index = s.Split(',');
                    foreach (string digit in index)
                    {
                        sIndex.Add(float.Parse((digit)));
                    }
                    groupIndex.Add(sIndex);
                }
            }
            catch (Exception ex)
            {
                string message = ex.Message;
                MessageBox.Show(message);
            }
            Microsoft.Office.Interop.Excel.Range wallstrRange = (Microsoft.Office.Interop.Excel.Range)ws_python.Cells[iteration_num + 1, "J"];
            string wallFull = wallstrRange.Value2;
            string wallParse1 = wallFull.Remove(wallFull.Length - 1);
            string[] wallSplit1 = wallParse1.Split('|');
            List<string> wallSingleDigits = new List<string>();
            List<Tuple<List<float>,List<float>>> wallIndex = new List<Tuple<List<float>, List<float>>>();
            try
            {
                foreach (string s in wallSplit1)
                {
                    List<List<float>> wIndex = new List<List<float>>();
                    string[] wallSplit2 = s.Split('/');
                    foreach (string digit in wallSplit2)
                    {
                        List<float> wwIndex = new List<float>();
                        string[] wallSplit3 = digit.Split(',');
                        foreach (string digit2 in wallSplit3)
                        {
                            wwIndex.Add(float.Parse(digit2));
                        }
                        wIndex.Add(wwIndex);
                    }
                    Tuple<List<float>, List<float>> tup = Tuple.Create(wIndex[0], wIndex[1]);
                    wallIndex.Add(tup);
                }
            }
            catch (Exception ex)
            {
                string message = ex.Message;
                MessageBox.Show(message);
            }
            wb_python.Close(0);
            excel_python.Quit();
            return Tuple.Create(groupIndex, wallIndex);
        }

        private static Tuple<List<GroupType>, List<string>> GetClusterTypes(int iteration_num, List<GroupType> m_groups)
        {
            List<GroupType> clusters = new List<GroupType>();
            List<String> clusterNames = new List<String>();
            try
            {
                string oPath = @"C:\temp\Output.xlsx";
                X.Application excel_python = new X.Application();
                X.Workbook wb_python = excel_python.Workbooks.Open(oPath);
                X._Worksheet ws_python = (X._Worksheet)wb_python.Sheets["Sheet1"];
                Microsoft.Office.Interop.Excel.Range strRange = (Microsoft.Office.Interop.Excel.Range)ws_python.Cells[iteration_num + 1, "F"];
                string full = strRange.Value;
                string parse1 = full.Remove(full.Length - 1);
                string[] parse2 = parse1.Split(',');
                foreach (string s in parse2)
                {
                    clusterNames.Add(s); 
                }
                List<string> groupNames = new List<string>();
                List<string> count = new List<string>();
                foreach (GroupType groupType in m_groups)
                {
                    groupNames.Add(groupType.Name);
                }
                foreach (string s in clusterNames)
                {
                    //MessageBox.Show(s);
                    int elem_index = groupNames.IndexOf(s);
                    //MessageBox.Show(elem_index.ToString());
                    clusters.Add(m_groups[elem_index]);
                }
                
                wb_python.Close(0);
                excel_python.Quit();
            }

            catch (Exception ex)
            {
                string message = ex.Message;
                MessageBox.Show(message);
            }

            return Tuple.Create(clusters, clusterNames);
        }

        private List<List<XYZ>> matrixGen()
        {
            //MessageBox.Show("Matrix Gen");
            string mPath = @"C:\temp\SOA_Copy.xlsx";
            X.Application excel_matrix = new X.Application();
            X.Workbook wb_matrix = excel_matrix.Workbooks.Open(mPath);
            X._Worksheet ws_matrix = (X._Worksheet)wb_matrix.Sheets["Orthopaedic"];
            List<List<XYZ>> m_xyz = new List<List<XYZ>>();
            try
            {
                Microsoft.Office.Interop.Excel.Range xRange = (Microsoft.Office.Interop.Excel.Range)ws_matrix.Cells[3, "G"];
                Microsoft.Office.Interop.Excel.Range yRange = (Microsoft.Office.Interop.Excel.Range)ws_matrix.Cells[4, "G"];
                Microsoft.Office.Interop.Excel.Range zRange = (Microsoft.Office.Interop.Excel.Range)ws_matrix.Cells[5, "G"];
                float xVal = (float)xRange.Value;
                float yVal = (float)yRange.Value;
                float zVal = (float)zRange.Value;
                XYZ bottom_left_pt = new XYZ(xVal, yVal, zVal);
                float grid_size = 3.93701f;

                Microsoft.Office.Interop.Excel.Range y_axRange = (Microsoft.Office.Interop.Excel.Range)ws_matrix.Cells[3, "F"];
                Microsoft.Office.Interop.Excel.Range x_axRange = (Microsoft.Office.Interop.Excel.Range)ws_matrix.Cells[4, "F"];
                int y_ax = (int)y_axRange.Value;
                int x_ax = (int)x_axRange.Value;
                try
                {
                    for (int y = 0; y < y_ax; y++)
                    {
                        var y_id = new List<XYZ>();
                        for (int x = 0; x < x_ax; x++)
                        {
                            XYZ newpt = new XYZ(((float)bottom_left_pt.X + x * grid_size), ((float)bottom_left_pt.Y + y * grid_size), 0);
                            y_id.Add(newpt);
                        }
                        m_xyz.Add(y_id);
                    }
                    wb_matrix.Close(0);
                    excel_matrix.Quit();
                }

                catch (Exception ex)
                {
                    string message = ex.Message;
                    MessageBox.Show(message);
                }

            }

            catch (Exception ex)
            {
                string message = ex.Message;
                MessageBox.Show(message);
            }
            return m_xyz;

        }

        private List<ElementId> PlaceModelGroup(UIApplication uiapp, List<List<XYZ>> m_xyz, dynamic info, List<GroupType> m_groups)
        {
            
            Document doc = uiapp.ActiveUIDocument.Document;
            Autodesk.Revit.DB.View active = doc.ActiveView;
            Transaction trans = new Transaction(doc);

            try
            {
                List<ElementId> groups = info[4];
                List<ElementId> newGroups = new List<ElementId>();
                List<List<float>> gIndex = info[0].Item1;
                List<Tuple<List<float>, List<float>>> wallIndex = info[0].Item2;
                List<string> clustTypes = info[1];
                List<int> transform = info[2];
                List<int> door_dir = info[3];
                List<string> clusterH = new List<string>();
                List<GroupType> clusters = new List<GroupType>();
                float grid_size = 3.93701f;
                foreach (string s in clustTypes)
                {
                    string parse = s.Replace('L', 'H');
                    clusterH.Add(parse);
                }
                List<string> groupNames = new List<string>();
                foreach (GroupType groupType in m_groups)
                {
                    groupNames.Add(groupType.Name);
                }
                foreach (string s in clusterH)
                {
                    int elem_index = groupNames.IndexOf(s);
                    clusters.Add(m_groups[elem_index]);
                }
                try
                {
                    ElementId level_id = new ElementId(311);
                    Level lv1 = doc.GetElement(level_id) as Level;
                    FilteredElementCollector collector = new FilteredElementCollector(doc);
                    FamilySymbol doorSymbol = GetFirstSymbol(FindDoorFamilies(doc).FirstOrDefault(), doc);
                    WallType wType = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().FirstOrDefault(q => q.Name == "Generic - 100mm");
                    ElementId wId = wType.Id;

                    trans.Start("High F Group Placement");
                    doorSymbol.Activate();
                    List<Wall> walls = new List<Wall>();

                    if (door_dir[0] == 0)
                    {
                        foreach (Tuple<List<float>, List<float>> t in wallIndex)
                        {
                            XYZ start_pt = m_xyz[0][0];
                            XYZ pt1 = new XYZ(start_pt.X + (t.Item1[1] * grid_size), start_pt.Y + (t.Item1[0] * grid_size) + (0.05 * grid_size), start_pt.Z);
                            XYZ pt2 = new XYZ(start_pt.X + (t.Item2[1] * grid_size), start_pt.Y + (t.Item2[0] * grid_size) + (0.05 * grid_size), start_pt.Z);
                            Line wallLine = Line.CreateBound(pt1, pt2);
                            List<Curve> wallCrv = new List<Curve> { wallLine as Curve };
                            Wall wall = Wall.Create(doc, wallLine, wId, level_id, 3 * grid_size, 0.0, false, false);
                            ElementId wallId = wall.Id;
                            newGroups.Add(wallId);
                        }
                    }

                    else
                    {
                        foreach (Tuple<List<float>, List<float>> t in wallIndex)
                        {
                            XYZ start_pt = m_xyz[0][0];
                            XYZ pt1 = new XYZ(start_pt.X + (t.Item1[1] * grid_size), start_pt.Y + (t.Item1[0] * grid_size) - (0.05 * grid_size), start_pt.Z);
                            XYZ pt2 = new XYZ(start_pt.X + (t.Item2[1] * grid_size), start_pt.Y + (t.Item2[0] * grid_size) - (0.05 * grid_size), start_pt.Z);
                            Line wallLine = Line.CreateBound(pt1, pt2);
                            List<Curve> wallCrv = new List<Curve> { wallLine as Curve };
                            Wall wall = Wall.Create(doc, wallLine, wId, level_id, 3 * grid_size, 0.0, false, false);
                            ElementId wallId = wall.Id;
                            newGroups.Add(wallId);
                        }
                    }
                    for (int i = 0; i < gIndex.Count(); i++)
                    {
                        List<float> index = gIndex[i];

                        XYZ start_pt = m_xyz[0][0];
                        XYZ pt = new XYZ(start_pt.X + (index[1] * grid_size), start_pt.Y + (index[0] * grid_size), start_pt.Z);
                        Group place = doc.Create.PlaceGroup(pt, clusters[i]);
                        List<ElementId> place_id = new List<ElementId> { place.Id };
                        if (door_dir[i] == 0)
                        {
                            XYZ pt1 = new XYZ(pt.X, pt.Y, pt.Z + 1);
                            Line axis = Line.CreateBound(pt, pt1);
                            ElementTransformUtils.RotateElements(doc, place_id, axis, Math.PI);
                        }

                        foreach (ElementId id in place_id)
                        {
                            newGroups.Add(id);
                        }
                    }
                    trans.Commit();
                    return newGroups;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    return null;
                }
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        private List<ElementId> loFiGroups(UIApplication uiapp, List<List<XYZ>> m_xyz, List<GroupType> clustTypes, List<ElementId> asd,
            Tuple<List<List<float>>, List<Tuple<List<float>, List<float>>>> groupIndex, List<int> transform, List<int> door_dir)
        {
            //MessageBox.Show("Generating Floorplan...");
            Document doc = uiapp.ActiveUIDocument.Document;
            Autodesk.Revit.DB.View active = doc.ActiveView;
            Transaction trans = new Transaction(doc);
            ElementId level_id = new ElementId(311);
            Level lv1 = doc.GetElement(level_id) as Level;
            float grid_size = 3.93701f;
            List<List<float>> group_list = groupIndex.Item1;
            List<Tuple<List<float>, List<float>>> wallIndex = groupIndex.Item2;
            FilteredElementCollector collector = new FilteredElementCollector(doc);
            FamilySymbol doorSymbol = GetFirstSymbol(FindDoorFamilies(doc).FirstOrDefault(), doc);
            WallType wType = new FilteredElementCollector(doc).OfClass(typeof(WallType)).Cast<WallType>().FirstOrDefault(q=> q.Name == "Generic - 100mm");
            ElementId wId = wType.Id;
            List<ElementId> new_elemId = new List<ElementId>();
            try
            {
                trans.Start("Model Group Placement");
                doorSymbol.Activate();
                List<Wall> walls = new List<Wall>();
                if (door_dir[0] == 0)
                {
                    foreach (Tuple<List<float>, List<float>> t in wallIndex)
                    {
                        XYZ start_pt = m_xyz[0][0];
                        XYZ pt1 = new XYZ(start_pt.X + (t.Item1[1] * grid_size), start_pt.Y + (t.Item1[0] * grid_size) + (0.05 * grid_size), start_pt.Z);
                        XYZ pt2 = new XYZ(start_pt.X + (t.Item2[1] * grid_size), start_pt.Y + (t.Item2[0] * grid_size) + (0.05 * grid_size), start_pt.Z);
                        Line wallLine = Line.CreateBound(pt1, pt2);
                        List<Curve> wallCrv = new List<Curve> { wallLine as Curve };
                        Wall wall = Wall.Create(doc, wallLine, wId, level_id, 3 * grid_size, 0.0, false, false);
                        ElementId wallId = wall.Id;
                        new_elemId.Add(wallId);
                    }
                }

                else
                {
                    foreach (Tuple<List<float>, List<float>> t in wallIndex)
                    {
                        XYZ start_pt = m_xyz[0][0];
                        XYZ pt1 = new XYZ(start_pt.X + (t.Item1[1] * grid_size), start_pt.Y + (t.Item1[0] * grid_size) - (0.075 * grid_size), start_pt.Z);
                        XYZ pt2 = new XYZ(start_pt.X + (t.Item2[1] * grid_size), start_pt.Y + (t.Item2[0] * grid_size) - (0.075 * grid_size), start_pt.Z);
                        Line wallLine = Line.CreateBound(pt1, pt2);
                        List<Curve> wallCrv = new List<Curve> { wallLine as Curve };
                        Wall wall = Wall.Create(doc, wallLine, wId, level_id, 3 * grid_size, 0.0, false, false);
                        ElementId wallId = wall.Id;
                        new_elemId.Add(wallId);
                    }
                }

                for (int i = 0; i < group_list.Count(); i++)
                {
                    List<float> index = group_list[i];
                    XYZ start_pt = m_xyz[0][0];
                    XYZ pt = new XYZ(start_pt.X + (index[1] * grid_size), start_pt.Y + (index[0] * grid_size), start_pt.Z);
                    Group place = doc.Create.PlaceGroup(pt, clustTypes[i]);
                    List<ElementId> place_id = new List<ElementId> { place.Id };
                    //if (transform[i] == 1)
                    //{
                    //    Plane plane = Plane.CreateByThreePoints(pt, new XYZ(pt.X + 1, pt.Y, pt.Z), new XYZ(pt.X, pt.Y, pt.Z + 1));
                    //    ElementTransformUtils.MirrorElements(doc, place_id, plane, false);
                    //}
                    //else if (transform[i] == 2)
                    //{
                    //    XYZ pt1 = new XYZ(pt.X, pt.Y + 1, pt.Z);
                    //    Plane plane = Plane.CreateByThreePoints(pt, new XYZ(pt.X, pt.Y + 1, pt.Z), new XYZ(pt.X, pt.Y, pt.Z + 1));
                    //    ElementTransformUtils.MirrorElements(doc, place_id, plane, false);
                    //}

                    //else if (transform[i] == 3)
                    //{
                    //    XYZ pt1 = new XYZ(pt.X, pt.Y, pt.Z + 1);
                    //    Line axis = Line.CreateBound(pt, pt1);
                    //    ElementTransformUtils.RotateElements(doc, place_id, axis, Math.PI);
                    //}

                    //Rotating model groups for correct door orientation
                    if (door_dir[i] == 0)
                    {
                        XYZ pt1 = new XYZ(pt.X, pt.Y, pt.Z + 1);
                        Line axis = Line.CreateBound(pt, pt1);
                        ElementTransformUtils.RotateElements(doc, place_id, axis, Math.PI);
                    }
                    new_elemId.Add(place.Id);
                }



                trans.Commit();
                return new_elemId;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }

        }


        public List<Line> FireEgressPaths(UIApplication uiapp, List<List<XYZ>> m_xyz, List<ElementId> elem_list)
        {
            try
            {
                //Instantiating document variables
                Autodesk.Revit.DB.Document doc = uiapp.ActiveUIDocument.Document;
                Selection sel = uiapp.ActiveUIDocument.Selection;
                Transaction trans = new Transaction(doc);
                DoorPickFilter doorFilter = new DoorPickFilter();
                float grid_size = 3.93701f;

                //Selecting egress point
                TaskDialog.Show("Define Fire Escape", "Please select a fire escape door.");
                Reference pickedegress = sel.PickObject(ObjectType.Element, doorFilter, "Select fire escape door."); //Calling method to prompt for door input
                Element egress = doc.GetElement(pickedegress);
                Location egress_loc = egress.Location;
                LocationPoint egress_lp = egress_loc as LocationPoint;
                XYZ ePt = egress_lp.Point;
                List<XYZ> ePts = new List<XYZ> { ePt };
                List<Element> doorElem = new List<Element>();
                List<XYZ> group_pts = new List<XYZ>();

                //Finding index of closest boundary point to fire egress point
                List<double> boundXYZDist = new List<double>();
                List<int> boundXIndex = new List<int>();
                List<int> boundIndex = new List<int>();

                for (int y = 0; y < m_xyz.Count - 1; y++)
                {
                    List<double> boundXDist = new List<double>();
                    for (int x = 0; x < m_xyz[y].Count - 1; x++)
                    {
                        boundXDist.Add(m_xyz[y][x].DistanceTo(ePt));
                    }
                    boundXYZDist.Add(boundXDist.Min());
                    boundXIndex.Add(boundXDist.IndexOf(boundXDist.Min()));
                }

                boundIndex.Add(boundXYZDist.IndexOf(boundXYZDist.Min()));
                boundIndex.Add(boundXIndex[boundXYZDist.IndexOf(boundXYZDist.Min())]);

                XYZ boundPt = m_xyz[boundIndex[0]][boundIndex[1]];


                foreach (ElementId elemId in elem_list)
                {
                    if (doc.GetElement(elemId).Category.Name == "Model Groups")
                    {
                        LocationPoint mgroup_loc = doc.GetElement(elemId).Location as LocationPoint;
                        group_pts.Add(mgroup_loc.Point);
                    }
                }
                //USING GROUP ORIGIN POINTS
                List<double> distances = new List<double>();
                foreach (XYZ pt in group_pts)
                {

                    double dist = pt.DistanceTo(boundPt);
                    distances.Add(dist);
                }
                List<double> sorted_dist = distances.OrderByDescending(dist => dist).ToList();
                List<XYZ> furthest_group_pt = new List<XYZ>();
                for (int i = 0; i < sorted_dist.Count; i++)
                {
                    int d_ind = distances.IndexOf(sorted_dist[i]);
                    furthest_group_pt.Add(group_pts[d_ind]);
                }
                trans.Start("Find Shortest Path");

                //USING GROUP ORIGIN
                PathOfTravel route = PathOfTravel.Create(doc.ActiveView, boundPt, furthest_group_pt[0]);
                PathOfTravel boundRoute = PathOfTravel.Create(doc.ActiveView, boundPt, ePt);
                trans.Commit();
                IList<Curve> routeCrv = route.GetCurves();
                IList<Curve> boundrouteCrv = boundRoute.GetCurves();
                List<ElementId> crvListID = new List<ElementId>();
                List<ElementId> crvID_remove = new List<ElementId>();
                List<ElementId> crvID_keep = new List<ElementId>();
                List<Line> lList = new List<Line>();
                double POTLength = 0;
                foreach (Curve c in routeCrv)
                {
                    POTLength = POTLength + c.Length;
                    Line l = c as Line;
                    lList.Add(l);
                }
                foreach (Curve c in boundrouteCrv)
                {
                    POTLength = POTLength + c.Length;
                    Line l = c as Line;
                    lList.Add(l);
                }
                string POTLengthStr = (POTLength/grid_size).ToString() + "m";
                TaskDialog.Show("Furthest Distance", "Furthest travel distance from fire egress is " + POTLengthStr);
                return lList;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        private void Corridor_EdgeIndexes(int itr)
        {
            string tPath = @"C:\temp\Output.xlsx";
            X.Application excel_Output = new X.Application();
            X.Workbook wb_Output = excel_Output.Workbooks.Open(tPath);
            X._Worksheet ws_Output = (X._Worksheet)wb_Output.Sheets["Sheet1"];
            Microsoft.Office.Interop.Excel.Range xRange = (Microsoft.Office.Interop.Excel.Range)ws_Output.Cells[itr, "I"];
            string indexes = xRange.Value2;


            string nPath = @"C:\temp\Corridor_Edge.xlsx";
            if (!File.Exists(nPath))
            {
                X.Application excel_CorridorEdge = new X.Application();
                X.Workbook wb_CorridorEdge = excel_CorridorEdge.Workbooks.Add();
                X.Worksheet main = (X.Worksheet)wb_CorridorEdge.Worksheets.Add();
                main.Name = "Indexes";
            }

            else if (File.Exists(nPath))
            {
                X.Application excel_CorridorEdge = new X.Application();
                X.Workbook wb_CorridorEdge = excel_CorridorEdge.Workbooks.Open(nPath);
                X.Worksheet ws_CorridorEdge = (X.Worksheet)wb_CorridorEdge.Sheets["Indexes"];
                for (int i = 0; i < 6; i++)
                {
                    Microsoft.Office.Interop.Excel.Range xVal = (Microsoft.Office.Interop.Excel.Range)ws_CorridorEdge.Cells[1, i+1];
                    if (xVal.Value2 == null)
                    {
                        ws_CorridorEdge.Cells[1, i] = indexes;
                        break;
                    }
                }

            }

            wb_Output.Close(0);
            excel_Output.Quit();

        }
        private static IEnumerable<Family> FindDoorFamilies(Document doc)
        {
            return new FilteredElementCollector(doc)
                .OfClass(typeof(Family))
                .Cast<Family>()
                .Where(e => e.FamilyCategory != null
                        && e.FamilyCategory.Id.IntegerValue == (int)BuiltInCategory.OST_Doors);
        }

        private static FamilySymbol GetFirstSymbol(Family family, Document doc)
        {
            ISet<ElementId> famIds = family.GetFamilySymbolIds();
            ElementId famSymId = famIds.FirstOrDefault();
            FamilySymbol famSymbol = doc.GetElement(famSymId) as FamilySymbol;
            return famSymbol;
        }

        //Creating filter to constrain the picking of elements to doors
        public class DoorPickFilter : ISelectionFilter
        {
            public bool AllowElement(Element e)
            {
                return (e.Category.Name == "Doors"); //Defining element category name to allow for selection

            }
            public bool AllowReference(Reference refer, XYZ point)
            {
                return false;
            }
        }

    }


}