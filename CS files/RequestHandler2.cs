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
    using Autodesk.Revit.ApplicationServices;
    using Autodesk.Revit.Attributes;
    using Autodesk.Revit.DB;
    using Autodesk.Revit.UI;
    using Autodesk.Revit.UI.Selection;
    using Autodesk.Revit.DB.Architecture;
    using X = Microsoft.Office.Interop.Excel;

    namespace TBO_Plugin
    {
        public class RequestHandler2 : IExternalEventHandler
        {
            private EventWaitHandle pause = new EventWaitHandle(false, EventResetMode.ManualReset);
            private Request m_request = new Request();
            private List<ElementId> elemId_list = new List<ElementId>();
            List<List<XYZ>> m_xyz = null;
            List<GroupType> m_groups = null;
            private int count = 0;

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
                //MessageBox.Show(m_xyz.Count.ToString());
                //MessageBox.Show(m_groups.Count.ToString());
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
                                MessageBox.Show("Option 1");
                            try
                            {
                                //MessageBox.Show("In Iteration_1 " + count.ToString());
                                //MessageBox.Show("In Iteration_1 " + m_xyz.Count.ToString());
                                //MessageBox.Show("In Iteration_1 " + m_groups.Count.ToString());
                            }
                            catch (Exception ex)
                            {
                                string message = ex.Message;
                                MessageBox.Show(message);
                            }
                                MessageBox.Show("Option 1 - Deleting Model Groups");
                                elemId_list = DeleteModelGroups(uiapp, elemId_list);
                                MessageBox.Show("Option 1 - Getting Grid Ids");
                                List<List<int>> GridId = GetGridId(1);
                                MessageBox.Show("Option 1 - Getting Cluster Types");
                                List<GroupType> clusterTypes = GetClusterTypes(1, m_groups);
                                MessageBox.Show("Option 1 - Getting Cluster Transforms");
                                List<int> transform = GetTransform(1);
                                MessageBox.Show("Option 1 - Placing Model Groups");
                                elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, GridId, transform);
                                break;
                            }
                        case RequestId.Iteration_2:
                            {
                                MessageBox.Show("Option 2");
                                elemId_list = DeleteModelGroups(uiapp, elemId_list);
                                List<List<int>> GridId = GetGridId(2);
                                List<GroupType> clusterTypes = GetClusterTypes(2, m_groups);
                                List<int> transform = GetTransform(2);
                                elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, GridId, transform);
                                break;
                            }
                        case RequestId.Iteration_3:
                            {
                                MessageBox.Show("Option 3");
                                elemId_list = DeleteModelGroups(uiapp, elemId_list);
                                List<List<int>> GridId = GetGridId(3);
                                List<GroupType> clusterTypes = GetClusterTypes(3, m_groups);
                                List<int> transform = GetTransform(3);
                                elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, GridId, transform);
                                break;
                            }
                        case RequestId.Iteration_4:
                            {
                                MessageBox.Show("Option 4");
                                elemId_list = DeleteModelGroups(uiapp, elemId_list);
                                List<List<int>> GridId = GetGridId(4);
                                List<GroupType> clusterTypes = GetClusterTypes(4, m_groups);
                                List<int> transform = GetTransform(4);
                                elemId_list = PlaceModelGroup(uiapp, m_xyz, clusterTypes, elemId_list, GridId, transform);
                                break;
                            }
                    default:
                            {
                                // some kind of a warning here should
                                // notify us about an unexpected request 
                                break;
                            }
                        }
                    }
                    finally
                    {
                    //wb_matrix.Close(0);
                    //excel_matrix.Quit();
                    //wb_python.Close(0);
                    //excel_python.Quit();
                    count++;
                    MessageBox.Show("End");
                }
                    return;
                }

            private List<ElementId> DeleteModelGroups(UIApplication uiapp, List<ElementId> list)
            {
                Document doc = uiapp.ActiveUIDocument.Document;
                List<Element> m_groups = new List<Element>();
                //MessageBox.Show("In DeleteModelGroups, before deleting. The list count is " + list.Count.ToString());
                MessageBox.Show("Clearing Floorplan...");
                List<ElementId> delete_list = elemId_list;
                Transaction trans = new Transaction(doc);
                trans.Start("Delete Model Groups");
                if (delete_list.Count != 0)
                {
                    try
                    {
                        foreach (ElementId elemid in elemId_list)
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
            //MessageBox.Show("In DeleteModelGroups, after deleting. The list count is " + list.Count.ToString());
            return list;
            }
            private List<GroupType> GetModelGroups(UIApplication uiapp)
            {
                List<string> clusterNames = new List<string> {
                "CD",
                "CR1",
                "CR1_2",
                "CR2_1",
                "CR2_2",
                "CR2_3",
                "CR3_1",
                "CR3_2",
                "RC1",
                "RC2",
                "RC3",
                "SI",
                "ST_1",
                "ST_2",
                "interviewroom"
            };


                Document doc = uiapp.ActiveUIDocument.Document;
                FilteredElementCollector collector = new FilteredElementCollector(doc);
                ICollection<Element> elem_groups = collector.OfCategory(BuiltInCategory.OST_IOSModelGroups).ToElements();
                List<GroupType> all_groups = new List<GroupType>();

                foreach (GroupType grouptype in elem_groups)
                {
                    if (clusterNames.Contains(grouptype.Name))
                    {
                        all_groups.Add(grouptype);

                    }
                }
                return all_groups;
            }


            private List<int> GetTransform(int iteration_num)
            {
                string oPath = @"C:\temp\testoutput.xlsx";
                X.Application excel_python = new X.Application();
                X.Workbook wb_python = excel_python.Workbooks.Open(oPath);
                X._Worksheet ws_python = (X._Worksheet)wb_python.Sheets["Sheet1"];
                Microsoft.Office.Interop.Excel.Range strRange = (Microsoft.Office.Interop.Excel.Range)ws_python.Cells[iteration_num + 1, "F"];
                string full = strRange.Value2;
                string parse_1 = full.Replace("[", "");
                string parse_2 = parse_1.Replace("]", "");
                string[] allDigits = parse_2.Split(',');
                List<string> singleDigits = new List<string>();
                List<int> transform_int = new List<int>();
                foreach (string digit in allDigits)
                {
                    transform_int.Add(Convert.ToInt32(digit));
                }
                wb_python.Close(0);
                excel_python.Quit();
                return transform_int;

            }
            private List<List<int>> GetGridId(int iteration_num)
            {
                string oPath = @"C:\temp\testoutput.xlsx";
                MessageBox.Show("GetGridId - Opening Excel");
                X.Application excel_python = new X.Application();
                MessageBox.Show("GetGridId - Opening Excel Workbook");
                X.Workbook wb_python = excel_python.Workbooks.Open(oPath);
                X._Worksheet ws_python = (X._Worksheet)wb_python.Sheets["Sheet1"];
                Microsoft.Office.Interop.Excel.Range strRange = (Microsoft.Office.Interop.Excel.Range)ws_python.Cells[iteration_num + 1, "D"];
                string full = strRange.Value2;
                string parse_1 = full.Replace("[", "");
                string parse_2 = parse_1.Replace("]", "");
                string parse_3 = parse_2.Replace(" ", "");
                string[] allDigits = parse_3.Split(',');
                List<string> singleDigits = new List<string>();
                List<List<int>> GridIds = new List<List<int>>();
                MessageBox.Show("GetGridId - Parsing Strings");
                foreach (string s in allDigits)
                {

                    singleDigits.Add(s);
                }

                for (int i = 0; i < (singleDigits.Count); i += 2)
                {

                    List<int> index = new List<int>();
                    index.Add(Convert.ToInt32(singleDigits[i]));
                    index.Add(Convert.ToInt32(singleDigits[i + 1]));
                    GridIds.Add(index);
                }
                MessageBox.Show("GetGridId - Exiting Excel");
                wb_python.Close(0);
                excel_python.Quit();
                return GridIds;
            }

            private List<GroupType> GetClusterTypes(int iteration_num, List<GroupType> m_groups)
            {
                List<GroupType> clusters = new List<GroupType>();
                try
                {
                    string oPath = @"C:\temp\testoutput.xlsx";
                    X.Application excel_python = new X.Application();
                    X.Workbook wb_python = excel_python.Workbooks.Open(oPath);
                    X._Worksheet ws_python = (X._Worksheet)wb_python.Sheets["Sheet1"];
                    Microsoft.Office.Interop.Excel.Range strRange = (Microsoft.Office.Interop.Excel.Range)ws_python.Cells[iteration_num + 1, "E"];
                    string full = strRange.Value;
                    string[] parse1 = full.Split(',');
                    List<String> clusterNames = new List<String>();
                    foreach (string s in parse1)
                    {
                        clusterNames.Add(s);
                    }
                    List<string> groupNames = new List<string>();
                    foreach (GroupType groupType in m_groups)
                    {
                        groupNames.Add(groupType.Name);
                    }

                    foreach (string s in clusterNames)
                    {

                        int elem_index = groupNames.IndexOf(s);
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
                
                return clusters;
            }

            private List<List<XYZ>> matrixGen()
            {
                string mPath = @"C:\temp\SOA_Copy.xlsx";
                X.Application excel_matrix = new X.Application();
                X.Workbook wb_matrix = excel_matrix.Workbooks.Open(mPath);
                X._Worksheet ws_matrix = (X._Worksheet)wb_matrix.Sheets["Orthopaedic"];
                List<List<XYZ>> m_xyz = new List<List<XYZ>>();
                try
                {
                    Microsoft.Office.Interop.Excel.Range xRange = (Microsoft.Office.Interop.Excel.Range)ws_matrix.Cells[2, "L"];
                    Microsoft.Office.Interop.Excel.Range yRange = (Microsoft.Office.Interop.Excel.Range)ws_matrix.Cells[3, "L"];
                    Microsoft.Office.Interop.Excel.Range zRange = (Microsoft.Office.Interop.Excel.Range)ws_matrix.Cells[4, "L"];
                    float xVal = (float)xRange.Value;
                    float yVal = (float)yRange.Value;
                    float zVal = (float)zRange.Value;
                    XYZ bottom_left_pt = new XYZ(xVal, yVal, zVal);
                    float grid_size = 1000f / 304.8f;

                    Microsoft.Office.Interop.Excel.Range y_axRange = (Microsoft.Office.Interop.Excel.Range)ws_matrix.Cells[2, "F"];
                    Microsoft.Office.Interop.Excel.Range x_axRange = (Microsoft.Office.Interop.Excel.Range)ws_matrix.Cells[3, "F"];
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

            private List<ElementId> PlaceModelGroup(UIApplication uiapp, List<List<XYZ>> m_xyz, List<GroupType> clusterTypes, List<ElementId> asd, List<List<int>> GridId, List<int> transform)
            {
                MessageBox.Show("Generating Floorplan...");
                Document doc = uiapp.ActiveUIDocument.Document;
                Autodesk.Revit.DB.View active = doc.ActiveView;
                Transaction trans = new Transaction(doc);
                List<ElementId> asd_1 = asd;
                
                try
                {
                    //if (asd_1.Count() != 0)
                    //{

                    //    foreach (ElementId elem_id in asd_1)
                    //    {
                    //        trans.Start("Delete Me");
                    //        doc.Delete(elem_id);
                    //        trans.Commit();
                    //    }
                    //}
                    //FilteredElementCollector collector = new FilteredElementCollector(doc);
                    //ICollection<Element> elem_groups = collector.OfCategory(BuiltInCategory.OST_IOSModelGroups).ToElements();
                    //List<GroupType> all_groups = new List<GroupType>();
                    trans.Start("Model Group Placement");

                //foreach (GroupType grouptype in elem_groups)
                //    if (grouptype.Name == "interviewroom")
                //    {
                //        all_groups.Add(grouptype);
                //    }
                //asd = new List<ElementId>();
                //foreach (GroupType group in all_groups)
                //{
                //    Group testplace = doc.Create.PlaceGroup(XYZ.Zero, group);
                //    asd.Add(testplace.Id);
                //}

                for (int i = 0; i < GridId.Count(); i++)
                {
                    List<int> index = GridId[i];

                    XYZ pt = m_xyz[index[0]][index[1]];
                    float add_x = (4300f / 304.8f) + (float)pt.X;
                    float add_y = (3500f / 304.8f) + (float)pt.Y;
                    XYZ new_pt = new XYZ(add_x, add_y, pt.Z);
                    Group place = doc.Create.PlaceGroup(new_pt, clusterTypes[i]);
                    List<ElementId> place_id = new List<ElementId>{place.Id};
                    if (transform[i] == 1)
                    {
                        XYZ pt1 = new XYZ(new_pt.X + 1, new_pt.Y, new_pt.Z);
                        Plane plane = Plane.CreateByThreePoints(new_pt, new XYZ(new_pt.X+1, new_pt.Y, new_pt.Z), new XYZ(new_pt.X, new_pt.Y, new_pt.Z + 1));
                        ElementTransformUtils.MirrorElements(doc, place_id, plane,false);
                    }
                    else if (transform[i] == 2)
                    {
                        XYZ pt1 = new XYZ(new_pt.X, new_pt.Y+1, new_pt.Z);
                        Plane plane = Plane.CreateByThreePoints(new_pt, new XYZ(new_pt.X, new_pt.Y+1, new_pt.Z), new XYZ(new_pt.X, new_pt.Y, new_pt.Z + 1));
                        ElementTransformUtils.MirrorElements(doc, place_id, plane, false);
                    }

                    else if (transform[i] == 3)
                    {
                        XYZ pt1 = new XYZ(new_pt.X, new_pt.Y, new_pt.Z+1);
                        Line axis = Line.CreateBound(new_pt, pt1);
                        ElementTransformUtils.RotateElements(doc, place_id, axis, Math.PI);
                    }
                    asd.Add(place.Id);
                }



                trans.Commit();
                //MessageBox.Show(asd.Count.ToString());
                return asd;
                    //transGroup.RollBack();
                }

                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                return null;
                }

            }



        }


    }