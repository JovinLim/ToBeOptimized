using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Diagnostics;
using System.Reflection;
using System.IO;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.Creation;
using System.Windows.Forms;
using X = Microsoft.Office.Interop.Excel;

namespace TBO_Plugin
{
	[Transaction(TransactionMode.Manual)]
	public class BoundaryInput : IExternalCommand
	{
		public virtual Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			// Get the application and document from external command data.
			UIApplication uiApp = commandData.Application;
			Autodesk.Revit.DB.Document doc = uiApp.ActiveUIDocument.Document;
			Autodesk.Revit.DB.View active = doc.ActiveView;
			Selection sel = uiApp.ActiveUIDocument.Selection;
			Transaction trans = new Transaction(doc);
			FloorPickFilter floorFilter = new FloorPickFilter();
			DoorPickFilter doorFilter = new DoorPickFilter();
			ModelLineFilter lineFilter = new ModelLineFilter();
			ElementId level_id = new ElementId(311);
			Level lv1 = doc.GetElement(level_id) as Level;
			string tPath = @"C:\temp\SOA.xlsx";
			string cDir = Directory.GetCurrentDirectory();
			string nPath = @"C:\temp\SOA_Copy.xlsx";

			try
			{
				// Checking if excel file exists
				try
				{
					if (!File.Exists(tPath))
					{
						TaskDialog.Show("Error", "Please input SOA");
						return Result.Failed;
					}
				}

				catch (Exception ex)
				{
					message = ex.Message;
					TaskDialog.Show("failed", message);
					return Result.Failed;
				}

				try
				{
					if (File.Exists(nPath))
					{
						File.Delete(nPath);
					}
				}

				catch (Exception ex)
				{
					message = ex.Message;
					TaskDialog.Show("failed", message);
					return Result.Failed;
				}

			}
			catch (Exception ex)
			{
				message = ex.Message;
				TaskDialog.Show("failed", message);
				return Result.Failed;
			}

			try
            {
				MessageBox.Show("READ ME \n" +
                    "Step 1 : Select a floor plate.\n" +
					"Step 2 : Select model lines which indicate the main corridor paths connecting to the clinic.\n" +
					"Step 3 : Generate!"
                    );

				//Picking a floor
				TaskDialog.Show("Define Boundary", "Please select a floor");
				Reference pickedfloor = sel.PickObject(ObjectType.Element, floorFilter, "Select a floor"); //Calling method to prompt for floor input

				//Getting Floor as element and checking for voids
				Element floorElem = doc.GetElement(pickedfloor);
				Floor floor = floorElem as Floor;

				//Opening original SOA excel to check total area required for room input
				X.Application excel_original = new X.Application();
				X.Workbook wb_original = excel_original.Workbooks.Open(tPath);
				wb_original.Worksheets["READ ME"].Delete();

				//Getting area of floor
				IList<Parameter> area_param = floor.GetParameters("Area");
				int floor_area = Convert.ToInt32(area_param[0].AsDouble() / 10.764);

				//Checking if input floor boundary is too big or small for input SOA
				X._Worksheet ws_original = (X._Worksheet)wb_original.Sheets["Orthopaedic"];
				Microsoft.Office.Interop.Excel.Range xRange = (Microsoft.Office.Interop.Excel.Range)ws_original.Cells[15, "D"];
				double RoomAreas = (double)xRange.Value2;
				double CorArea = RoomAreas / 2;
				double totalArea = RoomAreas + CorArea;
				if ((floor_area - totalArea) < 0)
				{
					TaskDialog.Show("Error", "Floor area is too small for SOA input.");
					TaskDialog.Show("Close", "Exiting Boundary Input now...");
					return Result.Failed;
				}

				else if (floor_area - totalArea > 30)
				{
					double remainingArea = Math.Round(Math.Abs(floor_area - totalArea) / 10.764 , 2);
					string remainingAreaStr = remainingArea.ToString() + "sqm";
					TaskDialog.Show("Error", "Floor area is " + remainingAreaStr + " more than 30sqm than total required area for SOA input. Excessive space will be generated.");
				}

				//Select connecting path
				TaskDialog.Show("Define Corridor", "Select Model Lines");
				IList<Reference> pickedpath = sel.PickObjects(ObjectType.Element, lineFilter, "Select model lines"); //Calling method to prompt for model line input

				//Get sketch of floor
				Sketch floorSketch = doc.GetElement(floor.SketchId) as Sketch;
				IList<ElementId> floorelems_id = floorSketch.GetAllElements();
				List<ElementId> sketchLinesId = new List<ElementId>();
				foreach (ElementId floorelems in floorelems_id)
                {
					if (doc.GetElement(floorelems).Name == "Model Lines")
                    {
						sketchLinesId.Add(floorelems);
                    }
                }
				//Getting all model lines of floor sketch
				List<Curve> linesList = new List<Curve>();
				//MessageBox.Show(sketchLinesId.Count.ToString());
				foreach (ElementId sketchLineId in sketchLinesId)
                {
					
					CurveElement sketchLine = doc.GetElement(sketchLineId) as CurveElement;
					Curve sketchCrv = sketchLine.GeometryCurve;
					linesList.Add(sketchCrv);
                }
				List<XYZ> vertice = new List<XYZ>();
				List<int> vert_index = new List<int>();
				while (vertice.Count < linesList.Count)
                {
					for (int i = 0; i < linesList.Count; i++)
                    {
						if (i == 0)
                        {
							vertice.Add(linesList[i].GetEndPoint(0));
							vert_index.Add(0);
                        }

                        else
                        {
							for (int j = 0; j < linesList.Count; j++)
                            {
								if (j != vert_index[vert_index.Count - 1])
                                {
									if (Convert.ToInt32(linesList[j].GetEndPoint(0).X) == Convert.ToInt32(vertice[vertice.Count - 1].X) && Convert.ToInt32(linesList[j].GetEndPoint(0).Y) == Convert.ToInt32(vertice[vertice.Count - 1].Y))
									{
										vertice.Add(linesList[j].GetEndPoint(1));
										vert_index.Add(j);
                                    }
									else if (Convert.ToInt32(linesList[j].GetEndPoint(1).X) == Convert.ToInt32(vertice[vertice.Count - 1].X) && Convert.ToInt32(linesList[j].GetEndPoint(1).Y) == Convert.ToInt32(vertice[vertice.Count - 1].Y))
                                    {
										vertice.Add(linesList[j].GetEndPoint(0));
										vert_index.Add(j);
									}

								}
                            }
                        }
					}
                }

                //Getting connecting paths as XYZ Elements
                List<List<XYZ>> endPts = new List<List<XYZ>>();
				foreach (Reference line in pickedpath)
                {
					Options option = new Options();
					Element mLine = doc.GetElement(line);
					CurveElement crvElem = mLine as CurveElement;
					Curve crv = crvElem.GeometryCurve;
					XYZ startPoint = crv.GetEndPoint(0);
					XYZ endPoint = crv.GetEndPoint(1);
					List<XYZ> pts = new List<XYZ> { startPoint, endPoint };
					endPts.Add(pts);
                }

				//Getting bounding box of floor
				BoundingBoxXYZ bb = floor.get_BoundingBox(active);

				//Finding min and max point of bounding box
				XYZ min_pt = bb.Min;
				XYZ max_pt = bb.Max;
				float min_pt_x = (float)min_pt.X;
				float min_pt_y = (float)min_pt.Y;
				float min_pt_z = (float)min_pt.Z;

				//Defining grid size & big cluster grid size
				float grid_size = 3.93701f;

				//Finding x and y differences between right top corner and bottom left corner of bounding box
				float min_y = (float)min_pt.Y;
				float min_x = (float)min_pt.X;
				float max_y = (float)max_pt.Y;
				float max_x = (float)max_pt.X;

				int y_ax = 0;
				int x_ax = 0;
				float y_bound = max_y - min_y;
				float x_bound = max_x - min_x;

				if (y_bound % grid_size != 0)
				{
					y_ax = (int)((int)(y_bound / grid_size));
				}

				else
                {
					y_ax = (int)(y_bound / grid_size);
				}
				if (x_bound % grid_size != 0)
				{
					x_ax = (int)((int)(x_bound / grid_size));
				}

				else
				{
					x_ax = (int)(x_bound / grid_size);
				}

				trans.Start("matrix");
				//Forming matrix, m_xyz as list of Revit XYZ, m_id as matrix indexes
				var m_xyz = new List<List<XYZ>>();
				var m_id = new List<List<string>>();
				for (int y = 0; y < y_ax; y++)
				{
					var x_id = new List<string>();
					var y_id = new List<XYZ>();
					for (int x = 0; x < x_ax; x++)
					{
						XYZ newpt = new XYZ((min_x + x * grid_size), (min_y + y * grid_size), 0);
						y_id.Add(newpt);
						x_id.Add("0");
					}
					m_xyz.Add(y_id);
					m_id.Add(x_id);
				}

				trans.Commit();

				// Finding which matrix points are out of the floor boundary
				List<List<int>> voids = new List<List<int>>();
				trans.Start("testArc");
				for (int y = 0; y < y_ax; y++)
                {
					for (int x = 0; x < x_ax; x++)
                    {
						if (PolygonContain(m_xyz[y][x], vertice) == false)
                        {
							voids.Add(new List<int> { y, x });
                        }
                    }
                }
				trans.Commit();

				// Adding indexes which are connected to paths to a list
				List<List<int>> connected_id = new List<List<int>>();
				for (int c = 0; c < endPts.Count(); c++)
                {
					List<int> x_values = new List<int> { Convert.ToInt32(endPts[c][0].X), Convert.ToInt32(endPts[c][1].X) };
					List<int> y_values = new List<int> { Convert.ToInt32(endPts[c][0].Y), Convert.ToInt32(endPts[c][1].Y) };
					int max_x_val = x_values.Max();
					int min_x_val = x_values.Min();
					int max_y_val = y_values.Max();
					int min_y_val = y_values.Min();

					// If line is horizontal
					if (max_y_val == min_y_val)
                    {
						// If line is below matrix grid
						if (max_y_val < m_xyz[0][0].Y)
                        {
							for (int x = 0; x < x_ax; x++)
                            {
								if (m_xyz[0][x].X > min_x_val && m_xyz[0][x].X < max_x_val)
                                {
									List<int> index = new List<int> { 0, x };
									connected_id.Add(index);
                                }
                            }
                        }

						// If line is above matrix grid
						else if( min_y_val > m_xyz[y_ax - 1][0].Y)
                        {
							for (int x = 0; x < x_ax; x++)
                            {
								if (m_xyz[y_ax - 1][x].X > min_x_val && m_xyz[y_ax - 1][x].X < max_x_val)
                                {
									List<int> index = new List<int> { y_ax - 1, x };
									connected_id.Add(index);
                                }
                            }
                        }
                    }

					// If line is vertical
					else if (max_x_val == min_x_val)
                    {
						// If line is on left of matrix grid
						if (max_x_val < m_xyz[0][0].X)
                        {
							for (int y = 0; y < y_ax; y++)
                            {
								if (m_xyz[y][0].Y > min_y_val && m_xyz[y][0].Y < max_y_val)
                                {
									List<int> index = new List<int> { y, 0 };
									connected_id.Add(index);
                                }
                            }
                        }

						// If line is on right of matrix grid
						else if (min_x_val > m_xyz[0][x_ax-1].X)
                        {
							for (int y = 0; y < y_ax; y++)
                            {
								if (m_xyz[y][x_ax-1].Y > min_y_val && m_xyz[y][x_ax-1].Y < max_y_val)
                                {
									List<int> index = new List<int> { y, x_ax - 1 };
									connected_id.Add(index);
                                }
                            }
                        }
                    }
                }

				// Adding main corridor indexes into a complete string
				string indexes = "";
				for (int i=0; i < connected_id.Count(); i++)
                    {
						indexes += connected_id[i][0].ToString() + "," + connected_id[i][1].ToString() + "|";
					}

				// Adding void indexes into a complete string
				string v_index = "";
				for (int i=0; i < voids.Count(); i++)
                {
					v_index += voids[i][0] + "," + voids[i][1] + "|";
                }


				try
				{
                    wb_original.SaveAs(nPath);
                    wb_original.Close(0);
                    X.Application excel_python = new X.Application();
					X.Workbook wb_python = excel_python.Workbooks.Open(nPath);
					X._Worksheet ws_python = (X._Worksheet)wb_python.Sheets["Orthopaedic"];

					// Adding number of rows and columns for 1m x 1m grid matrix
					ws_python.Cells[2, "F"] = "1m Matrix Grid";
					ws_python.Cells[2, "E"] = "Axis";
					ws_python.Cells[3, "E"] = "y_ax";
					ws_python.Cells[4, "E"] = "x_ax";
					ws_python.Cells[3, "F"] = y_ax;
					ws_python.Cells[4, "F"] = x_ax;


					// Adding indexes which connect to main path
					ws_python.Cells[2, "J"] = "Main path indexes";
					ws_python.Cells[3, "J"] = indexes;
					
					// Adding void indexes
					ws_python.Cells[2, "K"] = "Void indexes";
					ws_python.Cells[3, "K"] = v_index;

					// Adding XYZ location of bottom left corner of matrix
					ws_python.Cells[2, "G"] = "Bottom left XYZ of Matrix";
					ws_python.Cells[3, "H"] = "X";
					ws_python.Cells[4, "H"] = "Y";
					ws_python.Cells[5, "H"] = "Z";
					ws_python.Cells[3, "G"] = min_pt_x;
					ws_python.Cells[4, "G"] = min_pt_y;
					ws_python.Cells[5, "G"] = min_pt_z;

					// Saving new SOA for Python
					wb_python.Save();
					wb_python.Close(0);
					excel_original.Quit();
					excel_python.Quit();
					//excel_python.Visible = true;
					return Result.Succeeded;
				}

				catch (Exception ex)
				{
					message = ex.Message;
					MessageBox.Show(message);
					return Result.Failed;
				}

			}

			catch (Exception ex)
			{
				message = ex.Message;
				MessageBox.Show(message);
				return Result.Failed;
			}

		}
		//Creating filter to constrain picking of elements to floors
		public class FloorPickFilter : ISelectionFilter
		{
			public bool AllowElement(Element e)
			{
				return (e.Category.Name == "Floors"); //Defining element category name to allow for selection

			}
			public bool AllowReference(Reference refer, XYZ point)
			{
				return false;
			}
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

		//Creating filter to constrain the picking of elements to model lines
		public class ModelLineFilter : ISelectionFilter
		{
			public bool AllowElement(Element e)
			{
				return (e.Category.Name == "Lines"); //Defining element category name to allow for selection

			}
			public bool AllowReference(Reference refer, XYZ point)
			{
				return false;
			}
		}



		//https://stackoverflow.com/questions/4243042/c-sharp-point-in-polygon
		public bool PolygonContain(XYZ pt, List<XYZ> vertices)
        {
            bool result = false;

            int j = vertices.Count - 1;
            for (int i = 0; i < vertices.Count; i++)
            {
                if (vertices[i].Y < pt.Y && vertices[j].Y > pt.Y || vertices[j].Y < pt.Y && vertices[i].Y > pt.Y)
                {
                    if (vertices[i].X + (pt.Y - vertices[i].Y) / (vertices[j].Y - vertices[i].Y) * (vertices[j].X - vertices[i].X) < pt.X)
                    {
                        result = true;
                    }
                }
                j = i;
            }
            return result;
        }


        //https://dominoc925.blogspot.com/2012/02/c-code-snippet-to-determine-if-point-is.html
        //public bool PolygonContain(XYZ pt, List<Curve> crvList)
        //{
        //	bool result = false;
        //	List<XYZ> vertices = new List<XYZ>();
        //	foreach (Curve c in crvList)
        //	{
        //		vertices.Add(c.GetEndPoint(0));
        //	}

        //	for (int i = 0,j = vertices.Count-1;  i < vertices.Count; j = i++)
        //	{
        //		if (((vertices[i].Y > pt.Y) != (vertices[j].Y > pt.Y)) && (pt.X < (vertices[j].X - vertices[i].X) * (pt.Y - vertices[i].Y) / (vertices[j].Y - vertices[i].Y) + vertices[i].X))
        //		{
        //			result = true;
        //		}
        //	}
        //	return result;
        //}

        public bool CurveContain(XYZ pt, Line crv)
        {
			XYZ startPt = crv.GetEndPoint(0);
			XYZ endPt = crv.GetEndPoint(1);
			double AB = Math.Sqrt(Math.Pow((startPt.X - endPt.X),2) + Math.Pow((startPt.Y - endPt.Y),2));
			double AP = Math.Sqrt(Math.Pow((startPt.X - pt.X),2) + Math.Pow((startPt.Y - pt.Y),2));
			double PB = Math.Sqrt(Math.Pow((endPt.X - pt.X),2) + Math.Pow((endPt.Y - pt.Y),2));
			if (AB == AP + AB)
            {
				return true;
            }

            else
            {
				return false;
            }
        }
	}
}


