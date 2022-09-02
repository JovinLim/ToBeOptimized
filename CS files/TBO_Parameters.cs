using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Data;
using System.Drawing;
using System.Threading.Tasks;
using System.Diagnostics;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using Autodesk.Revit.DB.Architecture;
using System.Windows.Forms;

namespace TBO_Plugin
{
	[Transaction(TransactionMode.Manual)]
	public class Parameters : IExternalCommand
	{
		public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
		{
			// Get the application and document from external command data.
			UIApplication uiApp = commandData.Application;
			Document doc = uiApp.ActiveUIDocument.Document;
			/*System.Windows.Forms.Form test_form = new LayoutGencs(doc);
			test_form.Show();*/
			using (System.Windows.Forms.Form form = new Param(doc))
			{
                if (form.ShowDialog() == DialogResult.OK)
                {
                    return Result.Succeeded;
                }
                else
                {
					form.Dispose();
                    return Result.Succeeded;
                }
            }
		}
	}
}

		