using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jungle_RVT_Automatic_ifc__export
{
    [Transaction(TransactionMode.Manual)]
    public class Comamnd : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            TaskDialog.Show("Сообщение", "Привет, мир!");
            return Result.Succeeded; ;
        }
    }
}
