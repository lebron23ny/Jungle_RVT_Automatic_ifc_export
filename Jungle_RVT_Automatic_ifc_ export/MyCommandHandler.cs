using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Jungle_RVT_Automatic_ifc__export
{
    public class MyCommandHandler : IExternalEventHandler
    {
        public void Execute(UIApplication app)
        {
            TaskDialog.Show("Revit API", "Команда выполнена!");
        }

        public string GetName()
        {
            return "My External Event Handler";
        }
    }
}
