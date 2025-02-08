using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using System.Reflection;
using System.IO;
using Jungle_RVT_Automatic_ifc_export.Tools;
using System.Windows.Forms;

namespace Jungle_RVT_Automatic_ifc_export
{
    public class MyCommandHandler : IExternalEventHandler
    {
        private string _nameView;
        private string _nameExportSetup;
        private string _dirExportFile; 
        private string _nameExportFile;
        private List<ExportFileSetting> _exportSettings;
                
        public void Execute(UIApplication app)
        {
            foreach (ExportFileSetting set in _exportSettings)
            {
                string _nameView = set.NameView;
                string _nameExportSetup = set.NameExportSetup;
                string _dirExportFile = set.DirExportFile;
                string _nameExportFile = set.NameExportFile;
                try
                {
                    View3D viewExport = new FilteredElementCollector(app.ActiveUIDocument.Document)
                    .OfClass(typeof(View3D))
                    .Where(x => x.Name == _nameView)
                    .Cast<View3D>().First();
                    ElementId elementIdView = viewExport.Id;

                    IFCExportOptions ifcOptions = new IFCExportOptions
                    {
                        FileVersion = IFCVersion.IFC2x3, // Используем общепринятую версию
                        ExportBaseQuantities = true, // Экспортировать количественные данные
                        SpaceBoundaryLevel = 0, // Пространственные границы
                                                //FamilyMappingFile = pathMapping,
                        WallAndColumnSplitting = true
                    };

                    ifcOptions.FilterViewId = elementIdView;

                    Rootobject rootobject = JsonConvert.DeserializeObject<Rootobject>(File.ReadAllText(_nameExportSetup));
                    PropertyInfo[] properties = rootobject.GetType().GetProperties();
                    foreach (PropertyInfo property in properties)
                    {
                        string propertyName = property.Name;
                        string valueProperty = property.GetValue(rootobject).ToString();
                        ifcOptions.AddOption(propertyName, valueProperty);
                    }

                    using (Transaction tr = new Transaction(app.ActiveUIDocument.Document))
                    {
                        tr.Start("Export");

                        app.ActiveUIDocument.Document.Export(_dirExportFile, _nameExportFile, ifcOptions);
                        tr.Commit();                        
                    }
                    HistoryBuilder.WriteExportFile(_nameView,
                            _nameExportSetup,
                            _dirExportFile,
                            _nameExportFile);
                }
                catch (Exception ex)
                {                    
                    HistoryBuilder.WriteError(ex.Message);
                }
            }                            
        }

        public string GetName()
        {
            return "My External Event Handler";
        }

        public void SetProperty(string nameView, string nameExportSetup,
            string dirExportFile, string nameExportFile)
        {
            _nameView = nameView;
            _nameExportFile = nameExportFile;
            _dirExportFile = dirExportFile;
            _nameExportFile = nameExportFile;
        }

        public void SetProperty(List<ExportFileSetting> settings)
        {
            _exportSettings = settings;
        }

    }
}
