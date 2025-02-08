using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SWF = System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Spire.Xls;
using Newtonsoft.Json;
using System.Reflection;
using Jungle_RVT_Automatic_ifc_export.Tools;



namespace Jungle_RVT_Automatic_ifc_export
{
    public class Jungle_ExportIFC_App : IExternalApplication
    {
        private static ExternalEvent _externalEvent;
        private static MyCommandHandler _handler;

        ControlledApplication controlledApplication { get; set; }
        
        
        Document Document { get; set; }
        SWF.Timer UpTimer;
        private int _index_interval = -1;
        private List<int> _list_interval;
        public Result OnShutdown(UIControlledApplication application)
        {            
            return Result.Succeeded;
        }

        public Result OnStartup(UIControlledApplication application)
        {
            _handler = new MyCommandHandler();
            _externalEvent = ExternalEvent.Create(_handler);

            
            controlledApplication = application.ControlledApplication;           
            controlledApplication.DocumentOpened += ControlledApplication_DocumentOpened;
            
            return Result.Succeeded;
        }
       
        private void ControlledApplication_DocumentOpened(object sender, 
            DocumentOpenedEventArgs e)
        {
            if(Document == null)
            {                
                var pathModel = e.Document.PathName;
                FileInfo fileInfo = new FileInfo(pathModel);
                var dir = fileInfo.DirectoryName;
                string pathFile = Path.Combine(
                    dir,
                    "Jungle_RVT_Automatic_ifc_export", 
                    "RVT_Description_ifc_export.xlsx");
                if(new FileInfo(pathFile).Exists)
                {
                    HistoryBuilder.Set_Path_File(Path.Combine(
                    dir,
                    "Jungle_RVT_Automatic_ifc_export",
                    "History.txt"));
                    HistoryBuilder.WriteConnect(pathModel);
                    Document = e.Document;
                    Document.DocumentClosing += Document_DocumentClosing;
                    SetTimer(pathFile);
                    
                }
            }            
        }

        private void Document_DocumentClosing(object sender, DocumentClosingEventArgs e)
        {
            HistoryBuilder.WriteClose(Document.PathName);
            Document = null;

        }

        private void SetTimer(string _path_file_description)
        {
            List<TimeSpan> list = new List<TimeSpan>();
            Workbook workbook = new Workbook();
            workbook.LoadFromFile(_path_file_description);
            string _name_sheet = "Время выгрузки";
            Worksheet sheet = workbook.Worksheets[_name_sheet];
            CellRange locatedRange = sheet.AllocatedRange;
            for (int i = 1; i < locatedRange.Rows.Length; i++)
            {
                list.Add(new TimeSpan(
                    Convert.ToInt32(locatedRange[i + 1, 2].Value),
                    Convert.ToInt32(locatedRange[i + 1, 3].Value),
                    Convert.ToInt32(locatedRange[i + 1, 4].Value)));
            }
            list.Sort();
            HistoryBuilder.WriteSchedule(list);

            DateTime dateNow = DateTime.Now;
            TimeSpan currentTimeSpan = new TimeSpan(dateNow.Hour, dateNow.Minute, dateNow.Second);
            list.Add(currentTimeSpan);
            list.Sort();

            int index = list.IndexOf(currentTimeSpan);

            List<TimeSpan> timeSpans = list.Skip(index)
                .Concat(list.Take(index)
                .Select(x => x + new TimeSpan(24, 0, 0))).ToList();

            var
                dg = timeSpans.Zip(timeSpans.Skip(1), (prev, cur) => cur - prev).ToList();
            dg.Add(timeSpans[1] -
                (timeSpans[timeSpans.Count - 1] - new TimeSpan(24, 0, 0)));

            int startInterval = Convert.ToInt32(dg[0].TotalMilliseconds);
            HistoryBuilder.WriteNextExport(startInterval);
            dg.RemoveAt(0);

            _list_interval = dg.Select(x => Convert.ToInt32(x.TotalMilliseconds)).ToList();
            UpTimer = new SWF.Timer { Enabled = true, Interval = Convert.ToInt32(startInterval) };
            UpTimer.Tick += UpTimer_Tick;
        }

        private void UpTimer_Tick(object sender, EventArgs e)              
        {
            RunExport();
            _index_interval += 1;
            UpTimer.Interval = _list_interval[_index_interval % _list_interval.Count];            
        }

        private void RunExport()
        {            
            FileInfo fileInfo = new FileInfo(Document.PathName);
            var dir = fileInfo.DirectoryName;
            string pathFile = Path.Combine(
                dir,
                "Jungle_RVT_Automatic_ifc_export",
                "RVT_Description_ifc_export.xlsx");

            Workbook workbook = new Workbook();
            workbook.LoadFromFile(pathFile);
            string _name_sheet = "Данные";
            Worksheet sheet = workbook.Worksheets[_name_sheet];
            CellRange locatedRange = sheet.AllocatedRange;

            List<ExportFileSetting> settings = new List<ExportFileSetting>();
            for(int i = 1; i < locatedRange.Rows.Length; i++)
            {
                settings.Add(new ExportFileSetting()
                {
                    NameView = locatedRange[i + 1, 1].Value,
                    NameExportSetup = locatedRange[i + 1, 2].Value,
                    DirExportFile = locatedRange[i + 1, 3].Value,
                    NameExportFile = locatedRange[i + 1, 4].Value,
                });
            }
            ExportIFC(settings);
        }

        private void ExportIFC(
            List<ExportFileSetting> settings)
        {                        
            _handler.SetProperty(settings);
            _externalEvent.Raise();            
        }
    }
}
