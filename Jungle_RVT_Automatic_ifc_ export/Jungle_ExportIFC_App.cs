using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SWF = System.Windows.Forms;
using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Events;
using Autodesk.Revit.UI;
using Spire.Xls;

namespace Jungle_RVT_Automatic_ifc__export
{
    public class Jungle_ExportIFC_App : IExternalApplication
    {
        ControlledApplication controlledApplication { get; set; }
        Application Application { get; set; }
        UIControlledApplication UIControlledApplication { get; set; }
        Document Document { get; set; }
        SWF.Timer UpTimer;
        private int _index_interval = -1;
        private List<int> _list_interval;
        public Result OnShutdown(UIControlledApplication application)
        {
            throw new NotImplementedException();
        }

        public Result OnStartup(UIControlledApplication application)
        {
            UIControlledApplication = application;
            controlledApplication = application.ControlledApplication;

            controlledApplication.DocumentOpened += ControlledApplication_DocumentOpened;
            
            return Result.Succeeded;
        }
        

        private void ControlledApplication_DocumentOpened(object sender, 
            DocumentOpenedEventArgs e)
        {
            if(Document == null)
            {
                
                var s = e.Document.PathName;
                FileInfo fileInfo = new FileInfo(s);
                var dir = fileInfo.DirectoryName;
                string pathFile = Path.Combine(
                    dir,
                    "Jungle_RVT_Automatic_ifc_export", 
                    "RVT_Description_ifc_export.xlsx");
                if(new FileInfo(pathFile).Exists)
                {
                    Document = e.Document;
                    Document.DocumentClosing += Document_DocumentClosing;
                    SetTimer(pathFile);
                    TaskDialog.Show("Revit API", $"Приложение подключилось к модели\n{s}!");
                }
            }            
        }

        private void Document_DocumentClosing(object sender, DocumentClosingEventArgs e)
        {
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
            dg.RemoveAt(0);

            _list_interval = dg.Select(x => Convert.ToInt32(x.TotalMilliseconds)).ToList();
            UpTimer = new SWF.Timer { Enabled = true, Interval = Convert.ToInt32(startInterval) };
            UpTimer.Tick += UpTimer_Tick;
        }

        private void UpTimer_Tick(object sender, EventArgs e)
        {
            
            _index_interval += 1;
            UpTimer.Interval = _list_interval[_index_interval % _list_interval.Count];
            DateTime nowDT = DateTime.Now;
            TimeSpan timeSpan = TimeSpan.FromMilliseconds(_list_interval[_index_interval % _list_interval.Count]);
            DateTime newDt = nowDT + timeSpan;
            TaskDialog.Show("Revit API", $"Текущее время {nowDT}! Следующее время: {newDt}");
        }
    }
}
