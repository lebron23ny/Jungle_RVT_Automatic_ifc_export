using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace Jungle_RVT_Automatic_ifc_export.Tools
{
    public static class HistoryBuilder
    {
        private static string path_file;
        public static void Set_Path_File(string path)
        {
            path_file = path;
        }

        public static void WriteConnect(string name)
        {
            using (StreamWriter writer = new StreamWriter(path_file, true))
            {
                writer.WriteLine($"{DateTime.Now}");
                writer.WriteLine($"Приложение подключилось к модели {name}");
                
            }
        }
        public static void WriteSchedule(List<TimeSpan> schedule)
        {
            using (StreamWriter writer = new StreamWriter(path_file, true))
            {
                writer.WriteLine("Расписание выгрузки:");
                foreach (TimeSpan t in schedule)
                {
                    writer.WriteLine(t.ToString());                    
                }                
            }
        }

        public static void WriteClose(string name)
        {
            using (StreamWriter writer = new StreamWriter(path_file, true))
            {
                writer.WriteLine();
                writer.WriteLine($"{DateTime.Now} cеанс завершен");
                writer.WriteLine($"Модель {name} закрыта");
                writer.WriteLine();
            }
        }

        public static void WriteNextExport(int milliseconds)
        {
            TimeSpan timeSpan = TimeSpan.FromMilliseconds(milliseconds);
            DateTime nextDate = DateTime.Now + timeSpan;
            using (StreamWriter writer = new StreamWriter(path_file, true))
            {
                writer.WriteLine($"Следующий экспорт {nextDate}");                
                writer.WriteLine();
            }
        }

        public static void WriteExportFile(
            string _nameView,
            string _nameExportSetup,
            string _dirExportFile,
            string _nameExportFile
            )
        {
            string path_file_export = Path.Combine(_dirExportFile, _nameExportFile);
            using (StreamWriter writer = new StreamWriter(path_file, true))
            {                
                writer.WriteLine();
                writer.WriteLine($"{DateTime.Now} Успешный успех!");
                writer.WriteLine($"Вид {_nameView} экспортирован" +
                    $"\nс настройками {_nameExportSetup} " +
                    $"в файл {path_file_export}");
                writer.WriteLine();
            }
        }

        public static void WriteError(string message)
        {
            using (StreamWriter writer = new StreamWriter(path_file, true))
            {
                writer.WriteLine();
                writer.WriteLine($"{DateTime.Now} Провальный провал!");
                writer.WriteLine(message);
                writer.WriteLine();
            }
        }
    }
}
