using System;
using System.Collections.Generic;
using System.Linq;
using System.Configuration;
using System.IO;
using NUnrar.Archive;
using NUnrar.Common;
using System.Diagnostics;
using System.Threading;


namespace CheckArchive
{
    class Program
    {
        // путь к папке куда сохраним архив
        static string PathToSave = ConfigurationManager.AppSettings["pathToSave"];

        // путь к папке с архивами из конфига
        static readonly string PathToArchive = ConfigurationManager.AppSettings["pathToArchive"];
         
        // маска архива
        static readonly List<string> CorrectFoldersItem = new List<string>(ConfigurationManager.AppSettings["maskOfArchive"].Split(new char[] { ';' }));

        // полный путь к файлу
        static readonly string PathToLastFileInArchive = GetLastArchive(PathToArchive);

        // информация о файле (дата создания и т.д.)
        static readonly FileInfo LastFileInfo = new FileInfo(PathToLastFileInArchive);
        
        // KGA_YYYY_MM_DD.rar
        static readonly string CorrectNameOfFile = FormatNameOfFile(LastFileInfo.LastWriteTime.Month, LastFileInfo.LastWriteTime.Day);

        static void Main(string[] args)
        {
            CheckPathToSaveExists();

            ShowPathToArhiveAndToSave();

            CheckInfoFile();
        }

        /// <summary>
        /// Проверяем есть ли файл с информацией об изменениях в папке с архивами
        /// </summary>
        static void CheckInfoFile()
        {
            if (File.Exists(Environment.CurrentDirectory + "\\" + "LastFileInfo.txt"))
            {
                if (!(LastFileInfo.CreationTime.ToString() == File.ReadAllLines(Environment.CurrentDirectory + "\\" + "LastFileInfo.txt")[0]))
                {
                    CreateInfoFileAndCheckAndSaveArchive();
                }
            }
            else
            {
                CreateInfoFileAndCheckAndSaveArchive();
            }
        }

        /// <summary>
        /// Основной блок программы, который включает содание файла с информацией об последнем созданном файле, 
        /// а так же проверка и сохранение нового архива, 
        /// И удаление временных файлов
        /// </summary>
        static void CreateInfoFileAndCheckAndSaveArchive()
        {
            File.AppendAllText(Environment.CurrentDirectory + "\\" + "LastFileInfo.txt", LastFileInfo.CreationTime.ToString());

            CheckAndSaveArhive();

            DeleteTempFiles();
        }
        
        /// <summary>
        /// Выводит информацию об пути к папке сохранения и откуда беруться архивы
        /// </summary>
        static void ShowPathToArhiveAndToSave()
        {

            Console.WriteLine("Путь к папке с архивами: {0}", PathToArchive);

            Console.WriteLine("Путь к папке для сохранения: {0}", PathToSave);
        }

        /// <summary>
        /// Проверяет есть ли папка для сохранения из конфига. если нету то сохраняет в месте вызова
        /// </summary>
        static void CheckPathToSaveExists()
        {
            if (!Directory.Exists(PathToSave))
            {
                PathToSave = Environment.CurrentDirectory + "\\";
            }
        }

        /// <summary>
        /// Процедура проверки архива и если он не корректный сохранение его в указанную в конфиге папку
        /// </summary>
        static void CheckAndSaveArhive(){

            // Проверка на соответсивие имени файла шаблону
            if (CorrectNameOfFile == PathToLastFileInArchive.Remove(0, PathToArchive.Length))
            {
                Console.WriteLine("Имя файла {0} сответствует шаблону : {1}", PathToLastFileInArchive.Remove(0, PathToArchive.Length), CorrectNameOfFile);

                // создали папку с файлами из архива
                CreateFolderWithArchiveFiles(PathToLastFileInArchive, CorrectNameOfFile);

                // проверили структуру папки
                if (!CheckItemsFolder(CorrectNameOfFile, CorrectFoldersItem))
                {
                    IfNotCorrectStructure(CorrectNameOfFile);
                }
            }
            else
            {
                Console.WriteLine("Имя файла {0} не сответствует шаблону : {1}", PathToLastFileInArchive.Remove(0, PathToArchive.Length), CorrectNameOfFile);

                CreateFolderWithArchiveFiles(PathToLastFileInArchive, CorrectNameOfFile);

                if (CheckItemsFolder(CorrectNameOfFile, CorrectFoldersItem))
                {
                    IfCorrectStructure(CorrectNameOfFile);
                }
                else
                {
                    IfNotCorrectStructure(CorrectNameOfFile);
                }

            }
        }
        
        /// <summary>
        /// Удаляет времнные файлы и папки
        /// </summary>
        static void DeleteTempFiles()
        {
            if (Directory.Exists(Path.GetTempPath() + CorrectNameOfFile.Replace(".rar", "")))
            {
                Directory.Delete(Path.GetTempPath() + CorrectNameOfFile.Replace(".rar", ""),true);
            }

            if (Directory.Exists(PathToSave + CorrectNameOfFile.Replace(".rar", "")))
            {
                Directory.Delete(PathToSave + CorrectNameOfFile.Replace(".rar", ""), true);
            }

        }

        /// <summary>
        /// Если структура папки корректна, то делаем это
        /// </summary>
        /// <param name="correctNameOfFile"></param>
        static void IfCorrectStructure(string correctNameOfFile)
        {
            Console.WriteLine("Структура корректна");

            CreateArchive(PathToSave, correctNameOfFile, Path.GetTempPath() + correctNameOfFile.Replace(".rar", "") + "\\");
        }

        /// <summary>
        /// Если структура папки не корректна, до делаем это
        /// </summary>
        /// <param name="correctNameOfFile"></param>
        static void IfNotCorrectStructure(string correctNameOfFile)
        {
            Console.WriteLine("Структура не корректна");

            string pathToFolderForArchived = CreateCorrectFolders(correctNameOfFile);//создали кореектную структуру

            CreateArchive(PathToSave, correctNameOfFile, pathToFolderForArchived);//записали в архив
        }

        /// <summary>
        /// Создаем архив с данными
        /// </summary>
        /// <param name="pathToSave">куда</param>
        /// <param name="correctNameOfFile">в какой файл записать</param>
        /// <param name="pathToFolderForArchived">откуда</param>
        static void CreateArchive(string pathToSave,string  correctNameOfFile,string pathToFolderForArchived)
        {
            // string arg = String.Format("a D:\temp\result\test.rar -ep1 -r D:\temp\result\KGA_2020-03-18\",);
            
            Console.WriteLine("Создаем архив...");

            string arg = String.Format("m {0} -ep1 -r {1}", pathToSave + correctNameOfFile, pathToFolderForArchived);

            ProcessStartInfo ps = new ProcessStartInfo();

            ps.FileName = @"C:\Program Files\WinRAR\RAR.exe";

            ps.Arguments = arg;

            Process.Start(ps);

            DelayForProcess();
        }

        /// <summary>
        /// создает правильно сформированную папку с файлами для архива
        /// </summary>
        /// <param name="correctNameOfFile"></param>
        /// <returns>путь, где лежит папка с файлами правильными</returns>
        static string CreateCorrectFolders(string correctNameOfFile)
        {
            Console.WriteLine("Создние времеменной папки для хранения правильного набора данных...");

            Directory.CreateDirectory(PathToSave + correctNameOfFile.Replace(".rar", ""));

            for (int i = 0; i < CorrectFoldersItem.Count - 1; i++)
            {
                Directory.CreateDirectory(PathToSave + correctNameOfFile.Replace(".rar", "") + "\\" + CorrectFoldersItem[i]);
            }

            RarArchive archive = RarArchive.Open(PathToLastFileInArchive);//добавили файлы из архива в поток

            foreach (RarArchiveEntry entry in archive.Entries) // проходимся по ним в цикле и записываем в папку темп
            {
                string filename = Path.GetFileName(entry.FilePath);//файл;

                if (entry.ToString().Contains(CorrectFoldersItem[0]))//all_twn
                {
                    entry.WriteToFile(PathToSave + correctNameOfFile.Replace(".rar", "") + "\\" + CorrectFoldersItem[0] + "\\" + filename, ExtractOptions.ExtractFullPath | ExtractOptions.Overwrite);

                }
                else if (entry.ToString().Contains(CorrectFoldersItem[1]))//cl_cod
                {
                    entry.WriteToFile(PathToSave + correctNameOfFile.Replace(".rar", "") + "\\" + CorrectFoldersItem[1] + "\\" + filename, ExtractOptions.ExtractFullPath | ExtractOptions.Overwrite);

                }
                else//rgis.xml
                {
                    entry.WriteToFile(PathToSave + correctNameOfFile.Replace(".rar", "") + "\\" + filename, ExtractOptions.ExtractFullPath | ExtractOptions.Overwrite);

                }

            }

            return PathToSave + correctNameOfFile.Replace(".rar", "")+"\\";
        }

        /// <summary>
        /// Возвращает шаблон названия архива
        /// </summary>
        /// <param name="countMonth">Номер месяца</param>
        /// <returns>Форматированное название</returns>
        static string FormatNameOfFile(int countMonth, int countDay)
        {
            string correctNameOfFile = null;//

            if (countMonth < 10) //форматирую вывод
            {
                correctNameOfFile = String.Format("KGA_{0}-0{1}-{2}.rar", LastFileInfo.LastWriteTime.Year, LastFileInfo.LastWriteTime.Month, LastFileInfo.LastWriteTime.Day);
            }
            else if ( countDay < 10) //форматирую вывод
            {
                correctNameOfFile = String.Format("KGA_{0}-{1}-0{2}.rar", LastFileInfo.LastWriteTime.Year, LastFileInfo.LastWriteTime.Month, LastFileInfo.LastWriteTime.Day);
            }
            else if (countMonth < 10 && countDay < 10) //форматирую вывод
            {
                correctNameOfFile = String.Format("KGA_{0}-0{1}-0{2}.rar", LastFileInfo.LastWriteTime.Year, LastFileInfo.LastWriteTime.Month, LastFileInfo.LastWriteTime.Day);
            }
            else
            {
                correctNameOfFile = String.Format("KGA_{0}-{1}-{2}.rar", LastFileInfo.LastWriteTime.Year, LastFileInfo.LastWriteTime.Month, LastFileInfo.LastWriteTime.Day);
            }

            return correctNameOfFile;
        }

        /// <summary>
        /// Проверка структуры папки
        /// </summary>
        /// <param name="correctNameOfFile"></param>
        /// <param name="correctFoldersItem"></param>
        /// <returns>true , если корректна</returns>
        static bool CheckItemsFolder(string correctNameOfFile, List<string> correctFoldersItem)
        {
            Console.WriteLine("Проверка соответствия структуры папки...");

            List<string> itemInFolder = new List<string>(Directory.GetDirectories(Path.GetTempPath() + correctNameOfFile.Replace(".rar", "")).Concat(Directory.GetFiles(Path.GetTempPath() + correctNameOfFile.Replace(".rar", ""))));

            for (var i = 0; i < itemInFolder.Count; i++)
            {
                itemInFolder[i] = itemInFolder[i].Replace(Path.GetTempPath() + correctNameOfFile.Replace(".rar", "") + "\\", "");
            }

            return itemInFolder.SequenceEqual(correctFoldersItem) ? true : false;
        }

        /// <summary>
        /// Создает в папке temp папку с разархиваированными файлами и папками
        /// </summary>
        /// <param name="lastFile">путь к архиву</param>
        static void CreateFolderWithArchiveFiles(string lastFile, string correctNameOfFile)
        {
            Console.WriteLine("Создание временной папки с данными архива...");

            Directory.CreateDirectory(Path.GetTempPath() + correctNameOfFile.Replace(".rar", ""));

            string arg = String.Format("x {0} {1} ",lastFile , Path.GetTempPath() + correctNameOfFile.Replace(".rar", ""));

            ProcessStartInfo ps = new ProcessStartInfo();

            ps.FileName = @"C:\Program Files\WinRAR\RAR.exe";

            ps.Arguments = arg;
            
            Process.Start(ps);

            DelayForProcess();
        }

        /// <summary>
        /// Задержка, пока выполняется процесс 
        /// </summary>
        static void DelayForProcess()
        {

            string i = "";

            while (true)
            {
                if (Process.GetProcessesByName("rar").Length == 0)
                {
                    break;
                }

                Console.Write("\r{0}", i);
                i += ".";
                Thread.Sleep(1000);

            }

            Console.WriteLine();
        }   

        /// <summary>
        ///Отдает путь к последнему архиву в папке 
        /// </summary>
        /// <param name="pathToArchive"></param>
        /// <returns></returns> 
        static string GetLastArchive(string pathToArchive)
        {
            DirectoryInfo di = new DirectoryInfo(PathToArchive);

            var sortedFeed = (from s in di.GetFiles() orderby s.LastWriteTime descending select s).First();

            return di.ToString() + sortedFeed;
        }

    }
}
