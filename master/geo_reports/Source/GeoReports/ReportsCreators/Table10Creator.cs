using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Esrc.VstoTasksLib;
using MapViewLibrary;
using Esrc.PlayVideoLib;
using GeoReports.Words;
using System.IO;
using System.Globalization;


namespace GeoReports.ReportsCreators
{
    class Table10Creator
    {
        private static IMapViewer mapViewer;
        // строка соединения с базой данных
        private static string videoDbConStr = "";
        private static VideoPresenterEGP presenterEGP;
        private static string templateName;

        /// <summary>
        /// Инициализация создателя документов
        /// </summary>
        /// <param name="_mapViewer"></param>
        public static void Init(IMapViewer _mapViewer, string _templateName)
        {
            mapViewer = _mapViewer;
            templateName = _templateName;
            if (mapViewer != null)
            {
                // чтение строки
                string conStr = mapViewer.GisProject.ExtraDataConnectionStr("Видео");
                if (conStr != null) videoDbConStr = conStr;

                if (videoDbConStr == null)
                    throw new Exception("Не определена база данных аэровизуальных наблюдений!");
                presenterEGP = new VideoPresenterEGP(mapViewer, null, true);
            }
        }

        private static int CompareEGP(Egp x, Egp y)
        {
            if (x.Code > y.Code)
                return 1;
            if (x.Code < y.Code)
                return -1;
            return 0;
        }

        /// <summary>
        /// Общие данные по количеству участков с экзогенными геологическими процессами
        /// </summary>
        /// <param name="resultFile">имя выходного файла</param>
        public static void Create(string resultFile)
        {
            WordDocument doc = null;
            string error = string.Empty;
            try
            {
                List<Egp> egps = presenterEGP.ReadEgpFromDb();

                egps.Sort(CompareEGP);
                doc = new WordDocument(templateName, false);
                NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;

                int cat1, cat2, cat3;
                cat1 = cat2 = cat3 = 0;
                string currentName, previousName;
                currentName = previousName = string.Empty;
                for (int i = 0; i <= egps.Count; i++)
                {
                    if (i != egps.Count)
                        currentName = egps[i].Type;
                    if ((currentName != previousName && i != 0) || i == egps.Count)
                    {
                        string[] row = new string[5];
                        row[0] = previousName; //Геологический процесс
                        row[1] = cat1.ToString();     // 1 категория
                        row[2] = cat2.ToString();     // 2 категория
                        row[3] = cat3.ToString();     // 3 категория
                        row[4] = (cat1 + cat2 + cat3).ToString(); // общее число
                        doc.AddDataToTable(1, row);
                        cat1 = cat2 = cat3 = 0;

                    }

                    if (i != egps.Count)
                    {
                        switch (egps[i].Category)
                        {
                            case 1: cat1++;
                                break;
                            case 2: cat2++;
                                break;
                            default: cat3++;
                                break;
                        }
                        previousName = currentName;
                    }
                }
                doc.Save(resultFile);
            }
            catch (Exception ex)
            {
                error = ex.Message;
            }
            finally
            {
                if (doc != null)
                {
                    doc.CloseDocument();
                    doc.Close();
                }
                if (error != string.Empty)
                    throw new Exception(error);
            }
        }
    }
}
