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
    class Table11Creator
    {
        private static IMapViewer mapViewer;
        // строка соединения с базой данных
        private static string videoDbConStr = "";
        private static VideoPresenterEGP presenterEGP;
        private static string templateName;

        private static string kmStringTemplate = "#.000";

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
        /// Общие данные по протяженности участков с экзогенными геологическими процессами
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

                double cat1, cat2, cat3;
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
                        row[0] = previousName;          //Геологический процесс
                        row[1] = cat1.ToString(kmStringTemplate);       // 1 категория - протяженность
                        row[2] = cat2.ToString(kmStringTemplate);       // 2 категория - протяженность
                        row[3] = cat3.ToString(kmStringTemplate);       // 3 категория - протяженность
                        row[4] = (cat1 + cat2 + cat3).ToString(kmStringTemplate); // общее число
                        doc.AddDataToTable(1, row);
                        cat1 = cat2 = cat3 = 0;

                    }

                    if (i != egps.Count)
                    {
                        char[] splitChar = { '-' };
                        string[] km = egps[i].Name.Split(splitChar);

                        double beginKm = 0;
                        double endKm = 0;
                        try
                        {
                            beginKm = Double.Parse(km[0], nfi);
                            endKm = Double.Parse(km[1], nfi);
                        }
                        catch (Exception)
                        {
                            doc.CloseDocument();
                            doc.Close();
                            throw new Exception("Невозможно прочитать километры в таблице EGP");
                        }

                        double delta = endKm - beginKm;
                        switch (egps[i].Category)
                        {
                            case 1: cat1 += delta;
                                break;
                            case 2: cat2 += delta;
                                break;
                            default: cat3 += delta;
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
