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
    public class Table3Creator
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

        /// <summary>
        /// Категории опасности участков с проявлениями экзогенных геологических процессов
        /// </summary>
        /// <param name="resultFile">Имя выходного файла</param>
        public static void Create(string resultFile)
        {
            WordDocument doc = null;
            string error = string.Empty;
            try
            {
                List<Egp> egps = presenterEGP.ReadEgpFromDb();
                doc = new WordDocument(templateName, false);
                NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;

                for (int i = 0; i < egps.Count; i++)
                {
                    char[] splinChar = { '-' };
                    string[] km = egps[i].Name.Split(splinChar);

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
                    string[] row = new string[13];
                    row[0] = i.ToString() + 1;
                    row[1] = egps[i].Code.ToString();
                    row[2] = egps[i].Type;
                    row[3] = egps[i].Category.ToString();
                    row[4] = beginKm.ToString();
                    row[5] = egps[i].StartPointWgs().X.ToString();
                    row[6] = egps[i].StartPointWgs().Y.ToString();
                    row[7] = egps[i].Start.Z.ToString();
                    row[8] = endKm.ToString();
                    row[9] = egps[i].EndPointWgs().X.ToString();
                    row[10] = egps[i].EndPointWgs().Y.ToString();
                    row[11] = egps[i].End.Z.ToString();
                    row[12] = (endKm - beginKm).ToString();

                    doc.AddDataToTable(1, row);
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
