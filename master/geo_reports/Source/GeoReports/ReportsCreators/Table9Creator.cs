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
    public class Table9Creator
    {
        private static IMapViewer mapViewer;
        // строка соединения с базой данных
        private static string videoDbConStr = "";
        private static VideoPresenterEGP presenterEGP;
        private static string templateName;

        private static string gradStringTemplate = "#.0000";

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
        /// Ведомость выявленных на момент проведения обследования форм рельефа, связанных с развитием ЭГП/ареалов СГУ  
        /// </summary>
        /// <param name="resultFile">имя выходного файла</param>
        public static void Create(string resultFile)
        {
            WordDocument doc = null;
            string error = string.Empty;
            try
            {
                List<Egp> egps = presenterEGP.ReadEgpFromDb();
                egps = egps.FindAll(x => x.Code == 9);
                doc = new WordDocument(templateName, false);

                for (int i = 0; i < egps.Count; i++)
                {

                    string[] row = new string[5];
                    row[0] = i.ToString() + 1; //todo check - Номер участка с проявлением  ЭГП/ареала СГУ
                    row[1] = egps[i].Type;     // Название  ЭГП/ СГУ

                    row[2] = egps[i].ObjId.ToString();  //Идентификационный номер формы рельефа ЭГП/ареала СГУ
                    row[3] = egps[i].StartPointWgs().X.ToString(gradStringTemplate); //todo check Координаты контура формы рельефа*
                    row[4] = egps[i].StartPointWgs().Y.ToString(gradStringTemplate);
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
