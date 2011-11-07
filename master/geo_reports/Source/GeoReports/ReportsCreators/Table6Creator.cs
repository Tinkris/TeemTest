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
    //Ведомость участков пересечения болот 
    public class Table6Creator
    {
        private static IMapViewer mapViewer;
        // строка соединения с базой данных
        private static string videoDbConStr = "";
        private static VideoPresenterEGP presenterEGP;
        private static string templateName;

        private static string kmStringTemplate = "#.000";
        private static string gradStringTemplate = "#.0000";

        /// <summary>
        /// Изменение времени начала трека создателя документов
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
        /// Ведомость участков пересечения болот 
        /// </summary>
        /// <param name="resultFile"></param>
        public static void Create(string resultFile)
        {
            WordDocument doc = null;
            string error = string.Empty;
            try
            {
                List<Egp> egps = presenterEGP.ReadEgpFromDb();
                egps = egps.FindAll(x => x.Code == 7);
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
                    catch (Exception ex)
                    {
                        doc.CloseDocument();
                        doc.Close();
                        throw new Exception("Невозможно прочитать километры в таблице EGP");
                    }
                    string[] row = new string[9];
                    row[0] = i.ToString() + 1; // №
                    row[1] = egps[i].ObjId.ToString(); //Идентификационный номер участка  пересечения болот
                    row[2] = beginKm.ToString(kmStringTemplate); //начало участка - экспл. км
                    row[3] = egps[i].StartPointWgs().X.ToString(gradStringTemplate); // начало участка - Х градус
                    row[4] = egps[i].StartPointWgs().Y.ToString(gradStringTemplate); // начало участка - Y градус
                    row[5] = endKm.ToString(kmStringTemplate); //конец участка - экспл. км
                    row[6] = egps[i].EndPointWgs().X.ToString(gradStringTemplate); //конец участка - Х градус
                    row[7] = egps[i].EndPointWgs().Y.ToString(gradStringTemplate); //конец участка - Y градус
                    row[8] = (endKm - beginKm).ToString(kmStringTemplate); // протяженность участка, км.
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
