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
using PictureGPSBinding; 

namespace GeoReports.ReportsCreators
{
    public class GCreator
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

        public static void Create(string resultFile)
        {
            List<Egp> egps = presenterEGP.ReadEgpFromDb();
            string conStr = mapViewer.GisProject.ExtraDataConnectionStr("Снимки");
            //mapViewer
           // presenterEGP.SetSelectEgpMode(true);
           // mapViewer.DrawMap();
            WordDocument doc = new WordDocument();
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

                WordDocument tempDocument = new WordDocument(templateName, false);

                PictureSet pictureSet = new PictureSet(conStr);
                List<Picture> pictures =  pictureSet.getPicturesAtIntervalOfMarks(((int)beginKm).ToString(), ((int)endKm).ToString());
                DateTime pointTime = DateTime.Parse(egps[i].AppearanceTime);
                String longMin = egps[i].Start.Long.ToString();
                String latMin = egps[i].Start.Lat.ToString();
                String longMax = egps[i].End.Long.ToString();
                String latMax = egps[i].End.Lat.ToString();
                String reliefWidth = egps[i].MaxWidth.ToString();

                String orientir = egps[i].Relief;
                String water = egps[i].Watering;
                String grass = egps[i].VegetationAbnormality;
                if (pictures != null)
                {
                    
                    tempDocument.ReplaceIncludePicture(1, new Uri(pictures[0].Session.Folder + "\\" + pictures[0].Name), 1);
                    tempDocument.ReplaceVariables("DayPhoto", pictures[0].Session.Date.ToString());
                    tempDocument.ReplaceVariables("TimePhoto", pictures[0].Session.Date.TimeOfDay.ToString());
                }
                else
                {
                    tempDocument.ReplaceIncludePicture(1, new Uri("C:\\"), 1);
                    tempDocument.ReplaceVariables("DayPhoto", "________");
                    tempDocument.ReplaceVariables("TimePhoto", "________");
                }
                tempDocument.ReplaceVariables("PointNumber", i.ToString());
                tempDocument.ReplaceVariables("description_date", pointTime.Date.ToString("dd.MM.yyyy"));
                tempDocument.ReplaceVariables("time", pointTime.ToString("hh:mm"));
                tempDocument.ReplaceVariables("KmMin-KmMax", beginKm + "-" + endKm);
                tempDocument.ReplaceVariables("point_position1", longMin + ";" + latMin);
                tempDocument.ReplaceVariables("point_position2", longMax + ";" + latMax);
                tempDocument.ReplaceVariables("PhotoName", pictures[0].Name);
                tempDocument.ReplaceVariables("PhotoISN", i.ToString());
                tempDocument.ReplaceVariables("EGPType", egps[i].Category.ToString());
                tempDocument.ReplaceVariables("MaxWidth", reliefWidth);
                tempDocument.ReplaceVariables("Length", ((int)((endKm - beginKm)) * 1000).ToString());
                tempDocument.ReplaceVariables("Orientir", orientir);
                tempDocument.ReplaceVariables("Water", water);
                tempDocument.ReplaceVariables("Grass", grass);

                doc.AddDocument(tempDocument);
                tempDocument.CloseDocument();
                tempDocument.Close();
                tempDocument.Dispose();
            }
            doc.Save(resultFile);
            doc.Close();
            doc.Dispose();
        }
    }
}
