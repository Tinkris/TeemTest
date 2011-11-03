using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.ComponentModel;
using Microsoft.SqlServer.Types;

using MapClassLibrary;
using MapBase;
using D2Types;
using GeoReports.ReportsCreators;

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

namespace GeoReports
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
           // Application.EnableVisualStyles();
           // Application.SetCompatibleTextRenderingDefault(false);
           // Application.Run(new frmMain());
            WriteReport();
        }

        static void WriteReport()
        {
                MapObjSet objSet = MifReader.ReadMifMid(@"C:\Users\Тинкс\Documents\Visual Studio 2010\Projects\Работа\geo_reports\Source\ареалы_май_2011_final.mif");
                WordDocument doc = new WordDocument();
                int i = 109;
                StreamWriter writer = new StreamWriter("C:/noPictures.txt");

                List<MapEntity> sorted = new List<MapEntity>();
                foreach (MapEntity me in objSet)
                {
                    sorted.Add(me);
                }
                sorted.Sort(delegate(MapEntity me1, MapEntity me2)                        
                {
                    string kmMin1 = me1.Semantics["ISN"].ToString();
                    string kmMin2 = me2.Semantics["ISN"].ToString();
                    
                    return (Double.Parse(kmMin1)).CompareTo(Double.Parse(kmMin2)); 
                });
                foreach (MapEntity me in sorted)
                {

                    /// MapEntity me - объект, содержит географию и семантику
                    DPoint[] p = me.Points; // точки, которые описывают географию
                    DataRow r = me.Semantics; // строка таблицы, содержит всю семантику
                    
                    
                    string pointNumber = r["ISN"].ToString();
                    DateTime pointTime = DateTime.Parse( r["дата_время"].ToString());
                    string kmMin = r["KM_MIN"].ToString();
                    string kmMax = r["KM_MAX"].ToString();
                    string picturePath = r["PhotoISN"].ToString();
                    string longMin = r["Long_MIN"].ToString();
                    string latMin = r["Lat_MIN"].ToString();
                    string longMax = r["Long_MAX"].ToString();
                    string latMax = r["Lat_MAX"].ToString();
                    string PhotoName = r["Фотка"].ToString();


                    string egp = r["ЭГП"].ToString();
                    int egpType = Int32.Parse( r["Тип_ЭГП"].ToString());
                    string reliefWidth = r["макс_ширина"].ToString();
                    if (reliefWidth.Length == 0) reliefWidth = "  ";
                 //   string reliefLength = r["протяженность"].ToString();
                    string orientir = r["напр_измен_границ"].ToString();
                    string water = r["обводнение"].ToString();
                    string grass = r["нарушен_растительности"].ToString();


                    if (!File.Exists((Path.Combine("L:\\Photo_ЭГП", picturePath + ".jpg"))))
                    {
                      //  writer.WriteLine(picturePath + "     " + PhotoName);
                        continue;
                    }
                    if (egpType == 1 ||
                        egpType == 2 ||
                        egpType == 3 ||
                        egpType == 4 ||
                        egpType == 5 ||
                        egpType == 6 ||
                        egpType == 7 ||
                        egpType == 8 ||
                        egpType == 9 ||
                        egpType == 11 ||
                        egpType == 12 ||
                        egpType == 20 ||
                        egpType == 22 ||
                        egpType == 26 
                        )
                    {
                        WordDocument tempDocument = new WordDocument(@"C:\Users\Тинкс\Documents\Visual Studio 2010\Projects\Работа\geo_reports\Source\GeoReports\WordTemplates\Приложение Ж.dotx", false);

                        tempDocument.ReplaceVariables("PointNumber", pointNumber);
                        tempDocument.ReplaceVariables("description_date", pointTime.Date.ToString("dd.MM.yyyy"));
                        tempDocument.ReplaceVariables("time", pointTime.ToString("hh:mm"));
                        tempDocument.ReplaceVariables("KmMin-KmMax", kmMin + "-" + kmMax);
                        tempDocument.ReplaceVariables("point_position1", longMin + ";" + latMin);
                        tempDocument.ReplaceVariables("point_position2", longMax + ";" + latMax);
                        tempDocument.ReplaceVariables("PhotoName", PhotoName);
                        tempDocument.ReplaceVariables("PhotoISN", picturePath);
                        tempDocument.ReplaceVariables("EGPType", egp);
                        tempDocument.ReplaceVariables("MaxWidth", reliefWidth);
                        tempDocument.ReplaceVariables("Length", ((int)((Double.Parse(kmMax) - Double.Parse(kmMin))*1000)).ToString());
                        tempDocument.ReplaceVariables("Orientir", orientir);
                        tempDocument.ReplaceVariables("Water", water);
                        tempDocument.ReplaceVariables("Grass", grass);

                        if (Double.Parse(kmMax) < 200)
                        {
                            tempDocument.ReplaceVariables("Sorname", "Макарычева Л.");
                            tempDocument.ReplaceVariables("Proff", "Инженер-исследователь");
                        }
                        else if (Double.Parse(kmMax) < 450)
                        {
                            tempDocument.ReplaceVariables("Sorname", "Разумец М.");
                            tempDocument.ReplaceVariables("Proff", "Инженер-исследователь");
                        }
                        else if (Double.Parse(kmMax) < 960)
                        {
                            tempDocument.ReplaceVariables("Sorname", "Смирнова Л.");
                            tempDocument.ReplaceVariables("Proff", "Инженер-исследователь");
                        }
                        else if (Double.Parse(kmMax) < 1567)
                        {
                            tempDocument.ReplaceVariables("Sorname", "Разумец М.");
                            tempDocument.ReplaceVariables("Proff", "Инженер-исследователь");
                        }
                        else if (Double.Parse(kmMax) < 2300)
                        {
                            tempDocument.ReplaceVariables("Sorname", "Смирнова Л.");
                            tempDocument.ReplaceVariables("Proff", "Инженер-исследователь");
                        }
                        else
                        {
                            tempDocument.ReplaceVariables("Sorname", "Халилова Ю.");
                            tempDocument.ReplaceVariables("Proff", "м.н.с.");
                        }


                        tempDocument.ReplaceIncludePicture(1, new Uri(Path.Combine("L:\\Photo_ЭГП", picturePath + ".jpg")), 1);

                        doc.AddDocument(tempDocument);
                        tempDocument.CloseDocument();
                        tempDocument.Close();
                        tempDocument.Dispose();
                    }
                  }
                doc.Save("C:\\result.doc");
                doc.Close();
                doc.Dispose();

        }
    }
}
