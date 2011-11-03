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
    public class JCreator
    {
        private static IMapViewer mapViewer;
        // строка соединения с базой данных
        private static string videoDbConStr = "";
        private static VideoPresenterEGP presenterEGP;
        private static string templateName;



        public void Create()
        {
           
            WordDocument doc = new WordDocument();
            NumberFormatInfo nfi = new CultureInfo("en-US", false).NumberFormat;

            string templateTable1 = Path.Combine(Path.GetDirectoryName(this.GetType().Assembly.Location), ".//WordTemplates//Приложение Ж.dotx");
            WordDocument tempDocument = new WordDocument(templateTable1, true);


           // tempDocument.ReplaceIncludePicture(1, new Uri("C:\\" ), 1);
            tempDocument.ReplaceVariables("point_number", "____");
            tempDocument.ReplaceVariables("description_date", "_________");
            tempDocument.ReplaceVariables("time", "_________");
            tempDocument.ReplaceVariables("point_position", "_________");
            tempDocument.ReplaceVariables("process_type", "___________________________________________________________________________________________________________________________________");
            tempDocument.ReplaceVariables("relief_form", "____________________________________________________________________________________________________________________________________");
            tempDocument.ReplaceVariables("morfology", "______________________________________________________________________________________________________________________________________");
            tempDocument.ReplaceVariables("Water", "__________________________________________________________________________________________________________________________________________");
            tempDocument.ReplaceVariables("Grass", "__________________________________________________________________________________________________________________________________________");


        }
    }
}