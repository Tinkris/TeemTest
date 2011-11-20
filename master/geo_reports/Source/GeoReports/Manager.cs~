using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MapViewLibrary;
using GeoReports.ReportsCreators;
using System.IO;
using GeoReports;
using System.Windows.Forms;
using GeoReports.Words;

namespace GeoReports
{
    public class Manager
    {
        private frmMain form;
        private CreateReportFrm cform;
        private IMapViewer mapViewer;
        private string templateTable1;
        private string templateTable6;
        private string templateTable7;
        private string templateTable8;
        private string templateTable9;
        private string templateTable10;
        private string templateTable11;
        private string templateG;
        private string georeportsConnectionString;

        public Manager(frmMain _form)
        {
            form = _form;
           
            templateTable9 = Path.Combine(Path.GetDirectoryName(this.GetType().Assembly.Location), ".//WordTemplates//Таблица 9.dotx");
            templateTable10 = Path.Combine(Path.GetDirectoryName(this.GetType().Assembly.Location), ".//WordTemplates//Таблица 10.dotx");
            templateTable11 = Path.Combine(Path.GetDirectoryName(this.GetType().Assembly.Location), ".//WordTemplates//Таблица 11.dotx");
            templateG = Path.Combine(Path.GetDirectoryName(this.GetType().Assembly.Location), ".//WordTemplates//Приложение Баба .dotx");   
            templateTable1 = Path.Combine(Path.GetDirectoryName(this.GetType().Assembly.Location), ".//WordTemplates//Таблица 1.dotx");
            templateTable6= Path.Combine(Path.GetDirectoryName(this.GetType().Assembly.Location), ".//WordTemplates//Таблица 6.dotx");
            templateTable7 = Path.Combine(Path.GetDirectoryName(this.GetType().Assembly.Location), ".//WordTemplates//Таблица 7.dotx");
            templateTable8 = Path.Combine(Path.GetDirectoryName(this.GetType().Assembly.Location), ".//WordTemplates//Таблица 8.dotx");
        }

        public void SetViewer(IMapViewer _mapViewer)
        {
            mapViewer = _mapViewer;
            Table1Creator.Init(mapViewer,templateTable1);
            Table6Creator.Init(mapViewer, templateTable6);
            Table7Creator.Init(mapViewer, templateTable7);
            Table8Creator.Init(mapViewer, templateTable8);
            Table9Creator.Init(mapViewer, templateTable9);
            Table10Creator.Init(mapViewer, templateTable10);
            Table11Creator.Init(mapViewer, templateTable11);
            GCreator.Init(mapViewer, templateG);
            GetConnection();
            DataBaseWorker.Init(georeportsConnectionString);
        }

        public void WriteTable1(string resultFile)
        {
            Table1Creator.Create(resultFile);
        }

        public void WriteTable6(string resultFile)
        {
            Table6Creator.Create(resultFile);
        }

        public void WriteTable7(string resultFile)
        {
            Table7Creator.Create(resultFile);
        }

        public void WriteTable8(string resultFile)
        {
            Table8Creator.Create(resultFile);
        }

        public void WriteTable9(string resultFile)
        {
            Table9Creator.Create(resultFile);
        }

        public void WriteTable10(string resultFile)
        {
            Table10Creator.Create(resultFile);
        }

        public void WriteTable11(string resultFile)
        {
            Table11Creator.Create(resultFile);
        }

        public void WriteG(string resultFile)
        {
            GCreator.Create(resultFile);
        }

        private void GetConnection()
        {
            if (mapViewer != null)
            {
                DataBaseWorker.Init(georeportsConnectionString);
            }
        }

        public void AbortReportCreating()
        {
            WordDocument.KillAll();
        }

        
        public void UpdateTree(TreeView tree)
        {
            DataBaseWorker.GetReportsData(tree);
           /*// Georeports db = new Georeports(georeportsConnectionString);
            
            /*IEnumerable<DocumentsReport> table = 
                (from reports in db.DocumentsReports   
                 orderby reports.CatalogsReport
                 select reports);
            

            IEnumerable<CatalogsReport> table = 
                (from catalogs in db.CatalogsReports   
                 orderby catalogs.Id
                 select catalogs);

            tree.Nodes.Clear();

            foreach (CatalogsReport cr in table)
            {
                TreeNode newCatalogNode = new TreeNode(cr.Id.ToString() + " " + cr.Name);
                IEnumerable<DocumentsReport> tableDocs =
                    (from reports in db.DocumentsReports
                     where reports.CatalogsId == cr.Id
                     orderby reports.CreateDate
                     select reports);

                TreeNode dateParent = null;
                DateTime currCreateDate = DateTime.Now;
                foreach (DocumentsReport dr in tableDocs)
                {
                    if (dr.CreateDate != currCreateDate)
                    {
                        currCreateDate = dr.CreateDate;
                        dateParent = new TreeNode(currCreateDate.ToShortDateString());
                        tree.Nodes.Add(dateParent);
                    }
                    dateParent.Nodes.Add(dr.NameOfREportsSet.ToString() + ": " + dr.WordName);
                }
            }
            */
        }
    }
}
