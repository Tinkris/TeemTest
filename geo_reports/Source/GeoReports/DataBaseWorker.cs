using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using GeoReports.Words;
using System.IO;
using System.Data.SqlClient;
using System.Data;
using System.Windows.Forms;

namespace GeoReports
{
    public class DataBaseWorker
    {
        private static string connectionString;
        public static void Init(string _connectionString)
        {
            connectionString = _connectionString;
        }

        public static void CreateCatalog(string name, DateTime begin, DateTime end)
        {
        }

        public static void AddReports(String nameOfReportsSet, List<String> filePaths, int catalogID)
        {
            SqlConnection cn = null;
            try
            {
                cn = new SqlConnection(connectionString);
                //SqlCommand command = new SqlCommand("SELECT content, DirName FROM Docs WHERE LeafName LIKE N'И-Франковск.xls'", cn);
                cn.Open();
                foreach (String filePath in filePaths)
                {
                    Stream myStream = new FileStream(filePath, FileMode.Open);
                    String insertCmd = "insert into Documents values (@CatalogsId, @NameOfReportsSet, @WordDocument, @WordName);";
                    SqlCommand myCommand = new SqlCommand(insertCmd, cn);
                    myCommand.Parameters.Add(new SqlParameter("@WordDocument", SqlDbType.Image));
                    myCommand.Parameters.Add(new SqlParameter("@CatalogsId", SqlDbType.Int));
                    myCommand.Parameters.Add(new SqlParameter("@NameOfReportsSet", SqlDbType.NVarChar));
                    myCommand.Parameters.Add(new SqlParameter("@WordName", SqlDbType.NVarChar));
                    byte[] buf = new byte[myStream.Length];
                    myStream.Read(buf, 0, (int)myStream.Length);
                    myStream.Close();
                    myCommand.Parameters["@WordDocument"].Value = buf;
                    myCommand.Parameters["@CatalogsId"].Value = catalogID;
                    myCommand.Parameters["@NameOfReportsSet"].Value = nameOfReportsSet;
                    myCommand.Parameters["@WordName"].Value = "test";
                    myCommand.Connection.Open();
                    myCommand.ExecuteNonQuery();
                    myCommand.Connection.Close();
                }
            }
            catch (Exception)
            { }
            finally
            {
                if (cn != null)
                    cn.Close();
            }
        }

        /// <summary>
        /// Функция для вывода данных в дерево
        /// </summary>
        /// <param name="tree"></param>
        public static void GetReportsData(TreeView tree)
        {
            SqlConnection cn = null;
            DataTable dt = null;
            try
            {   
                tree.Nodes.Clear();
                cn = new SqlConnection(connectionString);
                SqlDataAdapter adapter = new SqlDataAdapter("select * From dbo.CatalogsReports c Inner join dbo.DocumentsReports docs ON docs.CatalogsId = c.Id ORDER BY c.Name, docs.CreateDate",
                    cn);
                dt = new DataTable();
                adapter.Fill(dt);

                int previousCatalog = -1;
                string currCatalogname = string.Empty;

                TreeNode flyAround = new TreeNode("Временный отчет");
                flyAround.Text = "1 - весна 2011";
                tree.Nodes.Add(flyAround);

                TreeNode cat = null;

                foreach (DataRow dr in dt.Rows)
                {
                    int num = (int)dr.ItemArray[0];
                    if (num != previousCatalog)
                    {
                        previousCatalog = num;
                        cat = new TreeNode((string)dr.ItemArray[3]);
                        cat.Text = cat.Name;
                        flyAround.Nodes.Add(cat);                        
                    }

                    TreeNode reportNode = new TreeNode();
                    reportNode.Text = (string)dr.ItemArray[9];
                    reportNode.Tag = dr.ItemArray[4];
                }
            }
            catch (Exception ex)
            { }
            finally
            {
                if (cn != null)
                    cn.Close();
            }
            //return reader;
        }
    }

        /*
        /// <summary>
        /// Функция для вывода данных в дерево
        /// </summary>
        /// <param name="tree"></param>
        public static void GetReportsData(TreeView tree)
        {
            SqlConnection cn = null;
            DataTable dt = null;
            try
            {   
                
                cn = new SqlConnection(connectionString);
                cn.Open();
                //SqlCommand cmd = new SqlCommand("select * From dbo.CatalogsReports c Inner join dbo.DocumentsReports docs ON docs.CatalogsId = c.Id ORDER BY c.Name, docs.CreateDate", 
                //    cn);
                //"select * From dbo.CatalogsReports"
                SqlDataAdapter adapter = new SqlDataAdapter("select * From dbo.CatalogsReports c Inner join dbo.DocumentsReports docs ON docs.CatalogsId = c.Id ORDER BY c.Name, docs.CreateDate",
                    cn);
                dt = new DataTable();
                adapter.Fill(dt);

                int previousCatalog = -1;
                string currCatalogname = string.Empty;
                DateTime currDate = DateTime.Now;
                TreeNode cat = null;
                TreeNode date = null;

                foreach (DataRow dr in dt.Rows)
                {
                    int num = (int)dr.ItemArray[0];
                    if (num != previousCatalog)
                    {
                        previousCatalog = num;
                        cat = new TreeNode((string)dr.ItemArray[3]);
                        cat.Text = cat.Name;
                        tree.Nodes.Add(cat);
                        currDate = ((DateTime)dr.ItemArray[7]);
                        date = new TreeNode(currDate.ToShortDateString());
                        date.Text = date.Name;
                        cat.Nodes.Add(date);
                    }

                    DateTime dateRow = (DateTime)dr.ItemArray[7];
                    if (dateRow != currDate)
                    {
                        currDate = dateRow;
                        date = new TreeNode(currDate.ToShortDateString());
                        date.Text = date.Name;
                        cat.Nodes.Add(date);
                    }

                    TreeNode reportNode = new TreeNode();
                    reportNode.Text = (string)dr.ItemArray[9];
                    reportNode.Tag = dr.ItemArray[4];
                }
            }
            catch (Exception)
            { }
            finally
            {
                if (cn != null)
                    cn.Close();
            }
            //return reader;
        }
    }
         * */
}
