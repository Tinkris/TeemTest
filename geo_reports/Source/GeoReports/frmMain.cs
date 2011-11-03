using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MapViewLibrary;

namespace GeoReports
{
    public partial class frmMain : Form
    {
        private Manager manager;
        public frmMain()
        {
            InitializeComponent();
            manager = new Manager(this);
        }

        public void SetViewer(IMapViewer _mapViewer)
        {
            try
            {
                if (manager == null)
                    manager = new Manager(this);
                manager.SetViewer(_mapViewer);
                UpdateTree();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                    "Невозможно установить MapViewer. Работа программы будет прекращена.",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                this.Close();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            manager.WriteG("C://result.doc");
        }


        private void CreateTable6()
        {
            manager.WriteTable6("C://result6.doc");
        }

        private void button2_Click(object sender, EventArgs e)
        {
            CreateTable6();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            manager.WriteTable11("C://result11.doc");
        }

        private void splitContainer1_SplitterMoved(object sender, SplitterEventArgs e)
        {

        }

        private void историяToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void всеОтчетыToolStripMenuItem_Click(object sender, EventArgs e)
        {
            CreateReportFrm frm = new CreateReportFrm(manager);
            frm.Visible = true;
        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {

        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void RestoreSettings()
        {
            SetTopMostState(Properties.Settings.Default.TopMost);
        }

        private void SaveSettings()
        {
            Properties.Settings.Default.TopMost = TopMost;
            Properties.Settings.Default.Save();
        }

        private void tsbtnTopMost_Click(object sender, EventArgs e)
        {
            ChangeTopMostState();
        }

        private void всегдаСверхуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            ChangeTopMostState();
        }

        private void SetTopMostState(bool isTopMost)
        {
            TopMost = isTopMost;
            tsmiTopMost.Checked = TopMost;
            if (isTopMost)
            {
                tsbtnTopMost.Image = Properties.Resources.Pin2_icon;
                tsbtnTopMost.ToolTipText = "Отключить режим 'Всегда сверху'";
            }
            else
            {
                tsbtnTopMost.Image = Properties.Resources.pin_icon;
                tsbtnTopMost.ToolTipText = "Включить режим 'Всегда сверху'";
            }
        }

        private void ChangeTopMostState()
        {
            try
            {               
                TopMost = !TopMost;
                tsmiTopMost.Checked = TopMost;
                if (TopMost)
                {
                    tsbtnTopMost.Image = Properties.Resources.Pin2_icon;
                    tsbtnTopMost.ToolTipText = "Отключить режим 'Всегда сверху'";
                }
                else
                {
                    tsbtnTopMost.Image = Properties.Resources.pin_icon;
                    tsbtnTopMost.ToolTipText = "Включить режим 'Всегда сверху'";
                }
            }
            catch (Exception ex)
            { }
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            RestoreSettings();
            //UpdateTree();
        }

        private void UpdateTree()
        {
            try
            {
                if (manager != null)
                    manager.UpdateTree(treeReports);
                if (treeReports.Nodes.Count == 0)
                {
                    treeReports.Visible = false;
                }
                else
                    treeReports.Visible = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show( "При подключении к базе данных \"Отчеты\" произошла следующая ошибка:\n" + ex.Message,
                    "Ошибка подключения к базе данных",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                treeReports.Nodes.Clear();
            }
        }
    }
}
