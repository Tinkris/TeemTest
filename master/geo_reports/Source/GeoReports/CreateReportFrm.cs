using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace GeoReports
{
    public partial class CreateReportFrm : Form
    {
        Manager manager;
        public CreateReportFrm(Manager _manager)
        {
            InitializeComponent();
            manager = _manager;
        }

        private void Fill()
	{
		try
		{
			if (manager != null)
				manager.FillReport();
		}
		catch (Exception ex)
		{
			MessageBox.Show("Не удалось заполнить форму отчета.");
		}
	}
}
