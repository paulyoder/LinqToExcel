using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using log4net;
using System.Reflection;
using System.IO;
using System.Linq.Expressions;

[assembly: log4net.Config.XmlConfigurator()]

namespace LinqToExcel.Prototype
{
    public partial class Form1 : Form
    {
        private static ILog _log = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //This is used for debugging purposes while building the LinqToExcel library

            string fileName = @"C:\Users\paul.yoder\Desktop\test.xls";
            IExcelRepository repo = new ExcelRepository(fileName);
            var people = from row in repo.Worksheet
                         where row["FirstName"].ToString() == "Paul"
                         select row;
            people.GetEnumerator();

            //IExcelRepository rep = new ExcelRepository(fileName);
            //var people = from p in rep.Worksheet
            //             where p["Age"].ValueAs<int>() > 25
            //             select p;


            //StringBuilder sb = new StringBuilder();
            //foreach (var p in people)
            //{
            //    sb.AppendFormat("{0} {1}", p.FirstName, p.LastName);
            //    sb.AppendLine();
            //}
            //textBox1.Text = sb.ToString();
        }
    }
}
