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

            string fileName = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ExcelData.xls");
            var people = from p in ExcelRepository.GetSheet<Person>(fileName)
                         where p.BirthDate == new DateTime(2008, 9, 10)
                         select p;

            StringBuilder sb = new StringBuilder();
            foreach (Person p in people)
            {
                sb.AppendFormat("Person: {0} {1} Color: {2}, Age: {3}", p.FirstName, p.LastName, p.FavoriteColor, p.Age);
                sb.AppendLine();
            }
            textBox1.Text = sb.ToString();
        }
    }
}
