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
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.ShowDialog();
            var people = from p in ExcelRepository.GetSheet<Person>(openFile.FileName)
                         where p.BirthDate == new DateTime(2008, 10, 9)
                         select p;

            StringBuilder sb = new StringBuilder();
            foreach (Person p in people)
            {
                sb.AppendFormat("Person: {0} {1} Color: {2}, Age: {3}", p.FirstName, p.LastName, p.FavoriteColor, p.Age);
                sb.AppendLine();
            }
            textBox1.Text = sb.ToString();

            /*
            string color = "Red";
            ExcelSheet<Person> data = new ExcelSheet<Person>();
            var redLovers = from p in data
                            where p.FavoriteColor == color
                            select p;

            redLovers.GetEnumerator();
            /*
            foreach (var lover in redLovers)
            {
                Console.WriteLine("Name: " + lover.FirstName);
            }
             * */
        }
    }
}
