using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace QBTest
{
    public partial class QBTest : Form
    {
        QBConnect qb;
        private string companyFile;
        public QBTest()
        {
            InitializeComponent();
        }

        private void QBTest_Load(object sender, EventArgs e)
        {

            string myExeDir = new FileInfo(Assembly.GetEntryAssembly().Location).Directory.ToString();
            string file = Path.Combine(myExeDir, "AutoBody.ini");
            if (File.Exists(file))
            {
                string fl = File.ReadAllText(file);
                companyFile = fl;
            }
            qb = new QBConnect(companyFile);
            LoadQBdata();
        }

        private async void LoadQBdata()
        {
            showStatus("Loading customers...");
            string[] cusomers = await qb.loadCustomers();
            comboBox_Customer.Items.AddRange(cusomers);
            showStatus("Loading items...");
            string[] items = await qb.loadItems();
            items = ChectList(items);
            comboBox_Item1.Items.AddRange(items);
            comboBox_Item2.Items.AddRange(items);
            comboBox_Item3.Items.AddRange(items);
            comboBox_Item4.Items.AddRange(items);
            comboBox_Item5.Items.AddRange(items);
            showStatus("Loading terms...");
            string[] terms = await qb.loadTerms();
            comboBox_Terms.Items.AddRange(terms);
            showStatus("Loading sales tax codes...");
            string[] taxs = await qb.loadSalesTaxCodes();
            comboBox_Tax1.Items.AddRange(taxs);
            comboBox_Tax2.Items.AddRange(taxs);
            comboBox_Tax3.Items.AddRange(taxs);
            comboBox_Tax4.Items.AddRange(taxs);
            comboBox_Tax5.Items.AddRange(taxs);
            showStatus("Loading customer msgs...");

            showStatus("Ready");
        }
        private string[] ChectList(string[] strings)
        {
            List<string> r = new List<string>();
            string[] res = new string[0];
            foreach(string s in strings)
            {
                if (!String.IsNullOrEmpty(s)) r.Add(s);
            }
            res = new string[r.Count];
            for(int i = 0; i < r.Count; i++)
            {
                res[i] = r[i];
            }
            return res;
        }
        private void showStatus(string text)
        {
            labelLoadStatus.Visible = true;
            labelLoadStatus.Text = text;
            this.Refresh();
        }

        private async void comboBox_Customer_SelectedIndexChanged(object sender, EventArgs e)
        {
            showStatus("Loading customer info...");
            string str = await qb.getBillShipTo(comboBox_Customer.Text, "BillAddress");
            label_BillTo.Text = str.Trim();
            str = await qb.getBillShipTo(comboBox_Customer.Text, "ShipAddress");
            label_ShipTo.Text = str;
        }
    }
}
