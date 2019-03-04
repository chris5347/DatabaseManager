using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CRMApp
{
    public partial class editColumn : Form
    {
        string colName = "";
        public editColumn(string c)
        {
            InitializeComponent();
            colName = c;
        }

        private void editColumn_Load(object sender, EventArgs e)
        {
            textBox1.Text = colName;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Form1 f = (Form1)Application.OpenForms["Form1"];
            f.editCol(colName,textBox1.Text);
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form1 f = (Form1)Application.OpenForms["Form1"];
            f.deleteCol(colName);
            this.Close();
        }


    }
}
