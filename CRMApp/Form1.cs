using MySql.Data.MySqlClient;
using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Data;
using System.Configuration;
using System.Data.SqlClient;
using System.Data.OleDb;
using Spire.Xls;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
/**
 * CRM --- This program is used for databases. You can store your excel files in the cloud and search
 * @author    Chris Campone
 */
namespace CRMApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        public MySqlConnection connection;
        string username = "";
        string MyConnectionString;
        int flow2ControlWidth = 120;
        bool forceStop = false;
        #region DbStuff

        public bool executeQuery(string q)
        {
            connection.Open();
            MySqlCommand mcd = new MySqlCommand(q, connection);
            if (mcd.ExecuteNonQuery() == 1)
            {
                connection.Close();
                return true;
            }
            else
            {
                connection.Close();
                return false;
            }
        }
        public string dbGetData(String q)
        {
            connection.Open();
            string s = "";
            try
            {
                MySqlCommand cmd1 = new MySqlCommand(q, connection);
                var firstColumn = cmd1.ExecuteScalar();

                if (firstColumn != null)
                {
                    s = firstColumn.ToString();
                }
            }
            catch (Exception e)
            {
                MessageBox.Show("getData: " + e.ToString());
                connection.Close();
                Environment.Exit(1);
            }
            connection.Close();
            return s;
        }
        #endregion


        private void Form1_Load(object sender, EventArgs e)
        {
            MyConnectionString = "Server=MYSQL5005.SmarterASP.NET;Database=db_a2770a_crm;Uid=a2770a_crm;Pwd=chris5347";
            connection = new MySqlConnection(MyConnectionString);
            username = "Admin";
            label7.Text = username;
            loadDatabases();
            dataGridView1.Font = new Font("Tahoma", 12);
            dataGridView1.ForeColor = Color.Black;
            panel3.BringToFront();
            button13.BackColor = Color.FromArgb(35, 168, 109);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                
                if(textBox1.Text.ToLower().Equals("id")){
                    MessageBox.Show("ID is automatically generated - no need to add this column!");
                    textBox1.Text = "";
                    return;
                }
                if(textBox1.Text.Length > 0){
                    addCol(textBox1.Text);
                    textBox1.Text = "";
                }
            }
            catch (Exception x)
            {
                MessageBox.Show(x.ToString());
            }
        }

        void addCol(string colName)
        {
            bool exist = false;
            foreach(Control x in flowLayoutPanel1.Controls){
                if(x.Name.ToLower().Equals("col_"+colName.ToLower())){
                    exist = true;
                }
            }
            if(!exist){
                addColButton(colName);
            }
            else
            {
                MessageBox.Show("Column already exists");
            }
        }

        void textBoxFocus(object sender, EventArgs e)
        {
            TextBox t = (TextBox)sender;
            MessageBox.Show(t.Name);
        }

        void t_KeyDown(object sender, KeyEventArgs e)
        {
            bool updateData = false;
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                //do it if its adding into the database
                foreach (Control item in flowLayoutPanel2.Controls)
                {
                    if(item.Text.ToLower().Equals("save")){
                        updateData = true;
                    }
                }
                if (updateData)
                {
                    saveEdit(null,null);
                }
                else
                {
                    addToDbEdit(null, null);
                }
                // these last two lines will stop the beep sound
                e.SuppressKeyPress = true;
                e.Handled = true;
                
            }
        }

        void addTextBox(Control c,string name,string placeholder)
        {
            TextBox t = new TextBox();
           // l.Name = "label" + i.ToString();
           // t.Margin = new Padding(5);
            t.ForeColor = Color.Black;
            t.Width = flow2ControlWidth;
            t.Name = name;
            t.Text = placeholder;
            if(name.ToLower().Equals("textboxid")){
                t.Enabled = false;
            }
            //t.Click += new System.EventHandler(this.textBoxFocus);
            t.KeyDown += new KeyEventHandler(t_KeyDown);
            c.Controls.Add(t);
        }

        void addLabel(Control c,string txt)
        {
            Label t = new Label();
            // l.Name = "label" + i.ToString();
            t.Margin = new Padding(0);
            t.ForeColor = Color.White;
            t.BackColor = Color.Transparent;
            t.Font = new Font("Serif",10,FontStyle.Regular);
            t.Text = txt;
            //string amount = (flowLayoutPanel1.Controls.Count + 1).ToString();
           // t.Name = "textBox" + amount;
            c.Controls.Add(t);
        }

        void addColButton(string text)
        {
            Button t = new Button();
            // l.Name = "label" + i.ToString();
            t.Margin = new Padding(0);
            t.ForeColor = Color.Black;
            t.BackColor = Color.Transparent;
            t.Text = text;
            t.Name = "col_" + text;
           
            if(text.Equals("ID")){
                t.Enabled = false;
            }
            else
            {
                t.Click += new System.EventHandler(this.colBtnClicked);
            }
            flowLayoutPanel1.Controls.Add(t);
        }

        

        /*void addDatabaseButton(string text)
        {
            Button t = new Button();
            // l.Name = "label" + i.ToString();
            t.Margin = new Padding(8);
            t.ForeColor = Color.Black;
            t.BackColor = Color.Transparent;
            t.Text = text.Substring(text.IndexOf("_")+1);
            t.Name = text;
            t.Click += new System.EventHandler(this.dbBtnClicked);
            flowLayoutPanel2.Controls.Add(t);
        }
        */
       
        void dbBtnClicked(object sender, EventArgs e)
        {
            Button t = (Button)sender;
            MessageBox.Show("name: "+t.Name);
        }


        void colBtnClicked(object sender, EventArgs e)
        {
            Button t = (Button)sender;
            editColumn s = new editColumn(t.Text);
            s.ShowDialog();
        }

        public void editCol(string colOldName,string colNewName)
        {
            foreach (Control x in flowLayoutPanel1.Controls)
            {
                if (x.Name.Equals("col_" + colOldName))
                {
                    x.Name = "col_"+colNewName;
                    x.Text = colNewName;
                  //  MessageBox.Show("Column Updated");
                }
            }
        }

        public void deleteCol(string colName)
        {
            foreach (Control x in flowLayoutPanel1.Controls)
            {
                if (x.Name.Equals("col_" + colName))
                {
                    flowLayoutPanel1.Controls.Remove(x);
                   // MessageBox.Show("Column Deleted");
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (textBox2.Text.Length == 0)
                {
                    MessageBox.Show("Provide a name for your database");
                    return;
                }
                if (flowLayoutPanel1.Controls.Count == 0)
                {
                    MessageBox.Show("Please create soem columns");
                    return;
                }

                string cols = "";
                foreach (Control x in flowLayoutPanel1.Controls)
                {
                    cols += x.Text + " varchar(255),";
                }
                executeQuery("CREATE TABLE " + username + "_" + textBox2.Text + "(ID int NOT NULL AUTO_INCREMENT," + cols + " PRIMARY KEY (ID));");
                MessageBox.Show("Created Database: " + textBox2.Text);
                statusLabel.Text = "Created Database: " + textBox2.Text;
                loadDatabases();
                button3_Click(null, null); //clear controls
                textBox2.Text = "";
                textBox3.Text = "";
            }
            catch (Exception ex)
            {
                MessageBox.Show("Oops! Something went wrong, please try again or use different column names");
                MessageBox.Show(ex.ToString());
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            flowLayoutPanel1.Controls.Clear();
        }

        public void loadDatabases()
        {
            listBox1.Items.Clear();
            string myConnectionString = "Server=MYSQL5005.SmarterASP.NET;Database=db_a2770a_crm;Uid=a2770a_crm;Pwd=chris5347";
            MySqlConnection connection = new MySqlConnection(myConnectionString);
            MySqlCommand command = connection.CreateCommand();
            command.CommandText = "select TABLE_NAME from information_schema.tables where TABLE_SCHEMA='db_a2770a_crm' AND TABLE_NAME LIKE '" + username + "_%'";
            MySqlDataReader Reader;
            connection.Open();
            Reader = command.ExecuteReader();
            while (Reader.Read())
            {
                string row = "";
                for (int i = 0; i < Reader.FieldCount; i++)
                    row += Reader.GetValue(i).ToString();
                listBox1.Items.Add(row.Substring(row.IndexOf("_")+1));
            }
            
        }

        private void textBox1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == System.Windows.Forms.Keys.Enter)
            {
                button1_Click(null, null);
                // these last two lines will stop the beep sound
                e.SuppressKeyPress = true;
                e.Handled = true;
            }
        }
     
        private void loadDB(string dbName)
        {
            try
            {
                MySqlDataAdapter sda = new MySqlDataAdapter();
            DataSet dt = new DataSet();
            clearDataGridView(table);

            dbName = username + "_" + dbName;
            connection = new MySqlConnection(MyConnectionString);
            connection.Open();
            sda = new MySqlDataAdapter("select * from `"+dbName+"` ORDER BY `ID` DESC", connection);
            sda.Fill(dt);
            table.DataSource = dt.Tables[0];
            foreach (DataGridViewRow row in table.Rows)
            {

                table.Rows[0].Selected = false;
            }
            foreach (DataGridViewColumn column in table.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            connection.Close();

            comboBox1.Items.Clear();
            foreach (DataGridViewColumn column in table.Columns)
            {
                
                string head = column.HeaderText;
                comboBox1.Items.Add(head);
            }
            comboBox1.Items.Add("");
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
               
            }
        }

        private void table_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                foreach (DataGridViewRow row in table.SelectedRows)
                {
                    string cpu = row.Cells["Computer"].Value.ToString();
                    DialogResult dr = MessageBox.Show("Delete Cpu Name: " + cpu, "", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Warning);
                    if (dr == DialogResult.Cancel)
                    {
                        return;
                    }
                    else if (dr == DialogResult.Yes)
                    {

                        string q = "DELETE FROM `instagram` WHERE Computer='" + cpu + "'";
                        executeQuery(q);
                      //  dt.Clear();
                      //  Admin_Load(null, null);
                        return;
                    }
                }
            }
            else
            {
                flowLayoutPanel2.Controls.Clear();
                foreach (DataGridViewRow row in table.SelectedRows)
                {
                    foreach (DataGridViewColumn column in table.Columns)
                    {

                        string head = column.HeaderText;
                        addLabel(flowLayoutPanel2, head);
                        addTextBox(flowLayoutPanel2,"textBox"+head,row.Cells[head].Value.ToString());
                    }
                    
                }
                if(flowLayoutPanel2.Controls.Count > 0){
                    Button t = new Button();
                    t.ForeColor = Color.Black;
                    t.BackColor = Color.Transparent;
                    t.FlatStyle = FlatStyle.Standard;
                    t.Width = flow2ControlWidth;
                    t.Text = "Save";
                    t.Name = "saveBtn";

                    t.Click += new System.EventHandler(this.saveEdit);
                    flowLayoutPanel2.Controls.Add(t);
                    t = new Button();
                    t.ForeColor = Color.Black;
                    t.BackColor = Color.Transparent;
                    t.FlatStyle = FlatStyle.Standard;
                    t.Width = flow2ControlWidth;
                    t.Text = "Delete";
                    t.Name = "deleteBtn";

                    t.Click += new System.EventHandler(this.deleteDataFromDbEdit);
                    flowLayoutPanel2.Controls.Add(t);
                }
            }
        }

        public void deleteDataFromDbEdit(object sender, EventArgs e)
        {
            string id = "";
            foreach (Control c in flowLayoutPanel2.Controls)
            {
                if (c.Name.Equals("textBoxID"))
                {
                    id = c.Text;
                }
            }
            string db = listBox1.GetItemText(listBox1.SelectedItem);
            executeQuery("DELETE FROM "+username+"_"+db+" WHERE `ID`='"+id+"'");
            flowLayoutPanel2.Controls.Clear();
            loadDB(db);
        }

        public void saveEdit(object sender, EventArgs e)
        {
            try
            {
                ArrayList values = new ArrayList();
                ArrayList cols = new ArrayList();
                string id = "";
                foreach (Control c in flowLayoutPanel2.Controls)
                {
                    if (c.Name.Contains("textBox"))
                    {
                        values.Add(c.Text);
                    }
                }

                foreach (DataGridViewColumn column in table.Columns)
                {
                    cols.Add(column.HeaderText);
                }
                //make string
                //update tablename set col1=value1,col2=value2 where condition
                string set = "";
                for (int i = 0; i < cols.Count; i++)
                {
                    if (set.Length == 0)
                    {
                        set += "`" + cols[i] + "`='" + values[i] + "'";
                    }
                    else
                    {
                        set += ", `" + cols[i] + "`='" + values[i] + "'";
                    }

                }
                string db = listBox1.GetItemText(listBox1.SelectedItem);
                executeQuery("UPDATE `" + username.ToLower() + "_" + db + "` SET " + set + " WHERE `ID`='" + values[0] + "'");
                loadDB(db);
                statusLabel.Text = "Data updated in database: " + db;
                //table.Rows[0].Selected = true;
                foreach (DataGridViewRow row in table.Rows)
                {
                    if (row.Cells[0].Value.ToString().Equals(values[0].ToString()))
                    {
                        table.Rows[row.Index].Selected = true;
                        table.FirstDisplayedScrollingRowIndex = table.SelectedRows[0].Index;
                        break;
                    }
                }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show("Couldnt save data");
            }
        }

        private void tabControl1_Selecting(object sender, TabControlCancelEventArgs e)
        {
           /* if (tabControl1.SelectedTab.Text.ToLower().Equals("your database"))
            {
                tabControl1.Width = 1159;
                this.Width = 1199;
            }
            else
            {
                tabControl1.Width = 793;
                this.Width = 840;
            }
            */
            this.Location = new Point((Screen.PrimaryScreen.Bounds.Size.Width / 2) - (this.Size.Width / 2), (Screen.PrimaryScreen.Bounds.Size.Height / 2) - (this.Size.Height / 2));  

        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                if (listBox1.SelectedItem == null)
                {
                    MessageBox.Show("Select a database");
                    return;
                }
                string db = listBox1.GetItemText(listBox1.SelectedItem);
                executeQuery("DROP TABLE " + username + "_" + db);
                statusLabel.Text = "Deleted table: "+db;
                loadDatabases();
            }
            catch (Exception x)
            {
                MessageBox.Show(x.ToString());
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedItem == null)
            {
                MessageBox.Show("Select a database");
                return;
            }
            exportToExcel();
        }

        private void exportToExcel()
        {
            string db = listBox1.GetItemText(listBox1.SelectedItem);
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);

            MySqlConnection conn = new MySqlConnection("Server=MYSQL5005.SmarterASP.NET;Database=db_a2770a_crm;Uid=a2770a_crm;Pwd=chris5347");
            conn.Open();
            MySqlCommand cmd = new MySqlCommand("select * from "+username+"_"+db, conn);
            MySqlDataReader dr = cmd.ExecuteReader();

            using (System.IO.StreamWriter fs = new System.IO.StreamWriter(path + @"\Database-" + db + ".csv"))
            {
                // Loop through the fields and add headers
                for (int i = 0; i < dr.FieldCount; i++)
                {
                    string name = dr.GetName(i);
                    if (name.Contains(","))
                        name = "\"" + name + "\"";

                    fs.Write(name + ",");
                }
                fs.WriteLine();

                // Loop through the rows and output the data
                while (dr.Read())
                {
                    for (int i = 0; i < dr.FieldCount; i++)
                    {
                        string value = dr[i].ToString();
                        if (value.Contains(","))
                            value = "\"" + value + "\"";

                        fs.Write(value + ",");
                    }
                    fs.WriteLine();
                }

                fs.Close();
            }
            MessageBox.Show("File Saved To Desktop: Database-"+db);
            statusLabel.Text = "File Saved To Desktop: Database-" + db;

        }

        private void listBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string db = listBox1.GetItemText(listBox1.SelectedItem);
            flowLayoutPanel2.Controls.Clear();
            if (db.Length == 0)
            {
                MessageBox.Show("Please select a database");
            }
            else
            {
                loadDB(db);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if(listBox1.SelectedItem == null){
                    MessageBox.Show("Select a database");
                    return;
                }
                flowLayoutPanel2.Controls.Clear();
                foreach (DataGridViewColumn column in table.Columns)
                {

                    string head = column.HeaderText;
                    if (!head.ToLower().Equals("id"))
                    {
                        addLabel(flowLayoutPanel2, head);
                        addTextBox(flowLayoutPanel2, "textBox" + head, "");
                    }

                }
                Button t = new Button();
                //t.Margin = new Padding(5);
                t.Width = flow2ControlWidth;
                t.ForeColor = Color.Black;
                t.BackColor = Color.Transparent;
                t.FlatStyle = FlatStyle.Standard;
                t.Text = "Add To Database";
                t.Name = "saveBtn";

                t.Click += new System.EventHandler(this.addToDbEdit);
                flowLayoutPanel2.Controls.Add(t);
            }
            catch (Exception x)
            {
                MessageBox.Show(x.ToString());
            }
        }

        void addToDbEdit(object sender, EventArgs e)
        {
            try
            {
                ArrayList values = new ArrayList();
                ArrayList cols = new ArrayList();
                foreach (Control c in flowLayoutPanel2.Controls)
                {
                    if (c.Name.Contains("textBox") && !c.Name.Equals("textBoxID"))
                    {
                        values.Add(c.Text);
                    }
                }

                foreach (DataGridViewColumn column in table.Columns)
                {
                    if (!column.HeaderText.Equals("ID"))
                    {
                        cols.Add(column.HeaderText);
                    }
                    
                }
                //make string
                //inert into `database`(`ID`,`Name`) VALUES('1','name')
                string colString = "";
                string vals = "";
                for (int i = 0; i < cols.Count; i++)
                {
                    if (colString.Length == 0)
                    {
                        colString += "`" + cols[i] + "`";
                    }
                    else
                    {
                        colString += ",`" + cols[i] + "`";
                    }

                }
                for (int i = 0; i < values.Count; i++)
                {
                    if (vals.Length == 0)
                    {
                        vals += "'" + values[i] + "'";
                    }
                    else
                    {
                        vals += ",'" + values[i] + "'";
                    }

                }

                string db = listBox1.GetItemText(listBox1.SelectedItem);
                string sql = "INSERT INTO `" + username.ToLower() + "_" + db + "` (" + colString + ") VALUES(" + vals + ")";
               // MessageBox.Show(sql);
                executeQuery(sql);
                loadDB(db);
                statusLabel.Text = "Data inserted into database: "+db;
                foreach (Control item in flowLayoutPanel2.Controls)
                {
                    if(item.Name.Contains("textBox")){
                        item.Text = "";
                    }
                }
            }
            catch (Exception x)
            {
                MessageBox.Show(x.ToString());
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            button10.Visible = true;
            uploadExcelToGridView();
            
            
        }

        string temppp = "";

       async private void uploadExcelFileToDB()
        {
            if (dataGridView1.Columns.Count == 0)
            {
                MessageBox.Show("Please upload an excel file");
                return;
            }
            if (textBox3.Text.Length == 0)
            {
                MessageBox.Show("Give the database a name");
                return;
            }
            if (Regex.IsMatch(textBox3.Text, @"^\d") || textBox3.Text.Contains(" "))
            {
                MessageBox.Show("Database name is invalid. Please choose another");
                return;
            }


            try
            {
                progressBar1.Value = 0;
                button8.Enabled = false;
                ArrayList colsArray = new ArrayList();
                string cols = "";
                foreach (DataGridViewColumn column in dataGridView1.Columns)
                {
                    string head = column.HeaderText;
                    if (head.ToLower().Equals("id"))
                    {
                        MessageBox.Show("Please rename your column: " + head);
                        clearDataGridView(dataGridView1);
                        return;
                    }
                    //get headers
                    if(head.Contains(" ")){
                        head = "";
                        for (int i = 0; i < column.HeaderText.Count(); i++)
                        {
                            if (column.HeaderText.Substring(i,1).Equals(" "))
                            {
                                head += "_";
                            }
                            else
                            {
                                head += column.HeaderText.Substring(i, 1);
                            }
                        }
                    }
                    cols += head + " varchar(255),";
                    colsArray.Add(head);
                }
                
                //make string
                //insert into `database`(`ID`,`Name`) VALUES('1','name')
                string colString = "";
                string vals = "";
                for (int i = 0; i < colsArray.Count; i++)
                {
                    if (colString.Length == 0)
                    {
                        colString += "`" + colsArray[i] + "`";
                    }
                    else
                    {
                        colString += ",`" + colsArray[i] + "`";
                    }

                }
                //loop thru each row of gridview
               
                string db = textBox3.Text;
                int rowCount = dataGridView1.Rows.Count;
                vals = "";
                foreach (DataGridViewRow item in dataGridView1.Rows)
                {
                    //get values
                    if(forceStop){
                        button7.Enabled = true;
                        button10.Visible = false;
                        statusLabel.Text = "Canceled your database upload to server";
                        return; 
                    }
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        //must make this string
                        //( Value1, Value2,val 3 ), ( Value1, Value2, val3 )
                        
                        
                        if(i==0 && dataGridView1.Columns.Count>1){
                            if (item.Cells[i].Value == null || item.Cells[i].Value == DBNull.Value || String.IsNullOrWhiteSpace(item.Cells[i].Value.ToString()))
                            {
                                vals += "(' ',";
                            }
                            else
                            {
                                vals += "('" + item.Cells[i].Value.ToString() + "',";
                            }
                            
                            continue;
                        }
                        else if (i == 0 && dataGridView1.Columns.Count == 1 && item.Index != dataGridView1.Rows.Count - 1)
                        {
                            if (item.Cells[i].Value == null || item.Cells[i].Value == DBNull.Value || String.IsNullOrWhiteSpace(item.Cells[i].Value.ToString()))
                            {
                                vals += "(' '),";
                            }
                            else
                            {
                                vals += "('" + item.Cells[i].Value.ToString() + "'),";
                            }
                            
                            continue;
                        }
                        else if (i == 0 && dataGridView1.Columns.Count == 1 && item.Index == dataGridView1.Rows.Count - 1)
                        {
                            if (item.Cells[i].Value == null || item.Cells[i].Value == DBNull.Value || String.IsNullOrWhiteSpace(item.Cells[i].Value.ToString()))
                            {
                                vals += "(' ')";
                            }
                            else
                            {
                                vals += "('" + item.Cells[i].Value.ToString() + "')";
                            }
                            
                            continue;
                        }

                        //not the last one or first one 
                        if (i != dataGridView1.Columns.Count - 1 && i != 0)
                        {
                            if (item.Cells[i].Value == null || item.Cells[i].Value == DBNull.Value || String.IsNullOrWhiteSpace(item.Cells[i].Value.ToString()))
                            {
                                vals += "' ',"; ;
                            }
                            else
                            {
                                vals += "'" + item.Cells[i].Value.ToString() + "',";
                            }
                            
                            continue;
                        }

                        //last row and last cell
                        if (i == dataGridView1.Columns.Count - 1 && item.Index == dataGridView1.Rows.Count - 1)
                        {
                            if (item.Cells[i].Value == null || item.Cells[i].Value == DBNull.Value || String.IsNullOrWhiteSpace(item.Cells[i].Value.ToString()))
                            {
                                vals += "' ')";
                            }
                            else
                            {
                                vals += "'" + item.Cells[i].Value.ToString() + "')";
                            }
                            
                            continue;
                        }
                        //last cell 
                        if (i == dataGridView1.Columns.Count - 1 )
                        {
                            if (item.Cells[i].Value == null || item.Cells[i].Value == DBNull.Value || String.IsNullOrWhiteSpace(item.Cells[i].Value.ToString()))
                            {
                                vals += "' '),";
                            }
                            else
                            {
                                vals += "'" + item.Cells[i].Value.ToString() + "'),";
                            }
                            
                            continue;
                        }
                        

                    }
                    progressBar1.Value += 1;
                    statusLabel.Text = "Gathering Information: "+item.Index+"/"+dataGridView1.Rows.Count;
                    await Task.Delay(1);
                }
                string sql = "INSERT INTO `" + username.ToLower() + "_" + db + "` (" + colString + ") VALUES " + vals + "";
               
                //creates the database
              //  MessageBox.Show("CREATE TABLE " + username + "_" + textBox3.Text + "(ID int NOT NULL AUTO_INCREMENT," + cols + " PRIMARY KEY (ID));");
                executeQuery("CREATE TABLE " + username + "_" + textBox3.Text + "(ID int NOT NULL AUTO_INCREMENT," + cols + " PRIMARY KEY (ID));");
                executeQuery(sql);
                statusLabel.Text = "Created Database: " + textBox3.Text;
                loadDatabases();
                clearDataGridView(dataGridView1);
                textBox3.Text = "";
                button10.Visible = false;
                
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString());
            }


        }

        async private void uploadExcelToGridView()
        {
            try
            {
                progressBar1.Value = 0;
                button7.Enabled = false;
                string fname = "";
                OpenFileDialog fdlg = new OpenFileDialog();
                fdlg.Title = "Excel File Dialog";
                // fdlg.InitialDirectory = @"c:\";
                fdlg.Filter = "All files (*.*)|*.*|All files (*.*)|*.*";
                fdlg.FilterIndex = 2;
                fdlg.RestoreDirectory = true;
                if (fdlg.ShowDialog() == DialogResult.OK)
                {
                    fname = fdlg.FileName;
                }
                else
                {
                    button7.Enabled = true;
                    button10.Visible = false;
                    return;
                }


                Microsoft.Office.Interop.Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fname);
                Microsoft.Office.Interop.Excel._Worksheet xlWorksheet = xlWorkbook.Sheets[1];
                Microsoft.Office.Interop.Excel.Range xlRange = xlWorksheet.UsedRange;

                int rowCount = xlRange.Rows.Count;
                int colCount = xlRange.Columns.Count;

                // dt.Column = colCount;  
                dataGridView1.ColumnCount = colCount;
                dataGridView1.RowCount = rowCount;

                progressBar1.Maximum = rowCount;
                for (int i = 1; i <= rowCount; i++)
                {
                    if(forceStop){
                        forceStop = false;
                        clearDataGridView(dataGridView1);
                        button7.Enabled = true;
                        button10.Visible = false;
                        statusLabel.Text = "Canceled your file upload";
                        return;
                    }
                    button7.Enabled = false;
                    for (int j = 1; j <= colCount; j++)
                    {


                        //write the value to the Grid  
                        

                        if (xlRange.Cells[i, j] != null && xlRange.Cells[i, j].Value2 != null)
                        {
                            if (xlRange.Cells[i, j].Value2.ToString().Contains(","))
                            {
                                string xx = "";
                                for (int w = 0; w < xlRange.Cells[i, j].Value2.ToString().Length; w++)
                                {
                                    if (xlRange.Cells[i, j].Value2.ToString().Substring(w, 1).Equals(","))
                                    {
                                        xx += "_";
                                    }
                                    else
                                    {
                                        xx += xlRange.Cells[i, j].Value2.ToString().Substring(w, 1);
                                    }
                                }
                                dataGridView1.Rows[i - 1].Cells[j - 1].Value = xx;
                            }
                            else
                            {
                                dataGridView1.Rows[i - 1].Cells[j - 1].Value = xlRange.Cells[i, j].Value2.ToString();
                            }
                            
                        }
                        // Console.Write(xlRange.Cells[i, j].Value2.ToString() + "\t");  

                        //add useful things here!   
                        await Task.Delay(05);
                        
                    }
                    progressBar1.Value += 1;
                    statusLabel.Text = "Downloading excel file. Rows to left complete: "+(rowCount-i).ToString();
                    
                }
                button7.Enabled = true;
                statusLabel.Text = "Loaded " + dataGridView1.Rows.Count.ToString() +" rows of data. Ready to create the database";
                //cleanup  
                GC.Collect();
                GC.WaitForPendingFinalizers();

                //rule of thumb for releasing com objects:  
                //  never use two dots, all COM objects must be referenced and released individually  
                //  ex: [somthing].[something].[something] is bad  

                //release com objects to fully kill excel process from running in the background  
                Marshal.ReleaseComObject(xlRange);
                Marshal.ReleaseComObject(xlWorksheet);

                //close and release  
                xlWorkbook.Close();
                Marshal.ReleaseComObject(xlWorkbook);
                
                //quit and release  
                xlApp.Quit();
                Marshal.ReleaseComObject(xlApp);

                //first row should be col headers
                int colCount2 = dataGridView1.Columns.Count;
                for (int i = 0; i < colCount2; i++)
                {
                    dataGridView1.Columns[i].HeaderText = dataGridView1.Rows[0].Cells[i].Value.ToString();
                }
                dataGridView1.Rows.RemoveAt(0);
                button8.Enabled = true;
                button10.Visible = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show("Couldnt read the excel file: "+ex.ToString());
                button7.Enabled = true;
                button8.Enabled = false;
                clearDataGridView(dataGridView1);
            }
        }

        private void clearDataGridView(DataGridView x)
        {
            x.DataSource = null;
            x.Rows.Clear();
            x.Columns.Clear();
            x.ClearSelection();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            button10.Visible = true;
            uploadExcelFileToDB();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            

            string search = textBox4.Text;
            string db = listBox1.GetItemText(listBox1.SelectedItem);

            string where = "";
            if(comboBox1.GetItemText(comboBox1.SelectedItem).Length == 0){
                foreach (DataGridViewColumn column in table.Columns)
                {

                    string head = column.HeaderText;
                    if (where.Length == 0)
                    {
                        where += head + " LIKE '%" + search + "%'";
                    }
                    else
                    {
                        where += " OR " + head + " LIKE '%" + search + "%'";
                    }

                }
            }else{
                string filter = comboBox1.GetItemText(comboBox1.SelectedItem);
                where = filter + " LIKE '%" + search + "%'";
            }

            MySqlDataAdapter sda = new MySqlDataAdapter();
            DataSet dt = new DataSet();
            clearDataGridView(table);
            connection = new MySqlConnection(MyConnectionString);
            connection.Open();
            sda = new MySqlDataAdapter("SELECT * FROM " + username + "_" + db + " WHERE " + where, connection);
            sda.Fill(dt);
            table.DataSource = dt.Tables[0];
            foreach (DataGridViewRow row in table.Rows)
            {

                table.Rows[0].Selected = false;
            }
            foreach (DataGridViewColumn column in table.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            connection.Close();
            
        }

        private void button10_Click(object sender, EventArgs e)
        {
            forceStop = true;
        }

       

        private void changeTab(Button btn, int tab)
        {
            foreach (Control item in panel5.Controls)
            {
                item.BackColor = panel1.BackColor;
            }
            btn.BackColor = Color.FromArgb(35, 168, 109);
            tabControl1.SelectedIndex = tab - 1;

        }

       

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            try
            {
                connection.Close();
            }
            catch (Exception)
            {
                
            }
            Application.Exit();
        }

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [DllImportAttribute("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [DllImportAttribute("user32.dll")]
        public static extern bool ReleaseCapture();
        private void panel3_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            changeTab((Button)sender, 3);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            changeTab((Button)sender, 2);
        }

        private void button13_Click(object sender, EventArgs e)
        {
            changeTab((Button)sender, 1);
        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {
            if (!this.WindowState.ToString().Equals("Normal"))
            {
                this.WindowState = FormWindowState.Normal;
            }
            else
            {
                this.WindowState = FormWindowState.Maximized;
            }
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }

       

        
       
        

        
        

        

        

    }
}
