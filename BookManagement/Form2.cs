using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MySql.Data;
using MySql.Data.MySqlClient;
using System.Data.OleDb;
using System.IO;

namespace BookManagement
{
    public partial class Form2 : Form
    {
        public MySqlConnection conn;
        private DataTable data;
        private MySqlDataAdapter da;
        //private System.Windows.Forms.DataGridView dataGrid;
       // private MySqlCommandBuilder cb;
        private string connStr,dataChoose;
        public Form2(string c1)
        {
            InitializeComponent();
            connStr = c1;
            Connection();
            GetDatabases();
            
        }
        private void Connection()
        {
            if (conn != null)
                conn.Close();
            conn = new MySqlConnection(connStr);
            conn.Open();
        }
        private void GetDatabases()
        {

            MySqlDataReader reader = null;
            MySqlCommand cmd = new MySqlCommand("SHOW DATABASES", conn);
            try
            {
                reader = cmd.ExecuteReader();
                comboBox1.Items.Clear();
                while (reader.Read())
                {
                    comboBox1.Items.Add(reader.GetString(0));
                }
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Failed to populate database list: " + ex.Message);
            }
            finally
            {
                if (reader != null) reader.Close();
            }
        }
        private void Select(string sqlselect,DataGridView d1)
        {
            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
                conn.Open();
                if (comboBox1.SelectedItem != null)
                {
                    conn.ChangeDatabase(comboBox1.SelectedItem.ToString());
                    data = new DataTable();
                    da = new MySqlDataAdapter(sqlselect, conn);
                    da.Fill(data);
                    if (d1.DataSource != null)
                        d1.DataSource = null;
                    d1.DataSource = data;
                }
                else MessageBox.Show("数据库不能为空");
                conn.Close();
            }
            catch (MySqlException ex)
            {

                MessageBox.Show("Error: " + ex);
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            MySqlDataReader reader = null;
            if (conn.State == ConnectionState.Closed)
            {
                conn.Open();
            }
            conn.ChangeDatabase(comboBox1.SelectedItem.ToString());

            MySqlCommand cmd = new MySqlCommand("SHOW TABLES", conn);
            try
            {
                reader = cmd.ExecuteReader();
                comboBox2.Items.Clear();
                while (reader.Read())
                {
                    comboBox2.Items.Add(reader.GetString(0));
                }
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Failed to populate table list: " + ex.Message);
            }
            finally
            {
                if (reader != null) reader.Close();
                conn.Close();
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                textBox1.Enabled = false;
                textBox2.Enabled = true;
                textBox2.Focus();
            }
            else
            {
                textBox1.Enabled = true;
                textBox2.Enabled = false;
                textBox1.Focus();
            }            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            IfDataChoose();
            if (radioButton2.Checked)
            {
                string sqlselect1 = "SELECT * FROM " + dataChoose + " WHERE ISBN='" + textBox2.Text + "'";
                Select(sqlselect1,dataGridView1);
            }
            else
            {
                string sqlselect2 = "SELECT * FROM " + dataChoose + " WHERE 书名 LIKE '%" + textBox1.Text + "%'";
                Select(sqlselect2,dataGridView1);
            }            
        }

        private void button3_Click(object sender, EventArgs e)
        {
            IfDataChoose();
            string sqlselect2 = "SELECT * FROM " + dataChoose;
            Select(sqlselect2,dataGridView1);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            dataChoose = comboBox2.SelectedItem.ToString();
            conn.Close();
            if (conn.State == ConnectionState.Closed)
            {
                MessageBox.Show("选择数据库成功");
                comboBox1.Enabled = false;
                comboBox2.Enabled = false;
            }
            else MessageBox.Show("选择数据库失败");
        }

        private void button4_Click(object sender, EventArgs e) //入库
        {
            if (IfInfoFilled()&IfDataChoose())
            {
                try
                {
                    if (conn.State == ConnectionState.Open)
                        conn.Close();
                    conn.Open();
                    if (comboBox1.SelectedItem != null)
                        conn.ChangeDatabase(comboBox1.SelectedItem.ToString());
                    data = new DataTable();
                    da = new MySqlDataAdapter("SELECT * FROM " + dataChoose + " WHERE ISBN='" + textBox3.Text + "'"+"AND 库位='"+textBox6.Text+"'", conn);
                    da.Fill(data);    //查询有无记录
                    if (data.Rows.Count == 0)//无记录
                    {
                        string query = "INSERT INTO " + dataChoose + " (书名, ISBN, 数量, 库位, 备注) VALUES('" + textBox4.Text + "', '" + textBox3.Text + "', " + textBox5.Text + ", '" + textBox6.Text + "', '" + textBox7.Text + "')";
                        MySqlCommand cmd = new MySqlCommand(query, conn);
                        cmd.ExecuteNonQuery();
                        da = new MySqlDataAdapter("SELECT * FROM " + dataChoose + " WHERE ISBN='" + textBox3.Text + "'" + "AND 库位='" + textBox6.Text + "'", conn);
                        da.Fill(data);
                        dataGridView2.DataSource = data;
                        conn.Close();
                        label6.Text = "入库成功，入库信息为：";
                    }
                    else                                          //有记录
                    {
                        data.PrimaryKey = new DataColumn[] { data.Columns["ISBN"],data.Columns["库位"]};
                        object[] key = { textBox3.Text,textBox6.Text};
                        DataRow dr = data.Rows.Find(key);
                        int n = (int)dr["数量"];
                        n = n + int.Parse(textBox5.Text);
                        string query = "UPDATE " + dataChoose + " SET 数量='" + n + "' WHERE ISBN='" + textBox3.Text + "'" + "AND 库位='" + textBox6.Text + "'";
                        MySqlCommand cmd = new MySqlCommand(query, conn);
                        cmd.ExecuteNonQuery();
                        da = new MySqlDataAdapter("SELECT * FROM " + dataChoose + " WHERE ISBN='" + textBox3.Text + "'" + "AND 库位='" + textBox6.Text + "'", conn);
                        da.Fill(data);
                        dataGridView2.DataSource = data;
                        conn.Close();
                        label6.Text = "入库成功，入库信息为：";
                    }
                }
                catch (MySqlException ex)
                {

                    MessageBox.Show("Error: " + ex);
                }
            }
            
        }


        private void textBox3_KeyPress(object sender, KeyPressEventArgs e) //ISBN检查
        {
            if (!char.IsNumber(e.KeyChar)&&!char.IsControl(e.KeyChar))
            {
                e.Handled = true;
            }
        }

        private void button5_Click(object sender, EventArgs e) //出库
        {
            if (IfDataChoose()&IfInfoFilled())
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
                conn.Open();
                if (comboBox1.SelectedItem != null)
                    conn.ChangeDatabase(comboBox1.SelectedItem.ToString());
                data = new DataTable();
                da = new MySqlDataAdapter("SELECT * FROM " + dataChoose + " WHERE ISBN='" + textBox3.Text + "'" + "AND 库位='" + textBox6.Text + "'", conn);
                da.Fill(data);    //查询有无记录
                try
                {
                    if (data.Rows.Count == 0)
                    {
                        MessageBox.Show("无该条记录，请确认");
                        conn.Close();
                    }
                    else
                    {
                        data.PrimaryKey = new DataColumn[] { data.Columns["ISBN"], data.Columns["库位"] };
                        object[] key = { textBox3.Text, textBox6.Text };
                        DataRow dr = data.Rows.Find(key);
                        int n = (int)dr["数量"];
                        n = n - int.Parse(textBox5.Text);
                        if (n < 0)
                        {
                            MessageBox.Show("该库位库存不足！");
                            conn.Close();
                        }
                        if (n == 0)
                        {
                            string query = "DELETE FROM " + dataChoose + " WHERE ISBN='" + textBox3.Text + "'" + "AND 库位='" + textBox6.Text + "'";
                            MySqlCommand cmd = new MySqlCommand(query, conn);
                            cmd.ExecuteNonQuery();
                            label6.Text = "出库成功 该库位已经没有库存";
                            conn.Close();
                            dataGridView2.DataSource = null;
                            LogAdd();
                        }
                        if (n > 0)
                        {
                            string query = "UPDATE " + dataChoose + " SET 数量='" + n + "' WHERE ISBN='" + textBox3.Text + "'" + "AND 库位='" + textBox6.Text + "'";
                            MySqlCommand cmd = new MySqlCommand(query, conn);
                            cmd.ExecuteNonQuery();
                            da = new MySqlDataAdapter("SELECT * FROM " + dataChoose + " WHERE ISBN='" + textBox3.Text + "'" + "AND 库位='" + textBox6.Text + "'", conn);
                            da.Fill(data);
                            dataGridView2.DataSource = data;
                            conn.Close();
                            label6.Text = "出库成功，出库信息为：";
                            LogAdd();
                        }
                    }
                }
                catch (MySqlException ex)
                {

                    MessageBox.Show("Error: " + ex);
                }
            }
            
        }
        private void LogAdd()//出库记录
        {
            if (conn.State == ConnectionState.Open)
                conn.Close();
            conn.Open();
            if (comboBox1.SelectedItem != null)
                conn.ChangeDatabase(comboBox1.SelectedItem.ToString());//连接部分
            try
            {
                string query = "INSERT INTO 出库表 (书名, ISBN, 数量, 库位, 备注, 出库时间) VALUES('" + textBox4.Text + "', '" + textBox3.Text + "', " + textBox5.Text + ", '" + textBox6.Text + "', '" + textBox7.Text + "', '" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "')";
                MySqlCommand cmd = new MySqlCommand(query, conn);
                cmd.ExecuteNonQuery();
                conn.Close();
                conn.Dispose();
            }
            catch (Exception)
            {

                throw;
            }
        }
        private bool IfInfoFilled()//填写信息检查
        {
            if (textBox3.Text == "" || textBox5.Text == "" || textBox6.Text == "")
            {
                MessageBox.Show("ISBN号、数量和库位为必填项");
                return false;
            }
            else return true;
            
        }
 
        private bool IfDataChoose()//数据库选择检查
        {
            if (comboBox1.Enabled || comboBox2.Enabled)
            {
                MessageBox.Show("请先选择数据并确定");
                return false;
            }
            else return true;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            textBox8.Text = openFileDialog1.FileName;
            if (textBox8.Text!="")
            {
                button6.Enabled = true;
                button7.Enabled = true;
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (dataGridView3.DataSource != null)
                dataGridView3.DataSource = null;
            dataGridView3.DataSource = ToDataTable(textBox8.Text, textBox8.Text).Tables[0];
            int n = dataGridView3.RowCount-1;
            label9.Text = "共读取"+n+"记录";
        }

        private void button7_Click(object sender, EventArgs e)//批量导入
        {
            int n = dataGridView3.RowCount - 1;
            int f = 0;
            if (IfDataChoose())
                for (int i = 0; i < n; i++)
            {
                    try
                    {
                        if (conn.State == ConnectionState.Open)
                            conn.Close();
                        conn.Open();
                        if (comboBox1.SelectedItem != null)
                            conn.ChangeDatabase(comboBox1.SelectedItem.ToString());
                        data = new DataTable();
                        da = new MySqlDataAdapter("SELECT * FROM " + dataChoose + " WHERE ISBN='" + dataGridView3.Rows[i].Cells[1].Value + "'" + "AND 库位='" + dataGridView3.Rows[i].Cells[3].Value + "'", conn);
                        da.Fill(data);    //查询有无记录
                        if (data.Rows.Count == 0)//无记录
                        {
                            string query = "INSERT INTO " + dataChoose + " (书名, ISBN, 数量, 库位, 备注) VALUES('" + dataGridView3.Rows[i].Cells[0].Value + "', '" + dataGridView3.Rows[i].Cells[1].Value + "', " + dataGridView3.Rows[i].Cells[2].Value + ", '" + dataGridView3.Rows[i].Cells[3].Value + "', '" + dataGridView3.Rows[i].Cells[4].Value + "')";
                            MySqlCommand cmd = new MySqlCommand(query, conn);
                            cmd.ExecuteNonQuery();
                            conn.Close();
                            conn.Dispose();
                            f++;
                            label9.Text = "共导入"+f+"/"+n+"条数据";
                        }
                        else
                        {
                            data.PrimaryKey = new DataColumn[] { data.Columns["ISBN"], data.Columns["库位"] };
                            object[] key = { dataGridView3.Rows[i].Cells[1].Value, dataGridView3.Rows[i].Cells[3].Value };
                            DataRow dr = data.Rows.Find(key);
                            int nu = (int)dr["数量"];
                            nu = nu +Convert.ToInt32(dataGridView3.Rows[i].Cells[2].Value);
                            string query = "UPDATE " + dataChoose + " SET 数量='" + nu + "' WHERE ISBN='" + dataGridView3.Rows[i].Cells[1].Value + "'" + "AND 库位='" + dataGridView3.Rows[i].Cells[3].Value + "'";
                            MySqlCommand cmd = new MySqlCommand(query, conn);
                            cmd.ExecuteNonQuery();
                            conn.Close();
                            conn.Dispose();
                            f++;
                            label9.Text = "共导入" + f + "/" + n + "条数据";
                        }
                        
                    }
                    catch (MySqlException ex)
                    {

                        MessageBox.Show("Error: " + ex);
                    }
            }
            MessageBox.Show("导入完成" + label9.Text);
            GC.Collect();
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            textBox3.Text = dataGridView2.CurrentRow.Cells[1].Value.ToString();
            textBox4.Text = dataGridView2.CurrentRow.Cells[0].Value.ToString();
            //textBox5.Text = dataGridView2.CurrentRow.Cells[2].Value.ToString();
            textBox6.Text = dataGridView2.CurrentRow.Cells[3].Value.ToString();
            textBox7.Text = dataGridView2.CurrentRow.Cells[4].Value.ToString();
        }

        private void button9_Click(object sender, EventArgs e)//出库查询
        {
            string dateStart = dateTimePicker1.Value.ToString("yyyy-MM-dd")+" 00:00:00";
            string dateEnd = dateTimePicker2.Value.ToString("yyyy-MM-dd")+" 23:59:59";
            if (IfDataChoose())
            {
                try
                {
                    if (conn.State == ConnectionState.Open)
                        conn.Close();
                    conn.Open();
                    if (comboBox1.SelectedItem != null)
                        conn.ChangeDatabase(comboBox1.SelectedItem.ToString());
                    data = new DataTable();
                    da = new MySqlDataAdapter("SELECT * FROM 出库表 WHERE 出库时间 between'" + dateStart + "'" + "AND '" + dateEnd + "'", conn);
                    da.Fill(data);
                    dataGridView4.DataSource = data;
                    conn.Close();
                    conn.Dispose();
                }
                catch (Exception)
                {

                    throw;
                }
            }   
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            if (IfDataChoose())
            {
                string sqlselect1 = "SELECT * FROM " + dataChoose + " WHERE ISBN='" + textBox3.Text + "'";
                Select(sqlselect1, dataGridView2);
                textBox4.Text = "";
                textBox5.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)//仅显示销量
        {
            string dateStart = dateTimePicker1.Value.ToString("yyyy-MM-dd") + " 00:00:00";
            string dateEnd = dateTimePicker2.Value.ToString("yyyy-MM-dd") + " 23:59:59";
            if (IfDataChoose())
            {
                if (checkBox1.Checked)
                {
                    try
                    {
                        if (conn.State == ConnectionState.Open)
                            conn.Close();
                        conn.Open();
                        if (comboBox1.SelectedItem != null)
                            conn.ChangeDatabase(comboBox1.SelectedItem.ToString());
                        data = new DataTable();
                        da = new MySqlDataAdapter("SELECT 书名, ISBN, COUNT(数量) AS 数量 FROM 出库表 WHERE 出库时间 between'" + dateStart + "'" + "AND '" + dateEnd + "' GROUP BY ISBN", conn);
                        da.Fill(data);
                        dataGridView4.DataSource = data;
                        conn.Close();
                        conn.Dispose();
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                }
                else
                {
                    try
                    {
                        if (conn.State == ConnectionState.Open)
                            conn.Close();
                        conn.Open();
                        if (comboBox1.SelectedItem != null)
                            conn.ChangeDatabase(comboBox1.SelectedItem.ToString());
                        data = new DataTable();
                        da = new MySqlDataAdapter("SELECT * FROM 出库表 WHERE 出库时间 between'" + dateStart + "'" + "AND '" + dateEnd + "'", conn);
                        da.Fill(data);
                        dataGridView4.DataSource = data;
                        conn.Close();
                        conn.Dispose();
                    }
                    catch (Exception)
                    {

                        throw;
                    }
                }
            }
               
        }

        public static DataSet ToDataTable(string filePath,string fileName)//读取excel
        {
            string connStr = "";
            string fileType = System.IO.Path.GetExtension(fileName);
            if (string.IsNullOrEmpty(fileType)) return null;

            if (fileType == ".xls")
                connStr = "Provider=Microsoft.Jet.OLEDB.4.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 8.0;HDR=YES;IMEX=1\"";
            else
                connStr = "Provider=Microsoft.ACE.OLEDB.12.0;" + "Data Source=" + filePath + ";" + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\"";
            string sql_F = "Select * FROM [{0}]";

            OleDbConnection conn = null;
            OleDbDataAdapter da = null;
            DataTable dtSheetName = null;

            DataSet ds = new DataSet();
            try
            {
                // 初始化连接，并打开
                conn = new OleDbConnection(connStr);
                conn.Open();

                // 获取数据源的表定义元数据                        
                string SheetName = "";
                dtSheetName = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                // 初始化适配器
                da = new OleDbDataAdapter();
                for (int i = 0; i < dtSheetName.Rows.Count; i++)
                {
                    SheetName = (string)dtSheetName.Rows[i]["TABLE_NAME"];

                    if (SheetName.Contains("$") && !SheetName.Replace("'", "").EndsWith("$"))
                    {
                        continue;
                    }

                    da.SelectCommand = new OleDbCommand(String.Format(sql_F, SheetName), conn);
                    DataSet dsItem = new DataSet();
                    da.Fill(dsItem, SheetName);

                    ds.Tables.Add(dsItem.Tables[0].Copy());
                }
            }
            catch (OleDbException ex)
            {
                MessageBox.Show("Error  "+ex);
            }
            finally
            {
                // 关闭连接
                if (conn.State == ConnectionState.Open)
                {
                    conn.Close();
                    da.Dispose();
                    conn.Dispose();
                }
            }
            return ds;
        }


    }
}
