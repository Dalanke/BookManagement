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

namespace BookManagement
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private  MySqlConnection conn;
        private void button1_Click(object sender, EventArgs e)
        {
            if (conn != null)
                conn.Close();
            string user = textBox1.Text;
            string password = textBox2.Text;
            string host = textBox3.Text;
            string connStr = String.Format("server={0};user id={1}; password={2}; database=mysql; pooling=false",
               host, user, password);                    
            try
            {
                conn = new MySqlConnection(connStr);
                conn.Open();
                label4.Text = "成功建立连接！";
                if (CloseCon())
                {
                    Form2 f2 = new Form2(connStr);
                    f2.Show();
                }    
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error connecting to the server: " + ex.Message);

            }
        }
        public bool CloseCon()
        {
            try
            {
                conn.Close();
                return true;
            }
            catch (MySqlException ex)
            {
                MessageBox.Show("Error :" + ex);
                return false;
            }
        }

    }
}
