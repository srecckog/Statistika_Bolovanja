using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data;
using System.Data.SqlClient;

namespace Statistika_Bolovanja
{

    public partial class LoginForm : Form
    {
        public string connectionString = @"Data Source=192.168.0.5;Initial Catalog=RFIND;User ID=sa;Password=AdminFX9.";
        public static string idusera1,idadmin,korisnik,idprijave;
        public LoginForm()
        {
            InitializeComponent();
        }
                

        private void btn_login_Click_1(object sender, EventArgs e)
        {
            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                SqlCommand sqlCommand = new SqlCommand("SELECT * FROM korisnici where username='" + textBox2.Text.Trim() + "' and password='" + textBox1.Text.Trim() + "'", cn);
                SqlDataReader reader = sqlCommand.ExecuteReader();
                idusera1 = "";
                idadmin = "0";
                while (reader.Read())
                {
                    idusera1 = reader["grupa"].ToString();
                    idadmin = reader["ID"].ToString();
                    DialogResult = DialogResult.OK;
                }

                if (idusera1 == "")
                    MessageBox.Show("Neispravno korisničko ime ili lozinka !");

                    cn.Close();
                
                    cn.Open();
                    korisnik   = textBox2.Text.Trim();
                    idprijave  = idadmin + "-" + DateTime.Now;
                    sqlCommand = new SqlCommand("insert into kks_log (datum,korisnik,idprijave,opis) values  ( getdate(),'" + korisnik + "','" + idprijave + "','Prijava')", cn);
                    reader = sqlCommand.ExecuteReader();
                    cn.Close();

                

            }
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {

        }
    }
}