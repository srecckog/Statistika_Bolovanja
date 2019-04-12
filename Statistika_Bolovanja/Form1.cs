using System;
using System.Collections.Generic;
using System.Data;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.IO;
using DocumentFormat.OpenXml.Spreadsheet;
using ClosedXML.Excel;
using System.Drawing;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Net;
using System.Net.Sockets;
//using Microsoft.Office.Interop.Word;


namespace Statistika_Bolovanja
{
    public partial class Form1 : Form
    {
        public string connectionString = @"Data Source=192.168.0.3;Initial Catalog = RFIND ; User ID=sa ; Password=AdminFX9.";
        public string connectionStringFeroapp = @"Data Source=192.168.0.3;Initial Catalog = Feroapp ; User ID=sa ; Password=AdminFX9.";
        public string connectionStringpa = @"Data Source=192.168.0.6;Initial Catalog =PantheonFxAt; User ID=sa ; Password=AdminFX9.";
        public string connectionStringpb = @"Data Source=192.168.0.6;Initial Catalog =PantheonTKB; User ID=sa ; Password=AdminFX9.";

        public int prikazinormu = 0, prikazibolovanje = 0, prikaziizostanak = 0, prikazigodisnji = 0;
        public int idradnika1;
        public int g1, m1, prikaz1;
        public string danx;

        public class radnici
        {
            public string id { get; set; }
            public string ime { get; set; }
            public string prezime { get; set; }
            public string datumrodjenja { get; set; }
            public string rfid { get; set; }
            public string rfid2 { get; set; }
            public string rfidhex { get; set; }
            public string lokacija { get; set; }
            public string poduzece { get; set; }
            public string mt { get; set; }
        }
        public class mjesto_troska
        {
            public string id { get; set; }
            public string naziv { get; set; }
        }
        public class project
        {
            public string id { get; set; }
            public string naziv { get; set; }
        }

        public class funkcijaa
        {
            public string id { get; set; }
            public string funkcija { get; set; }
        }
        public class skolovanje1
        {
            public string idskolovanje   { get; set; }
            public string opis  { get; set; }
        }


        public class linija
        {
            public string id { get; set; }
            public string naziv { get; set; }
        }

        public class CheckboxItem
        {
            public string Text { get; set; }
            public object Value { get; set; }
            public bool checked1 { get; set; }

            public override string ToString()
            {
                return Text;
            }
        }


        public class praznici
        {
            public string dan { get; set; }
            public string ime { get; set; }

        }

        public List<radnici> radnicii = new List<radnici>();
        public List<mjesto_troska> lista_mt = new List<mjesto_troska>();
        public List<project> lista_projects = new List<project>();
        public List<funkcijaa> lista_funkcija = new List<funkcijaa>();
        public List<linija> lista_linija = new List<linija>();
        public List<skolovanje1> lista_skolovanja = new List<skolovanje1>();
        public List<CheckboxItem> cboxlista = new List<CheckboxItem>();
        public List<CheckboxItem> cboxlistaskol = new List<CheckboxItem>();


        public List<praznici> praznicii = new List<praznici>();
        public static string idloged, idadm,korisnik, idprijave;
        
        public int ssindex1;

        public Form1()
        {
            InitializeComponent();
            idloged = LoginForm.idusera1.Trim();    // grupa
            idadm = LoginForm.idadmin.Trim();        // user
            idprijave = LoginForm.idprijave.Trim();        // user
            bt_print1.Visible = false;
            prikaz1 = 1;


            Ocisti();
            pnl_norme.Visible = false;
            dgv_norme.Visible = false;
            pnl_vjestine.Visible = false;

            // punjenje liste radnika
            korisnik = "";
            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                string sql1 = "select username from korisnici where id= " + idadm.ToString();
                SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                SqlDataReader reader = sqlCommand.ExecuteReader();
                
                while (reader.Read())
                {
                    korisnik = reader["username"].ToString();
                }
                cn.Close();
            }
            label129.Text = "Korisnik: "+korisnik;  // ime korisnika
            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                SqlCommand sqlCommand = new SqlCommand("SELECT r.id, r.ime, prezime,rfid,rfid2,rfidhex,lokacija,mt,poduzece FROM  radnici_ r left join mjestotroska m on r.mt = m.id where r.neradi=0 order by prezime", cn);

                if (idloged == "8")   // ako je grupa 8 ili tehn. direktor, može vidjeti sve, može printati
                {
                    bt_print1.Visible = true;
                    label33.Visible = true;
                    label34.Visible = true;
                    label32.Visible = true;

                }
                else
                {
                    if (idloged== "7")  // tehnički direktor može vidjeti određene grupe  1- tokarija,steleri, 2 održavanje,,alatnica,brusenje 5 kaljenje 13 tehnologija
                    {
                        sqlCommand = new SqlCommand("SELECT r.id, r.ime, prezime,rfid,rfid2,rfidhex,lokacija,mt,poduzece FROM  radnici_ r left join mjestotroska m on r.mt = m.id where r.neradi=0 and m.grupa1 in ( 1,2,3,5,13) order by prezime", cn);
                    }
                    else
                    {
                        sqlCommand = new SqlCommand("SELECT r.id, r.ime, prezime,rfid,rfid2,rfidhex,lokacija,mt,poduzece FROM  radnici_ r left join mjestotroska m on r.mt = m.id where r.neradi=0 and m.grupa1='" + idloged + "' order by prezime", cn);
                        label33.Visible = true;
                        label34.Visible = true;
                        label32.Visible = true;
                    }

                    // nemoj pokazati satnice
                    //label33.Visible = false;
                    //label34.Visible = false;
                    //label32.Visible = false;

                }


                SqlDataReader reader = sqlCommand.ExecuteReader();

                while (reader.Read())
                {
                    // reader["Datum"].ToString();

                    radnici radnik = new radnici();
                    radnik.id = ((int.Parse)(reader["ID"].ToString())).ToString();
                    radnik.ime = reader["Ime"].ToString().TrimEnd();
                    radnik.prezime = reader["Prezime"].ToString().TrimEnd();
                    radnik.rfidhex = reader["RFIDHex"].ToString();
                    radnik.rfid2 = reader["RFID2"].ToString();
                    radnik.rfid = reader["RFID"].ToString();
                    radnik.lokacija= reader["lokacija"].ToString();
                    radnik.mt = reader["mt"].ToString();
                    radnik.poduzece = "3";
                    if (reader["poduzece"].ToString().Contains("Fero"))
                    {
                        radnik.poduzece = "1";
                    }
                    //  radnik.datumrodjenja = reader["DatumRodjenja"].ToString();
                    radnicii.Add(radnik);

                }
                cn.Close();

                cn.Open();
                sqlCommand = new SqlCommand("SELECT * from praznici", cn);
                reader = sqlCommand.ExecuteReader();

                while (reader.Read())
                {
                    praznici praznik = new praznici();
                    praznik.dan = reader["datum"].ToString();
                    praznik.ime = reader["opis"].ToString();
                    praznicii.Add(praznik);
                }
                cn.Close();

                cn.Open();
                if (idloged == "8")
                {
                    sqlCommand = new SqlCommand("SELECT * from mjestotroska order by id", cn);
                }
                else
                {
                    if (idloged == "7")
                    {
                        sqlCommand = new SqlCommand("SELECT * from mjestotroska where grupa1 in (1,2,3,5,13) order by id", cn);
                    }
                    else
                    {
                        sqlCommand = new SqlCommand("SELECT * from mjestotroska where grupa1 = " + idloged.ToString()+" order by id", cn);
                    }
                }
                reader = sqlCommand.ExecuteReader();

                while (reader.Read())
                {
                    mjesto_troska mt1 = new mjesto_troska();
                    mt1.id = reader["id"].ToString();
                    mt1.naziv = reader["naziv"].ToString();
                    lista_mt.Add(mt1);
                }
                cn.Close();

                cn.Open();
                sqlCommand = new SqlCommand("SELECT * from projekti  order by grupa,naziv", cn);
                
                reader = sqlCommand.ExecuteReader();

                while (reader.Read())
                {
                    project project1  = new project();
                    project1.id = reader["id"].ToString();
                    project1.naziv = reader["naziv"].ToString();
                    lista_projects.Add(project1);
                }
                cn.Close();

                cn.Open();
                string gf = "";

                if (idloged == "1")   // Tokarenje
                    gf = "A";

                if (idloged == "5")   // Kaliona
                    gf = "T";

                if (idloged == "2")   // Kaliona
                    gf = "O";

                if (idloged == "3")   // Kaliona
                    gf = "A";


                if (gf!="")
                   sqlCommand = new SqlCommand("SELECT * from funkcije  where grupa1='"+gf+"' order by grupa1,funkcija", cn);
                else
                    sqlCommand = new SqlCommand("SELECT * from funkcije  order by grupa1,funkcija", cn);

                reader = sqlCommand.ExecuteReader();

                while (reader.Read())
                {
                    funkcijaa funkcija1 = new funkcijaa();
                    funkcija1.id = reader["id"].ToString();
                    funkcija1.funkcija = reader["funkcija"].ToString();
                    lista_funkcija.Add(funkcija1);
                }
                cn.Close();

                cn.Open();
                sqlCommand = new SqlCommand("SELECT * from linije  order by id", cn);

                reader = sqlCommand.ExecuteReader();

                while (reader.Read())
                {
                    linija linija1 = new linija();
                    linija1.id = reader["id"].ToString();
                    linija1.naziv = reader["naziv"].ToString();
                    lista_linija.Add(linija1);
                }
                cn.Close();

                label85.Text = "";
                label84.Text = "";
                label83.Text = "";
                label58.Text = "";
                label46.Text = "";
                label45.Text = "";
                label73.Text = "";
                label22.Text = "";

            }


        }
        #region checkdate
        protected bool CheckDate(String date)
        {
            try
            {
                DateTime dt = DateTime.Parse(date);
                return true;
            }
            catch
            {
                return false;
            }
        }
        #endregion

        #region calculate statbo2
        private void button1_Click(object sender, EventArgs e)
        {

            using (SqlConnection cnn1 = new SqlConnection(connectionString))
            {
                cnn1.Open();
                SqlCommand cmd = new SqlCommand("dbo.bolovanja2017", cnn1);
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.Parameters.Add("@id1", SqlDbType.Int, 5).Value = id1;  // dodati parametar godinu i mjesec
                //cmd.Parameters.Add("@datum1", SqlDbType.DateTime, 5).Value = datum1;  // dodati parametar godinu i mjesec
                //cmd.Parameters.Add("@datum2", SqlDbType.DateTime, 5).Value = datum2;  // dodati parametar godinu i mjesec
                SqlDataReader rdr = null;
                SqlConnection cnn = null;
                rdr = cmd.ExecuteReader();
                string string1 = "", string2 = "", string3 = "";
                string radnikid, mjesec1, ime1 = "", godina1 = "", firma1 = "";
                int jos = 1;
                bool joss = true;
                rdr.Read();

                while (joss)
                {
                    // reader["Datum"].ToString();
                    if (rdr.HasRows)
                    {

                        if (rdr["radnikid"] == DBNull.Value)
                        {
                            joss = false;
                            continue;
                        }


                        radnikid = rdr["radnikid"].ToString();
                        if (radnikid == "541")
                            radnikid = radnikid;


                        mjesec1 = rdr["Mjesec"].ToString();
                        godina1 = rdr["Godina"].ToString();
                        if (radnikid == "541" && mjesec1 == "5")
                            radnikid = radnikid;

                        ime1 = rdr["Ime"].ToString().TrimEnd();

                        //while (rdr["radnikid"].ToString() == radnikid)
                        // {

                        string1 = rdr["DAN01"].ToString() + "," + rdr["DAN02"].ToString() + "," + rdr["DAN03"].ToString() + "," + rdr["DAN04"].ToString() + "," + rdr["DAN05"].ToString() + "," + rdr["DAN06"].ToString() + "," + rdr["DAN07"].ToString() + "," + rdr["DAN08"].ToString() + "," + rdr["DAN09"].ToString() + "," + rdr["DAN10"].ToString() + "," + rdr["DAN11"].ToString() + "," + rdr["DAN12"].ToString() + "," + rdr["DAN13"].ToString() + "," + rdr["DAN14"].ToString() + "," + rdr["DAN15"].ToString() + "," + rdr["DAN16"].ToString() + "," + rdr["DAN17"].ToString() + "," + rdr["DAN18"].ToString() + "," + rdr["DAN19"].ToString() + "," + rdr["DAN20"].ToString() + "," + rdr["DAN21"].ToString() + "," + rdr["DAN22"].ToString() + "," + rdr["DAN23"].ToString() + "," + rdr["DAN24"].ToString() + "," + rdr["DAN25"].ToString() + "," + rdr["DAN26"].ToString() + "," + rdr["DAN27"].ToString() + "," + rdr["DAN28"].ToString() + "," + rdr["DAN29"].ToString() + "," + rdr["DAN30"].ToString() + "," + rdr["DAN31"].ToString();
                        //rdr.Read();

                        // 1., mjesec bolovanja

                        int bdana = 0, bcb = 0, bdana1 = 0;
                        int pn = 1;
                        string sql1 = "";
                        string[] s1 = string1.Split(',');
                        mjesec1 = rdr["Mjesec"].ToString();
                        godina1 = rdr["Godina"].ToString();
                        firma1 = rdr["Firma"].ToString();

                        if (radnikid == "218")
                            radnikid = radnikid;

                        for (int i = 0; i < s1.Length; i++)
                        {
                            string smjesec1;
                            if (mjesec1.TrimEnd().Length == 2)
                            {
                                smjesec1 = mjesec1;
                            }
                            else
                            {
                                smjesec1 = "0" + mjesec1;
                            }

                            if ((s1[i] == "8b") || (s1[i] == "7b") || ((s1[i] == "") && (pn == 0)))
                            {
                                if (pn == 1)
                                {
                                    bcb++;
                                }
                                bdana = bdana + 1;

                                if (s1[i] == "8b" || s1[i] == "7b")
                                    bdana1 = bdana1 + 1;

                                pn = 0;
                                if (i == 30)
                                {

                                    sql1 = "insert into stat_bo2 (id,ime,mjesec,godina,brojdana,razlog,firma) values( '" + radnikid.ToString() + "','" + ime1 + " - " + radnikid.ToString() + "'," + mjesec1 + "," + godina1 + "," + bdana1.ToString() + ",'B'," + firma1 + ")";
                                    using (SqlConnection cn = new SqlConnection(connectionString))
                                    {
                                        cn.Open();
                                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                                        SqlDataReader reader = sqlCommand.ExecuteReader();
                                        cn.Close();
                                    }
                                    sql1 = "update bolovanja_m set ["+smjesec1+ "]=ISNULL([" + smjesec1 + "],0) +" + bdana1.ToString()+" where id="+radnikid.ToString() +" and godina="+godina1+ "; IF @@ROWCOUNT=0 insert into bolovanja_m(id,poduzece,godina,ime,[" + smjesec1 + "]) values( '" + radnikid.ToString() + "'," + firma1 + "," + godina1 + ",'" + ime1 + " - " + radnikid.ToString() + "'," + bdana1.ToString() + ")";
                                    if (ime1.Contains("HUSELJI"))
                                        {
                                        int z = 0;
                                    }
                                    using (SqlConnection cn = new SqlConnection(connectionString))
                                    {
                                        cn.Open();
                                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                                        SqlDataReader reader = sqlCommand.ExecuteReader();
                                        
                                        cn.Close();
                                    }

                                    //sql1 = "insert into bolovanja_m(id,poduzece,godina,ime,["+smjesec1+"]) values( '" + radnikid.ToString() + "',"+firma1+","+godina1+",'" + ime1 + " - " + radnikid.ToString() + "'," + bdana1.ToString() +")";
                                    //using (SqlConnection cn = new SqlConnection(connectionString))
                                    //{
                                    //    cn.Open();
                                    //    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                                    //    SqlDataReader reader = sqlCommand.ExecuteReader();
                                    //    cn.Close();
                                    //}


                                }


                            }
                            else
                            {
                                if (pn == 0)
                                {

                                    sql1 = "insert into stat_bo2 (id,ime,mjesec,godina,brojdana,razlog,firma) values( '" + radnikid.ToString() + "','" + ime1 + " - " + radnikid.ToString() + "'," + mjesec1 + "," + godina1 + "," + bdana1.ToString() + ",'B'," + firma1 + ")";
                                    using (SqlConnection cn = new SqlConnection(connectionString))
                                    {
                                        cn.Open();
                                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                                        SqlDataReader reader = sqlCommand.ExecuteReader();
                                        cn.Close();
                                    }
                                    sql1 = "update bolovanja_m set [" + smjesec1 + "]=ISNULL([" + smjesec1 + "],0) +" + bdana1.ToString() + " where id=" + radnikid.ToString() + " and godina=" + godina1 + "; IF @@ROWCOUNT=0 insert into bolovanja_m(id,poduzece,godina,ime,[" + smjesec1 + "]) values( '" + radnikid.ToString() + "'," + firma1 + "," + godina1 + ",'" + ime1 + " - " + radnikid.ToString() + "'," + bdana1.ToString() + ")";
                                    using (SqlConnection cn = new SqlConnection(connectionString))
                                    {
                                        cn.Open();
                                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                                        SqlDataReader reader = sqlCommand.ExecuteReader();
                                        cn.Close();
                                    }


                                    bdana1 = 0;
                                }
                                pn = 1;
                            }

                        }

                        bdana = 0; bdana1 = 0; bcb = 0;
                        pn = 1;

                        // 1.mjesec 0e    Izostanci sa posla

                        string datum1 = "", dan1 = "";
                        int bs = 0, bn = 0, bg = 0;
                        for (int i = 0; i <= s1.Length; i++)
                        {
                            if (i < 9)
                            {
                                dan1 = "0" + (i + 1).ToString();
                            }
                            else
                            {
                                dan1 = (i + 1).ToString();
                            }

                            datum1 = godina1 + "-" + mjesec1 + "-" + dan1;
                            if (radnikid == "562" && mjesec1 == "5" && godina1 == "2017")
                                radnikid = radnikid;


                            DateTime dt = new DateTime(int.Parse(godina1), int.Parse(mjesec1), 1);

                            if (DateTime.DaysInMonth(int.Parse(godina1), int.Parse(mjesec1)) == i)
                            {

                                if (bg > 0)
                                {
                                    sql1 = "insert into stat_bo2 (id,ime,mjesec,godina,brojdana,sn,nn,razlog,firma) values( '" + radnikid.ToString() + "','" + ime1 + " - " + radnikid.ToString() + "'," + mjesec1 + "," + godina1 + "," + bg.ToString() + ",0,0,'G'," + firma1 + ")";
                                    using (SqlConnection cn = new SqlConnection(connectionString))
                                    {
                                        cn.Open();
                                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                                        SqlDataReader reader = sqlCommand.ExecuteReader();
                                        cn.Close();
                                    }

                                }


                                break;
                            }
                            else
                            {
                                dt = new DateTime(int.Parse(godina1), int.Parse(mjesec1), i + 1);

                            }

                            if (i == 29 && radnikid == "218" && mjesec1 == "4")
                            {
                                int t = DateTime.DaysInMonth(int.Parse(godina1), int.Parse(mjesec1));
                                bg = bg;

                            }

                            if ((dt.DayOfWeek == DayOfWeek.Saturday) && (s1[i] == "0e"))      // izostanci nedjeljom
                            {
                                bs = bs + 1;
                            }

                            if (s1[i] == "7g" || s1[i] == "5g")      // izostanci godišnji
                            {
                                bg = bg + 1;
                            }

                            if (i == 30)
                            {

                            }

                            if ((dt.DayOfWeek == DayOfWeek.Sunday) && (s1[i] == "0e"))         // izpstanci subotom
                            {
                                bn = bn + 1;
                                if (radnikid == "1278")
                                    radnikid = radnikid;
                            }


                            if ((s1[i] == "0e"))
                            {
                                if (pn == 1)
                                {
                                    bcb++;
                                }
                                bdana = bdana + 1;

                                if (s1[i] == "0e")
                                    bdana1 = bdana1 + 1;


                                pn = 0;
                                if (i == 30)
                                {

                                    sql1 = "insert into stat_bo2 (id,ime,mjesec,godina,brojdana,sn,nn,razlog,firma) values( '" + radnikid.ToString() + "','" + ime1 + " - " + radnikid.ToString() + "'," + mjesec1 + "," + godina1 + "," + bdana1.ToString() + "," + bs.ToString() + "," + bn.ToString() + ",'N'," + firma1 + ")";
                                    using (SqlConnection cn = new SqlConnection(connectionString))
                                    {
                                        cn.Open();
                                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                                        SqlDataReader reader = sqlCommand.ExecuteReader();
                                        cn.Close();
                                    }

                                }


                            }
                            else
                            {
                                if (pn == 0)
                                {

                                    sql1 = "insert into stat_bo2 (id,ime,mjesec,godina,brojdana,sn,nn,razlog,firma) values( '" + radnikid.ToString() + "','" + ime1 + " - " + radnikid.ToString() + "'," + mjesec1 + "," + godina1 + "," + bdana1.ToString() + "," + bs.ToString() + "," + bn.ToString() + ",'N'," + firma1 + ")";

                                    using (SqlConnection cn = new SqlConnection(connectionString))
                                    {
                                        cn.Open();
                                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                                        SqlDataReader reader = sqlCommand.ExecuteReader();
                                        cn.Close();
                                    }

                                    bdana1 = 0;
                                }

                                pn = 1;

                            }
                        }


                        // zavrsetak .mjesec 0e

                        string1 = "";

                    }


                    joss = rdr.Read();
                }
                cnn1.Close();
            }
            using (SqlConnection cnn2 = new SqlConnection(connectionString))
            {
                cnn2.Open();
                string sql1 = "SELECT [Idradnika],[idvjestine],v.naziv naziv,[Firma] firma1 FROM[RFIND].[dbo].[RadniciVjestine] rv left join vjestine v on v.id=rv.idvjestine order by idradnika";
                SqlCommand cmd2 = new SqlCommand(sql1, cnn2);
                //cmd.Parameters.Add("@id1", SqlDbType.Int, 5).Value = id1;  // dodati parametar godinu i mjesec
                //cmd.Parameters.Add("@datum1", SqlDbType.DateTime, 5).Value = datum1;  // dodati parametar godinu i mjesec
                //cmd.Parameters.Add("@datum2", SqlDbType.DateTime, 5).Value = datum2;  // dodati parametar godinu i mjesec
                SqlDataReader rdr2 = null;
                rdr2 = cmd2.ExecuteReader();
                string idradnik1 = "", vjes1 = "", idfirma1 = "", firma1 = "";
                int jos = 1;
                rdr2.Read();

                while (jos == 1)
                {

                    idradnik1 = rdr2["Idradnika"].ToString();
                    idfirma1 = rdr2["firma1"].ToString();
                    if (idradnik1=="688")
                    {
                        int z = 0;
                    }
                            
                    vjes1 = "";
                    while (idradnik1 == rdr2["Idradnika"].ToString() && (jos == 1) && (idfirma1 == rdr2["firma1"].ToString()))
                    {
                        vjes1 = vjes1 + "," + rdr2["naziv"].ToString();                        
                        if (rdr2.Read())
                        {
                            jos = 1;
                        }
                        else
                        {
                            jos = 0;
                            break;
                        }


                    }

                    firma1 = "";
                    if (idfirma1 == "1")
                    {
                        firma1 = "FX";
                    }
                    else
                    {
                        firma1 = "Tokabu";
                    }

                    if (idradnik1=="75")
                        {
                        int z = 0;
                    }


                    using (SqlConnection cnn3 = new SqlConnection(connectionString))
                    {
                        cnn3.Open();
                        vjes1 = vjes1.Substring(1, vjes1.Length-1).Trim();
                        int l1 = vjes1.Length;
                        string sql3 = "update kompetencije set vještine ='" + vjes1 + "' where id=" + idradnik1 + " and rtrim(ltrim(poduzece))='" + firma1+"'";
                        SqlCommand cmd3 = new SqlCommand(sql3, cnn3);
                        SqlDataReader rdr3 = null;
                        rdr3 = cmd3.ExecuteReader();
                        cnn3.Close();
                    }

                    




                }

            }
                        // satnice

                        using (SqlConnection cnn1 = new SqlConnection(connectionString))
            {
                cnn1.Open();
                SqlCommand cmd = new SqlCommand("dbo.satnice", cnn1);
                cmd.CommandType = CommandType.StoredProcedure;
                //cmd.Parameters.Add("@id1", SqlDbType.Int, 5).Value = id1;  // dodati parametar godinu i mjesec
                //cmd.Parameters.Add("@datum1", SqlDbType.DateTime, 5).Value = datum1;  // dodati parametar godinu i mjesec
                //cmd.Parameters.Add("@datum2", SqlDbType.DateTime, 5).Value = datum2;  // dodati parametar godinu i mjesec
                SqlDataReader rdr = null;
                SqlConnection cnn = null;
                rdr = cmd.ExecuteReader();
                string string1 = "", string2 = "", string3 = "";
                string radnikid, mjesec1, ime1 = "", godina1 = "", mjesec2 = "", godina2 = "", sql1 = "";
                int jos = 1, bio = 0;
                string stimulacija1 = "0";
                string stimulacija2 = "0";
                bool joss = true;

                rdr.Read();

                while (joss)
                {
                    // reader["Datum"].ToString();
                    if (rdr.HasRows)
                    {

                        try
                        {
                            if (rdr["radnikid"] == DBNull.Value)
                            {
                                joss = false;
                                continue;
                            }

                        }
                        catch (Exception ex)
                        {
                            // nakon što napravi stat_bo2, updateiraj kompetencije_  iz stat_bo2.....
                            cnn1.Close();
                            cnn1.Open();
                            cmd = new SqlCommand("dbo.sp_kompetencije_", cnn1);
                            cmd.CommandType = CommandType.StoredProcedure;
                            //cmd.Parameters.Add("@id1", SqlDbType.Int, 5).Value = id1;  // dodati parametar godinu i mjesec
                            //cmd.Parameters.Add("@datum1", SqlDbType.DateTime, 5).Value = datum1;  // dodati parametar godinu i mjesec
                            //cmd.Parameters.Add("@datum2", SqlDbType.DateTime, 5).Value = datum2;  // dodati parametar godinu i mjesec
                            rdr = null;
                            cnn = null;
                            rdr = cmd.ExecuteReader();
                            cnn1.Close();
                            joss = false;
                            continue;
                        }


                        if (joss == false)
                            continue;


                        radnikid = rdr["radnikid"].ToString();

                        if (radnikid == "5")
                            radnikid = radnikid;

                        mjesec1 = rdr["Mjesec"].ToString();
                        godina1 = rdr["Godina"].ToString();

                        ime1 = rdr["Ime"].ToString();

                        string satnica1 = rdr["SatnicaKn"].ToString();

                        stimulacija1 = "0.0";
                        bio = 0;


                        int mjesec11 = (int.Parse)(rdr["Mjesec"].ToString());
                        int godina11 = (int.Parse)(rdr["Godina"].ToString());
                        DateTime dat2 = DateTime.Now.AddMonths(-1);


                        if ((dat2.Month == mjesec11) && (dat2.Year == godina11) && (bio == 0))
                        {
                            stimulacija1 = stimulacija1.Replace(",", ".");
                            bio = 1;
                        }

                        DateTime dat3 = DateTime.Now.AddMonths(-3);
                        int dat3sada = DateTime.Now.Year * 100 + DateTime.Now.Month;
                        int dat3tab = dat3.Year * 100 + dat3.Month;
                        int raz3 = dat3sada - dat3tab;
                        int stimulacija3 = 0;

                        if ((raz3 <= 3 && raz3 >= 1) && (bio == 0))
                        {

                            //   int i17 = int.Parse(stimulacija1);

                            //stimulacija3 = stimulacija3 + (int.Parse(stimulacija1.Replace(",", "."))) ;
                            //stimulacija3 = stimulacija3 + (int.Parse(stimulacija1)) ;
                            bio = 1;
                        }


                        stimulacija2 = rdr["Stimulacija2"].ToString();
                        string satnica10 = rdr["SatnicaKn"].ToString();
                        bool joss2 = true;
                        int insert1 = 0;

                        while (joss2)
                        {

                            ime1 = rdr["Ime"].ToString();
                            string firma = rdr["Firma"].ToString();

                            mjesec11 = (int.Parse)(rdr["Mjesec"].ToString());
                            godina11 = (int.Parse)(rdr["Godina"].ToString());

                            string satnica2 = rdr["SatnicaKn"].ToString();
                            dat2 = DateTime.Now.AddMonths(-1);

                            if ((dat2.Month == mjesec11) && (dat2.Year == godina11) && (bio == 0))
                            {
                                stimulacija1 = stimulacija1.Replace(",", ".");
                                bio = 1;
                            }

                            satnica1 = satnica10.Replace(",", ".");
                            satnica2 = satnica2.Replace(",", ".");
                            dat3 = DateTime.Now.AddMonths(-3);
                            dat3sada = DateTime.Now.Year * 100 + DateTime.Now.Month;
                            dat3tab = dat3.Year * 100 + dat3.Month;
                            raz3 = dat3sada - dat3tab;


                            if ((raz3 <= 3 && raz3 >= 1) && (bio == 0))
                            {
                                stimulacija3 = stimulacija3 + (int.Parse(stimulacija1.Replace(",", ".")));
                                bio = 1;
                            }


                            sql1 = "insert into satnica (firma,radnikid,ime,godina,mjesec,satnicastara,satnica) values( " + firma + "," + radnikid.ToString() + ",'" + ime1 + "'," + godina2 + "," + mjesec2 + "," + satnica2 + "," + satnica1 + ")";


                            if (satnica10 != rdr["SatnicaKn"].ToString())
                            {

                                if (radnikid == "5")
                                {

                                    radnikid = radnikid;
                                }
                                ime1 = rdr["Ime"].ToString();
                                firma = rdr["Firma"].ToString();
                                satnica2 = rdr["SatnicaKn"].ToString();
                                //  stimulacija1 = rdr["Stimulacija1"].ToString();
                                //  stimulacija1 = stimulacija1.Replace(",", ".");
                                // godina2 = rdr["godina"].ToString();
                                // mjesec2 = rdr["mjesec"].ToString();


                                satnica1 = satnica1.Replace(",", ".");
                                satnica2 = satnica2.Replace(",", ".");

                                sql1 = "insert into satnica (firma,radnikid,ime,godina,mjesec,satnicastara,satnica) values( " + firma + "," + radnikid.ToString() + ",'" + ime1 + "'," + godina2 + "," + mjesec2 + "," + satnica2 + "," + satnica1 + ")";

                                using (SqlConnection cn = new SqlConnection(connectionString))
                                {
                                    cn.Open();
                                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                                    SqlDataReader reader = sqlCommand.ExecuteReader();
                                    cn.Close();

                                    if (radnikid == "5")
                                    {
                                        radnikid = radnikid;
                                    }
                                    if (firma == "1")
                                        sql1 = "update kompetencije  set satnicabruto= " + satnica1 + " ,satnicastara= " + satnica2 + " , satnicaNovaod =' " + godina2 + " - " + mjesec2 + "' where id=" + radnikid.ToString() + " and poduzece='FX'";
                                    else
                                        sql1 = "update kompetencije  set satnicabruto=" + satnica1 + " ,satnicastara= " + satnica2 + " , satnicaNovaod =' " + godina2 + " - " + mjesec2 + "' where id=" + radnikid.ToString() + " and poduzece='Tokabu'";


                                    cn.Open();
                                    sqlCommand = new SqlCommand(sql1, cn);
                                    reader = sqlCommand.ExecuteReader();
                                    cn.Close();

                                }
                                insert1 = 1;

                                bool joss3 = true;
                                while (joss3)
                                {
                                    joss3 = rdr.Read();
                                    if (joss3)
                                    {

                                        if (radnikid != rdr["radnikid"].ToString())
                                            joss3 = false;
                                        else
                                        {

                                        }

                                    }


                                }
                                joss2 = false;
                                continue;

                            }

                            godina2 = rdr["godina"].ToString();
                            mjesec2 = rdr["mjesec"].ToString();

                            joss2 = rdr.Read();
                            if (joss2)
                            {
                                if (radnikid != rdr["radnikid"].ToString())
                                {
                                    if (radnikid == "1037")
                                    {

                                        radnikid = radnikid;

                                    }
                                    joss2 = false;
                                    if (insert1 == 0)
                                    {

                                        using (SqlConnection cn = new SqlConnection(connectionString))
                                        {
                                            cn.Open();
                                            SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                                            SqlDataReader reader = sqlCommand.ExecuteReader();
                                            cn.Close();

                                            if (firma == "1")
                                                sql1 = "update kompetencije  set satnicabruto= " + satnica1 + " ,satnicastara= " + satnica2 + " , satnicaNovaod =' " + godina2 + " - " + mjesec2 + "' where id=" + radnikid.ToString() + " and poduzece='FX'";
                                            else
                                                sql1 = "update kompetencije  set satnicabruto=" + satnica1 + " ,satnicastara= " + satnica2 + " , satnicaNovaod =' " + godina2 + " - " + mjesec2 + "' where id=" + radnikid.ToString() + " and poduzece='TB'";


                                            cn.Open();
                                            sqlCommand = new SqlCommand(sql1, cn);
                                            reader = sqlCommand.ExecuteReader();
                                            cn.Close();

                                        }



                                    }
                                    else
                                    {

                                    }
                                }

                            }




                        }
                    }



                }  // end of satncie




            }

        }

        #endregion
        #region zbirni pregled kompetencija

        private void zbirniPregledToolStripMenuItem_Click(object sender, EventArgs e)

        {
            Ocisti();
            OcistiPanele();

            //panel1.Visible = true;            
            dgv_psr.Visible = true;
            dgv_psr.ReadOnly = true;

            pl_zbirni.Visible = true;

            pl_zbirni.Width = Form1.ActiveForm.Width - 100;
            pl_zbirni.Height = Form1.ActiveForm.Height - 100;


            dgv_zbirni.Width = pl_zbirni.Width;
            dgv_zbirni.Height = pl_zbirni.Height;
            dgv_zbirni.ReadOnly = true;

            string sql1 = "";

            if (idloged == "8")
            {
                sql1 = "select k.Id,PrezimeIme,Funkcija,Projekt,Hala,Linija,Vještine,Školovanje_Posto,k.Mjesto_troska,k.RadnoMjesto,k.DatumZaposlenja,Istek_Ugovora,Staz,SatnicaStara,SatnicaNovaOd,SatnicaBruto,SatnicaNeto,Godisnji_ostalo,Godisnji_dana1,Godisnji_dana3,Godisnji_dana6,Godisnji_dana12,DolasciNedjeljom1,DolasciNedjeljom3,DolasciNedjeljom6,DolasciNedjeljom12,NedolasciNedjeljom1,DolasciPraznikom1,DolasciPraznikom3,DolasciPraznikom6,DolasciPraznikom12,NedolasciNedjeljom3,NedolasciNedjeljom6,NedolasciNedjeljom12,NedolasciSubotom1,NedolasciSubotom3,NedolasciSubotom6,NedolasciSubotom12,Bolovanje_broj1,Bolovanje_broj3,Bolovanje_broj6,Bolovanje_broj12,Bolovanja_dana1,Bolovanja_dana3,Bolovanja_dana6,Bolovanja_dana12,Stimulacija1,Stimulacija3,Stimulacija6,Stimulacija12,NedostajeDo8sati1,NedostajeDo8sati3,NedostajeDo8sati6,NedostajeDo8sati12,Kasni1,Kasni3,Kasni6,Kasni12,PreranoOtisao1,PreranoOtisao3,PreranoOtisao6,PreranoOtisao12,Neopravdano_puta1,Neopravdano_puta3,Neopravdano_puta6,Neopravdano_puta12,NeopravdaniDani1,NeopravdaniDani3,NeopravdaniDani6,NeopravdaniDani12,NormaPosto1,NormaPosto3,NormaPosto6,NormaPosto12,Napomena,k.Poduzece" +
                    " from kompetencije k " +  "left join mjestotroska m  on k.mjesto_troska = m.naziv " + " left join radnici_ r on r.id = k.id" +  " where r.neradi=0 " +  " order by prezimeime";
            }
            else
            {
                if (idloged != "7")
                {
                  sql1 = "select K.Id,PrezimeIme,Funkcija,Projekt,Hala,Linija,Vještine,Školovanje_Posto,k.Mjesto_troska,K.RadnoMjesto,k.DatumZaposlenja,Istek_Ugovora,Staz,Godisnji_ostalo,Godisnji_dana1,Godisnji_dana3,Godisnji_dana6,Godisnji_dana12,DolasciNedjeljom1,DolasciNedjeljom3,DolasciNedjeljom6,DolasciNedjeljom12,NedolasciNedjeljom1,DolasciPraznikom1,DolasciPraznikom3,DolasciPraznikom6,DolasciPraznikom12,NedolasciNedjeljom3,NedolasciNedjeljom6,NedolasciNedjeljom12,NedolasciSubotom1,NedolasciSubotom3,NedolasciSubotom6,NedolasciSubotom12,Bolovanje_broj1,Bolovanje_broj3,Bolovanje_broj6,Bolovanje_broj12,Bolovanja_dana1,Bolovanja_dana3,Bolovanja_dana6,Bolovanja_dana12,Stimulacija1,Stimulacija3,Stimulacija6,Stimulacija12,NedostajeDo8sati1,NedostajeDo8sati3,NedostajeDo8sati6,NedostajeDo8sati12,Kasni1,Kasni3,Kasni6,Kasni12,PreranoOtisao1,PreranoOtisao3,PreranoOtisao6,PreranoOtisao12,Neopravdano_puta1,Neopravdano_puta3,Neopravdano_puta6,Neopravdano_puta12,NeopravdaniDani1,NeopravdaniDani3,NeopravdaniDani6,NeopravdaniDani12,NormaPosto1,NormaPosto3,NormaPosto6,NormaPosto12,Napomena,k.Poduzece from kompetencije k left join mjestotroska m  on k.mjesto_troska = m.naziv " + " left join radnici_ r on r.id = k.id" + " where r.neradi=0  and m.grupa1=" + idloged + " order by prezimeime";
//                    sql1 = "select Id,PrezimeIme,Funkcija,Projekt,Hala,Linija,Vještine,Školovanje_Posto,Mjesto_troska,RadnoMjesto,Vjestina,DatumZaposlenja,Istek_Ugovora,Staz,Zruto,SatnicaNeto,Godisnji_ostalo,Godisnji_dana1,Godisnji_dana3,Godisnji_dana6,Godisnji_dana12,DolasciNedjeljom1,DolasciNedjeljom3,DolasciNedjeljom6,DolasciNedjeljom12,NedolasciNedjeljom1,DolasciPraznikom1,DolasciPraznikom3,DolasciPraznikom6,DolasciPraznikom12,NedolasciNedjeljom3,NedolasciNedjeljom6,NedolasciNedjeljom12,NedolasciSubotom1,NedolasciSubotom3,NedolasciSubotom6,NedolasciSubotom12,Bolovanje_broj1,Bolovanje_broj3,Bolovanje_broj6,Bolovanje_broj12,Bolovanja_dana1,Bolovanja_dana3,Bolovanja_dana6,Bolovanja_dana12,Stimulacija1,Stimulacija3,Stimulacija6,Stimulacija12,NedostajeDo8sati1,NedostajeDo8sati3,NedostajeDo8sati6,NedostajeDo8sati12,Kasni1,Kasni3,Kasni6,Kasni12,PreranoOtisao1,PreranoOtisao3,PreranoOtisao6,PreranoOtisao12,Neopravdano_puta1,Neopravdano_puta3,Neopravdano_puta6,Neopravdano_puta12,NeopravdaniDani1,NeopravdaniDani3,NeopravdaniDani6,NeopravdaniDani12,NormaPosto1,NormaPosto3,NormaPosto6,NormaPosto12,Napomena,Poduzece" +
//                        " from kompetencije k " + "left join mjestotroska m  on k.mjesto_troska = m.naziv " + " left join radnici_ r on r.id = k.id" + " where r.neradi=0 " + " order by prezimeime";

                }
                else
                {// tehnički direktor,DARAPI
//                    sql1 = "select k.* from kompetencije k left join mjestotroska m  on k.mjesto_troska = m.naziv " + " left join radnici_ r on r.id = k.id" + " where r.neradi=0 and m.grupa1 in (1,2,3,5,13) order by prezimeime";
                    sql1 = "select K.Id,PrezimeIme,Funkcija,Projekt,Hala,Linija,Vještine,Školovanje_Posto,k.Mjesto_troska,K.RadnoMjesto,k.DatumZaposlenja,Istek_Ugovora,Staz,Godisnji_ostalo,Godisnji_dana1,Godisnji_dana3,Godisnji_dana6,Godisnji_dana12,DolasciNedjeljom1,DolasciNedjeljom3,DolasciNedjeljom6,DolasciNedjeljom12,NedolasciNedjeljom1,DolasciPraznikom1,DolasciPraznikom3,DolasciPraznikom6,DolasciPraznikom12,NedolasciNedjeljom3,NedolasciNedjeljom6,NedolasciNedjeljom12,NedolasciSubotom1,NedolasciSubotom3,NedolasciSubotom6,NedolasciSubotom12,Bolovanje_broj1,Bolovanje_broj3,Bolovanje_broj6,Bolovanje_broj12,Bolovanja_dana1,Bolovanja_dana3,Bolovanja_dana6,Bolovanja_dana12,Stimulacija1,Stimulacija3,Stimulacija6,Stimulacija12,NedostajeDo8sati1,NedostajeDo8sati3,NedostajeDo8sati6,NedostajeDo8sati12,Kasni1,Kasni3,Kasni6,Kasni12,PreranoOtisao1,PreranoOtisao3,PreranoOtisao6,PreranoOtisao12,Neopravdano_puta1,Neopravdano_puta3,Neopravdano_puta6,Neopravdano_puta12,NeopravdaniDani1,NeopravdaniDani3,NeopravdaniDani6,NeopravdaniDani12,NormaPosto1,NormaPosto3,NormaPosto6,NormaPosto12,Napomena,k.Poduzece" +
                    " from kompetencije k " + "left join mjestotroska m  on k.mjesto_troska = m.naziv " + " left join radnici_ r on r.id = k.id" + " where r.neradi=0 and m.grupa1 in (1,2,3,5,13)" + " order by prezimeime";
                }
            }

            SqlConnection connection = new SqlConnection(connectionString);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            dgv_zbirni.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgv_zbirni.DataSource = ds;
            dgv_zbirni.DataMember = "event";

            if ((idloged == "8") || (idadm == "14"))
            {

            }
            else
            {
                dgv_zbirni.Columns.RemoveAt(9);
                dgv_zbirni.Columns.RemoveAt(9);
                dgv_zbirni.Columns.RemoveAt(9);
                dgv_zbirni.Columns.RemoveAt(9);
            }

            dgv_zbirni.AutoResizeColumns();
            dgv_zbirni.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgv_zbirni.DoubleBuffered(true);
            this.dgv_zbirni.Columns["Prezimeime"].Frozen = true;

        }


        #endregion

        #region osobne kompetencije
        private void pregledOsobneKompetencijeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Ocisti();
            OcistiPanele();
            pnl_norme.Visible = false;


            panel1.Visible = true;
            pl_psr.Visible = true;
            dgv_psr.Visible = true;
            
            using (SqlConnection cnn1 = new SqlConnection(connectionString))
            {

                var dataSource = new List<radnici>();

                dataSource.Add(new radnici() { prezime = " - ", id = "0" });

                foreach (var radnikk in radnicii)
                {
                    dataSource.Add(new radnici() { prezime = radnikk.prezime + " " + radnikk.ime + " - " + radnikk.id.ToString() + " ", id = radnikk.id });
                }

                combo_listadjelatnika.MaxDropDownItems = 60;
                this.combo_listadjelatnika.DataSource = dataSource;
                this.combo_listadjelatnika.DisplayMember = "prezime";
                this.combo_listadjelatnika.ValueMember = "id";

            }

            using (SqlConnection cnn1 = new SqlConnection(connectionString))
            {

                var dataSource = new List<mjesto_troska>();

                dataSource.Add(new mjesto_troska() { naziv = " - ", id = "0" });

                foreach (var mt1 in lista_mt)
                {
                    dataSource.Add(new mjesto_troska() { naziv = mt1.naziv, id = mt1.id });
                }

                cbx_mjestotroska.MaxDropDownItems = 60;
                this.cbx_mjestotroska.DataSource = dataSource;
                this.cbx_mjestotroska.DisplayMember = "Naziv";
                this.cbx_mjestotroska.ValueMember = "id";
            }

        }

        #endregion

        #region izracun staza
        // izračun staža na temelju datuma zaposlenja
        private string staz( DateTime datumz , DateTime kraj )
        {

            int y1 = datumz.Year;
            int m1 = datumz.Month;
            int d1 = datumz.Day;

            int y2 = kraj.Year;
            int m2 = kraj.Month;
            int d2 = kraj.Day;


            d1 = d2 - d1;
            if (d1 < 0)
            {
                d1 = d1 + 30 ;
                m2 = m2 - 1;
            }

            m1 = m2 - m1;
            if (m1 < 0)
            {

                m1 = m1 + 12;
                y2 = y2 - 1;
            }

            y1 = y2 - y1;


            return (y1.ToString() + " g " + m1.ToString() + " m " + d1.ToString() + "d");
        }

        #endregion

        private string NormaRadnika(int Idradnika, DateTime datum1)
        {
            string Norma1 = "";
            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                string dat1 = datum1.Year.ToString() + "-" + datum1.Month.ToString() + "-" + datum1.Day.ToString();

                //SqlCommand sqlCommand = new SqlCommand("SELECT [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt FROM radniciAT0 union all select [Radnik id] as id,[ime x] as ime,[prezime x] as prezime,rfid,rfid2,rfidhex,lokacija,mt from radniciTB0 where mt in ( 700,702,703,710,716) order by prezime", cn);
                // SqlCommand sqlCommand = new SqlCommand("SELECT id, ime, prezime,rfid,rfid2,rfidhex,lokacija,mt FROM  radnici_ where mt in ( 700,702,703,710,716)  order by prezime", cn);
                string sql1 = "select e.Datum,e.Linija,e.Hala,e.Brojrn,e.Norma,e.vrijemeod,e.vrijemedo,e.Kolicinaok , e.OtpadObrada,Napomena1 from feroapp.dbo.evidencijanormiview e left join feroapp.dbo.radnici r on r.id_radnika = e.id_radnika left join feroapp.dbo.Proizvodi p on p.id_pro = e.id_pro where r.ID_Fink= " + Idradnika.ToString() + " and datum='" + dat1 + "' order by e.datum desc";
                SqlCommand sqlCommand;

                sqlCommand = new SqlCommand(sql1, cn);
                double minuta = 0.0;

                SqlDataReader reader = sqlCommand.ExecuteReader();

                while (reader.Read())
                {
                    // reader["Datum"].ToString();
                    if (reader.HasRows)
                    {
                        int kol = (int.Parse)(reader["Kolicinaok"].ToString());
                        int norm = (int.Parse)(reader["Norma"].ToString());
                        string brisi = reader["VrijemeOd"].ToString();
                        if (brisi.Length == 0)
                        {
                            Norma1 = "---";
                            continue;
                        }

                        TimeSpan t1 = (TimeSpan.Parse(reader["VrijemeOd"].ToString()));
                        TimeSpan t2 = (TimeSpan.Parse(reader["VrijemeDo"].ToString()));
                        if (t2 < t1)
                        {
                            minuta = t2.TotalMinutes + 1440 - t1.TotalMinutes;
                        }
                        else
                        {
                            minuta = t2.TotalMinutes - t1.TotalMinutes;
                        }

                        norm = (int)(norm * minuta / 480.00);

                        if (kol > norm)
                        {
                            Norma1 = "Prebačaj";
                        }
                        if (kol < norm)
                        {
                            Norma1 = "Podbačaj";
                        }
                    }


                }
                cn.Close();

            }

            return (Norma1);
        }

        #region djelatnika, update labels,datagridview


        private void combo_listadjelatnika_SelectedIndexChanged(object sender, EventArgs e)
        {

            panel1.Visible = true;
            pl_zbirni.Visible = false;
            pnl_norme.Visible = true;
            dgv_norme.Visible = true;
            label30.Text = "";
            string satnica = "0";
            if (idloged=="8")
            {
                satnica = "1";
            }


            using (SqlConnection cnn1 = new SqlConnection(connectionString))
            {

                if (combo_listadjelatnika.SelectedIndex <= 0)
                {
                    return;
                }

                int ssindex = int.Parse(combo_listadjelatnika.SelectedValue.ToString());
                ssindex1 = ssindex;

                string poduzece1 = radnicii.Find(item => item.id == ssindex.ToString()).poduzece;
                string poduzece1s = "";
                if (poduzece1 == "1")
                {
                    lbl_poduzece.Text = "Poduzece: Feroimpex AT";
                    poduzece1s = "Feroimpex";
                }
                else
                {
                    lbl_poduzece.Text = "Poduzece: ToKaBu";
                    poduzece1s = "Tokabu";
                }

                if ((ssindex > 8000) && ( ssindex<10000))
                {
                    ssindex1 = ssindex - 8000;

                }


                string sql = "";

                if (ssindex1 == 17)
                {

                    sql = "select e.Datum,e.Linija,e.Hala,e.Brojrn,e.Norma,e.Kolicinaok , e.OtpadObrada,Napomena1,e.Id_pro,p.nazivpro,e.Vrijemeod,'' UkupnoMinuta,e.Vrijemedo " +
                            "from feroapp.dbo.evidencijanormiview e left join feroapp.dbo.radnici r on r.id_radnika = e.id_radnika left join feroapp.dbo.Proizvodi p on p.id_pro = e.id_pro where r.ID_Fink= 973  and r.id_firme=1 and DATEDIFF(month,e.datum, GETDATE()) <= 13 order by e.datum desc";

                }
                else
                {
                    //string sql = "select month(p.datum) Mjesec, p.datum,p.dosao,p.otisao,p.hala,p.smjena,p.radnomjesto,p.napomena,p.kasni,p.preranootisao,p.ukupno_minuta from radnici_ r left join pregledvremena p on r.id=p.idradnika  where idradnika= " + ssindex + " order by datum desc";
                    sql = "select e.Datum,e.Linija,e.Hala,e.Brojrn,e.Norma,0 Planirano,e.Kolicinaok , e.OtpadObrada,Napomena1,e.Id_pro,p.nazivpro,e.Vrijemeod,e.Vrijemedo,''  UkupnoMinuta " +
                                "from feroapp.dbo.evidencijanormiview e left join feroapp.dbo.radnici r on r.id_radnika = e.id_radnika left join feroapp.dbo.Proizvodi p on p.id_pro = e.id_pro where r.ID_Fink= " + ssindex1 + "  and r.id_firme=" + poduzece1 + "  and DATEDIFF(month,e.datum, GETDATE()) <= 13 order by e.datum desc";
                }


                SqlConnection connection = new SqlConnection(connectionStringFeroapp);

                SqlDataAdapter dataadapter = new SqlDataAdapter(sql, connection);
                DataSet ds = new DataSet();
                connection.Open();
                dataadapter.Fill(ds, "eventn");
                connection.Close();
                label92.BackColor = System.Drawing.Color.LawnGreen;

                dgv_norme.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dgv_norme.DataSource = ds;
                dgv_norme.DataMember = "eventn";
                double minuta = 0.0;
                int norm1 = 0, norm = 0;
                string hala1 = "", linija1 = "";

                foreach (DataGridViewRow row in dgv_norme.Rows)
                {
                    if (row.Cells[1].Value == null)
                    { }
                    else
                    {
                        string s1 = row.Cells[0].Value.ToString(); ;
                        DateTime dat0 = DateTime.Parse(s1);
                        DayOfWeek dat1 = dat0.DayOfWeek;

                        foreach (var item in praznicii)                   // norma praznici
                        {
                            if (item.dan.ToString() == s1)
                                row.DefaultCellStyle.BackColor = System.Drawing.Color.LightPink;
                        }

                        if (dat1 == DayOfWeek.Saturday)
                        {
                            row.DefaultCellStyle.BackColor = System.Drawing.Color.Yellow;
                        }
                        if (dat1 == DayOfWeek.Sunday)
                        {
                            row.DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                        }

                        //hala1 = (int.Parse)(row.Cells[2].Value.ToString());
                        //linija1 = (int.Parse)(row.Cells[1].Value.ToString());
                        norm1 = (int.Parse)(row.Cells[4].Value.ToString());
                        int kol = (int.Parse)(row.Cells[6].Value.ToString());
                        string ts1 = row.Cells[11].Value.ToString();
                        string ts2 = row.Cells[12].Value.ToString();
                        
                        minuta = 0;
                        if (ts1.Length > 0 && ts2.Length > 0)
                        {
                            TimeSpan t1 = (TimeSpan.Parse(ts1));
                            TimeSpan t2 = (TimeSpan.Parse(ts2));
                            DateTime datum1s = new DateTime(dat0.Year, dat0.Month, dat0.Day, t1.Hours, t1.Minutes, t1.Seconds);
                            DateTime datum2s = new DateTime(dat0.Year, dat0.Month, dat0.Day, t2.Hours, t2.Minutes, t2.Seconds);

                            if (t2 < t1)
                            {
                                minuta = t2.TotalMinutes + 1440 - t1.TotalMinutes;
                                datum2s = new DateTime(dat0.Year, dat0.Month, dat0.AddDays(1).Day, t2.Hours, t2.Minutes, t2.Seconds);
                            }
                            else
                            {
                                minuta = t2.TotalMinutes - t1.TotalMinutes;
                            }

                            
                        }
                        row.Cells[13].Value = minuta.ToString();
                        norm = (int)(norm1 * minuta / (480.0));
                        row.Cells[5].Value = norm.ToString();


                        if ((kol) >= (norm))
                        {
                            //                            row.Cells[6].Style.BackColor = System.Drawing.Color.LawnGreen;
                            row.Cells[5].Style.BackColor = System.Drawing.Color.LawnGreen;
                            row.Cells[4].Style.BackColor = System.Drawing.Color.LawnGreen;
                        }
						if ((kol) < (norm*0.9))
                        {
                            //                            row.Cells[6].Style.BackColor = System.Drawing.Color.LawnGreen;
                            row.Cells[5].Style.BackColor = System.Drawing.Color.Red;
                            row.Cells[4].Style.BackColor = System.Drawing.Color.Red;
                        }

                    }
                }
                dgv_norme.AutoResizeColumns();
                dgv_norme.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                //


                sql = "select month(p.datum) Mjesec, p.datum,p.dosao,p.otisao,p.hala,p.smjena,p.radnomjesto,p.napomena,p.kasni,p.preranootisao,p.ukupno_minuta from radnici_ r left join pregledvremena p on r.id=p.idradnika  where p.idradnika= " + ssindex + " and DATEDIFF(month,p.datum, GETDATE()) <= 12 order by datum desc";

                connection = new SqlConnection(connectionString);
                dataadapter = new SqlDataAdapter(sql, connection);
                ds = new DataSet();
                connection.Open();
                dataadapter.Fill(ds, "event");
                connection.Close();

                dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dataGridView1.DataSource = ds;
                dataGridView1.DataMember = "event";

                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (row.Cells[1].Value == null)
                    { }
                    else
                    {
                        string s1 = row.Cells[1].Value.ToString();
                        string smjena1 = row.Cells[5].Value.ToString().TrimEnd();

                        DateTime dat0 = DateTime.Parse(s1);
                        DayOfWeek dat1 = dat0.DayOfWeek;

                        foreach (var item in praznicii)                   // pregledvemena praznici
                        {
                            if (item.dan.ToString() == s1)
                                row.DefaultCellStyle.BackColor = System.Drawing.Color.LightPink;
                        }


                        if (dat1 == DayOfWeek.Saturday)
                        {
                            if (smjena1 != "3")
                            {
                                row.DefaultCellStyle.BackColor = System.Drawing.Color.Yellow;
                            }
                            else  // subota 3.smjena
                            {
                                row.DefaultCellStyle.BackColor = System.Drawing.Color.LightSeaGreen;
                            }
                        }
                        if (dat1 == DayOfWeek.Sunday)
                        {
                            row.DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                        }
                    }
                }
                dataGridView1.AutoResizeColumns();
                dataGridView1.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);


                if ((idloged == "8") || (idadm == "14"))   // oni koji mogu vidjeti sve i tehnički direktor mogu vidjeti satnice
                {
                    sql = "select Firma,vrstaRM,MT,Satnicakn,Dandolaska,Danodlaska,Godina,mjesec,Dan01,dan02,dan03,dan04,dan05,dan06,dan07,dan08,dan09,dan10,dan11,dan12,dan13,dan14,dan15,dan16,dan17,dan18,dan19,dan20,dan21,dan22,dan23,dan24,dan25,dan26,dan27,dan28,dan29,dan30,dan31,stimulacija1,stimulacija2,stimulacija3,stimulacija4,stimulacija5,rfid from fxsap.dbo.plansatirada where radnikid=" + ssindex1 + " and firma= " + poduzece1 + " order by godina desc";
                }
                else
                {
                    sql = "select Firma,vrstaRM,MT,Satnicakn,Dandolaska,Danodlaska,Godina,mjesec,Dan01,dan02,dan03,dan04,dan05,dan06,dan07,dan08,dan09,dan10,dan11,dan12,dan13,dan14,dan15,dan16,dan17,dan18,dan19,dan20,dan21,dan22,dan23,dan24,dan25,dan26,dan27,dan28,dan29,dan30,dan31,stimulacija1,stimulacija2,stimulacija3,stimulacija4,stimulacija5,rfid from fxsap.dbo.plansatirada where radnikid=" + ssindex1 + " and firma= " + poduzece1 + " order by godina desc";
                    //sql = "select Firma,vrstaRM,MT,Dandolaska,Danodlaska,Godina,mjesec,Dan01,dan02,dan03,dan04,dan05,dan06,dan07,dan08,dan09,dan10,dan11,dan12,dan13,dan14,dan15,dan16,dan17,dan18,dan19,dan20,dan21,dan22,dan23,dan24,dan25,dan26,dan27,dan28,dan29,dan30,dan31,stimulacija1,stimulacija2,stimulacija3,stimulacija4,stimulacija5,rfid from fxsap.dbo.plansatirada where radnikid=" + ssindex1 + " and firma= " + poduzece1 + " order by godina desc";
                }

                connection = new SqlConnection(connectionStringFeroapp);
                dataadapter = new SqlDataAdapter(sql, connection);
                DataSet ds2 = new DataSet();
                connection.Open();
                dataadapter.Fill(ds2, "event2");
                connection.Close();

                dgv_psr.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dgv_psr.DataSource = ds2;
                dgv_psr.DataMember = "event2";
                //object row = this.dgv_psr.Items[0]; //Grab the first row
                //DataGridColumn col = this.dgv_psr.Columns[this.dgv_psr.Columns.Count - 1]; //Grab the last column
                //this.dgv_psr.ScrollIntoView(row, col); //Set the view

                dgv_psr.AutoResizeColumns();
                dgv_psr.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                int bc = dgv_psr.ColumnCount;
                foreach (DataGridViewRow row1 in dgv_psr.Rows)
                {
                    if (row1.Cells[1].Value == null)
                    { }
                    else
                    {
                        for (int i = 5; i < bc; i++)
                        {

                            if (i > 8)
                            {

                                int g1 = (int.Parse)(row1.Cells[6].Value.ToString());
                                int m1 = int.Parse(row1.Cells[7].Value.ToString());
                                int d1 = i - 8;
                                DateTime dat0 = new DateTime();

                                if (m1 < 12)
                                {
                                    dat0 = new DateTime(g1, m1 + 1, 1);
                                }
                                else
                                {
                                    dat0 = new DateTime(g1 + 1, 1, 1);
                                }

                                if (dat0.AddDays(-1).Day >= d1)
                                {
                                    DateTime dat11 = new DateTime(g1, m1, d1);


                                    DayOfWeek dat1 = dat11.DayOfWeek;

                                    if (dat1 == DayOfWeek.Saturday)
                                    {
                                        row1.Cells[i - 1].Style.BackColor = System.Drawing.Color.LightYellow;

                                    }
                                    if (dat1 == DayOfWeek.Sunday)
                                    {
                                        row1.Cells[i - 1].Style.BackColor = System.Drawing.Color.LightGreen;
                                    }
                                }
                            }

                            if (row1.Cells[i].Value.ToString().Contains("g"))
                                row1.Cells[i].Style.BackColor = System.Drawing.Color.LawnGreen;

                            if (row1.Cells[i].Value.ToString().Contains("b"))
                                row1.Cells[i].Style.BackColor = System.Drawing.Color.LightCyan;

                            if (row1.Cells[i].Value.ToString().Contains("e"))
                                row1.Cells[i].Style.BackColor = System.Drawing.Color.LightPink;


                            if ((idloged == "8") || (idadm == "14"))
                            {

                            }
                            else
                            {
                                row1.Cells[3].Value = 0.0;
                            }

                            if (i == 5)
                            {
                                if (row1.Cells[5].Value.ToString().Contains("."))
                                {
                                    row1.Cells[5].Style.BackColor = System.Drawing.Color.Red;
                                    label30.Text = label30.Text + " >> Prestanak radnog odnosa:" + row1.Cells[5].Value.ToString();
                                }
                            }
                        }

                    }
                }

                
                cnn1.Open();
                SqlCommand cmd1= new SqlCommand("insert into kks_log (datum,korisnik,idprijave,opis) values  ( getdate(),'" + korisnik + "','" + idprijave + "','Pregled O.K. za ssindex1="+ssindex+" poduzeće "+poduzece1+"')", cnn1);
                SqlDataReader reader1= cmd1.ExecuteReader();
                cnn1.Close();

                string cnn11 ="";
                if (poduzece1 == "1")
                {
                    cnn11 = connectionStringpa;
                }
                else
                {
                    cnn11 = connectionStringpb;
                }

                cnn1.Open();
                cmd1 = new SqlCommand("select id_radnika from radnici_ where id="+ssindex.ToString()+" and poduzece ='"+poduzece1s+"'", cnn1);
                reader1 = cmd1.ExecuteReader();
                
                while (reader1.Read())
                {
                    idradnika1 = (int.Parse)(reader1["id_radnika"].ToString());
                }



                    cnn1.Close();

                

                // izračun staža kod prekida radnog odnosa manjeg od dva dana

                SqlConnection cnn111 = new SqlConnection(cnn11);
                cnn111.Open();
                string sql11 = "select * from( select p.acworker, addate, adDateEnd from thr_prsn p left join thr_prsnjob j on p.acworker=j.acworker  where p.acregno="+ssindex+") x1 order by addate " ;
                SqlCommand cmd11 = new SqlCommand(sql11, cnn111);
                SqlDataReader reader11 = cmd11.ExecuteReader();
                int rbr = 1;
                DateTime dato1 = new DateTime();
                DateTime dato2 = new DateTime();
                DateTime datz1 = new DateTime();
                DateTime datppp = dato1 ;
                string datumz = "",datumo="";

                while (reader11.Read())
                {

                    if (rbr!=0)
                    {
                       // datumz = reader11["addate"].ToString();
                        datumo = reader11["addateend"].ToString();
                        //string[] datz1p = datumz.Split('.');
                        //datz1 = new DateTime((int.Parse)(datz1p[2]), (int.Parse)(datz1p[1]), (int.Parse)(datz1p[0]));
                        string[] datz1p = datumo.Split('.') ;
                        if ((datz1p.Length > 1) && (rbr != 0))
                            dato1 = new DateTime((int.Parse)(datz1p[2]), (int.Parse)(datz1p[1]), (int.Parse)(datz1p[0])) ;
                        else
                            dato1 = DateTime.Now ;
                    }
                    if (rbr==1)
                    {
                        datumz = reader11["addate"].ToString();
                        datumo = reader11["addateend"].ToString();
                        string[] datz1p = datumz.Split('.');
                        datz1 = new DateTime((int.Parse)(datz1p[2]), (int.Parse)(datz1p[1]), (int.Parse)(datz1p[0]));
                        datz1p = datumo.Split('.');
                        datppp = dato1;
                        string[] dato2p = datumo.Split('.');
                        if (dato2p.Length > 1)
                            dato2 = new DateTime((int.Parse)(dato2p[2]), (int.Parse)(dato2p[1]), (int.Parse)(dato2p[0]));
                        else
                            dato2 = DateTime.Now;
                    }
                                        
                    //DateTime datz = 
                    if ( rbr!=1)
                    {
                        string datumz2 = reader11["addate"].ToString();
                        string datumo2 = reader11["addateend"].ToString();
                        string[] datz2p = datumz2.Split('.');
                        string[] dato2p = datumo2.Split('.');
                        DateTime datz2 = new DateTime((int.Parse)(datz2p[2]), (int.Parse)(datz2p[1]), (int.Parse)(datz2p[0]));
                        if (dato2p.Length > 1)
                            dato2 = new DateTime((int.Parse)(dato2p[2]), (int.Parse)(dato2p[1]), (int.Parse)(dato2p[0]));
                        else
                            dato2 = DateTime.Now;
                        
                        //datz1p = datumo2.Split('.');
                        //if (datz1p.Length > 1)
                        //    dato1 = new DateTime((int.Parse)(datz1p[2]), (int.Parse)(datz1p[1]), (int.Parse)(datz1p[0]));

                        int danq = datz2.Subtract(datppp).Days;
                        if (danq>2)   // ako je prekid u radnom odnosu veći od dva dana
                        {
                            datumz = datumz2;                            
                        }
                        datumo = datumo2;
                    }
                    datppp = dato1;
                    rbr++;
                }
                cnn111.Close();
                string staz1 = staz(datz1, dato2);
                cnn1.Open();
                SqlCommand cmd = new SqlCommand("dbo.sp_Kompetencije1", cnn1);
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@id1", SqlDbType.Int, 5).Value = ssindex1;  // dodati parametar godinu i mjesec
                cmd.Parameters.Add("@pod1", SqlDbType.Int, 5).Value = poduzece1;  // dodati parametar godinu i mjesec
                //cmd.Parameters.Add("@datum1", SqlDbType.DateTime, 5).Value = datum1;  // dodati parametar godinu i mjesec
                //cmd.Parameters.Add("@datum2", SqlDbType.DateTime, 5).Value = datum2;  // dodati parametar godinu i mjesec
                SqlDataReader rdr = null;
                SqlConnection cnn = null;
                rdr = cmd.ExecuteReader();
                string string1 = "", string2 = "", string3 = "";
                string radnikid, mjesec1, ime1 = "", godina1 = "", firma1 = "";
                int jos = 1;

                bool joss = false;

                rdr.Read();
                if (rdr.HasRows)
                { }
                else
                {
                    MessageBox.Show("Nema podataka za dane uvjete u kompetencijama  !!!!!!!!!");
                    Ocisti();
                    return;
                }


                label24.Text = rdr["prezimeime"].ToString();
                //string[] sdatz = rdr["datumzaposlenja"].ToString().Split(' ');
                //sdatz = datumz.Split(' ');

                //label25.Text = sdatz[0];
                label25.Text = datumz;
                label128.Text = rdr["napomena"].ToString();
                label26.Text = rdr["Funkcija"].ToString();
                tb_vjestina.Visible = true;
                tb_vjestina.Text = rdr["vještine"].ToString();  // vjestina staro
                //label27.Text = rdr["vjestina"].ToString();
                label28.Text = rdr["mjesto_troska"].ToString();
                label113.Text = rdr["Godisnji_ostalo"].ToString();
                if (rdr["Hala"].ToString().Trim().Length > 0)
                {
                    label29.Text = "Hala" + rdr["Hala"].ToString() + " - " + rdr["radnomjesto"].ToString();
                }
                else
                {
                    label29.Text = rdr["radnomjesto"].ToString();
                }


                if (label30.Text.Contains("Prest"))
                {
                    label30.Text = "Prekid radnog odnosa: "+datumo;
                }
                else
                {
                    label30.Text = rdr["istek_ugovora"].ToString();
                }
                label31.Text = rdr["staz"].ToString();
                label31.Text = staz1 ;

                if (satnica == "1")   // dali smije vidjeti satnicu
                { 
                label32.Text = rdr["satnicastara"].ToString();
                label33.Text = rdr["satnicabruto"].ToString();
                label34.Text = rdr["satnicanovaod"].ToString();
                }

                // 1 mjesec
                label35.Text = rdr["nedolascinedjeljom1"].ToString();
                label36.Text = rdr["nedolascisubotom1"].ToString();
                label37.Text = rdr["bolovanje_broj1"].ToString();
                label38.Text = rdr["bolovanja_dana1"].ToString();
                label39.Text = rdr["stimulacija1"].ToString();
                label40.Text = rdr["nedostajedo8sati1"].ToString();
                label41.Text = rdr["kasni1"].ToString();
                label42.Text = rdr["preranootisao1"].ToString();
                label43.Text = rdr["neopravdanidani1"].ToString();
                label44.Text = rdr["normaposto1"].ToString();
                label46.Text = rdr["Dolascinedjeljom1"].ToString();
                label95.Text = rdr["Dolascipraznikom1"].ToString();
                label85.Text = rdr["Godisnji_dana1"].ToString();

                // 3.mjesec
                label57.Text = rdr["nedolascinedjeljom3"].ToString();
                label56.Text = rdr["nedolascisubotom3"].ToString();
                label55.Text = rdr["bolovanje_broj3"].ToString();
                label54.Text = rdr["bolovanja_dana3"].ToString();
                label53.Text = rdr["stimulacija3"].ToString();
                label52.Text = rdr["nedostajedo8sati3"].ToString();
                label51.Text = rdr["kasni3"].ToString();
                label50.Text = rdr["preranootisao3"].ToString();
                label49.Text = rdr["neopravdanidani3"].ToString();
                label48.Text = rdr["normaposto3"].ToString();
                label45.Text = rdr["Dolascinedjeljom3"].ToString();
                label96.Text = rdr["Dolascipraznikom3"].ToString();
                label84.Text = rdr["Godisnji_dana3"].ToString();

                // 6.mjesec
                label68.Text = rdr["nedolascinedjeljom6"].ToString();
                label67.Text = rdr["nedolascisubotom6"].ToString();
                label66.Text = rdr["bolovanje_broj6"].ToString();
                label65.Text = rdr["bolovanja_dana6"].ToString();
                label64.Text = rdr["stimulacija6"].ToString();
                label63.Text = rdr["nedostajedo8sati6"].ToString();
                label62.Text = rdr["kasni6"].ToString();
                label61.Text = rdr["preranootisao6"].ToString();
                label60.Text = rdr["neopravdanidani6"].ToString();
                label59.Text = rdr["normaposto6"].ToString();
                label23.Text = rdr["Dolascinedjeljom6"].ToString();
                label97.Text = rdr["Dolascipraznikom6"].ToString();
                label73.Text = rdr["Godisnji_dana6"].ToString();


                // 12.mjesec
                label83.Text = rdr["nedolascinedjeljom12"].ToString();
                label82.Text = rdr["nedolascisubotom12"].ToString();
                label81.Text = rdr["bolovanje_broj12"].ToString();
                label80.Text = rdr["bolovanja_dana12"].ToString();
                label79.Text = rdr["stimulacija12"].ToString();
                label78.Text = rdr["nedostajedo8sati12"].ToString();
                label77.Text = rdr["kasni6"].ToString();
                label76.Text = rdr["preranootisao12"].ToString();
                label75.Text = rdr["neopravdanidani12"].ToString();
                label74.Text = rdr["normaposto12"].ToString();
                label22.Text = rdr["Dolascinedjeljom12"].ToString();
                label98.Text = rdr["Dolascipraznikom12"].ToString();
                label58.Text = rdr["Godisnji_dana12"].ToString();

                while (joss)
                {
                    // reader["Datum"].ToString();
                    if (rdr.HasRows)
                    {
                        if (rdr["id"] is DBNull)
                        {
                            joss = false;
                            continue;
                        }

                        label24.Text = rdr["prezimeime"].ToString();


                    }
                }
            }
         }

        #endregion lista djelatnika 



        private void updatePodatakaRučnoToolStripMenuItem_Click(object sender, EventArgs e)
        {

            // int z = 0;
            if (idadm == "8")
            {
                button1.Visible = true;
                button1.PerformClick();
                MessageBox.Show("Gotovo !!");
                button1.Visible = false;
            }

        }
        #region export from datagriview zbirni
        private void btn_zbirni_export_Click(object sender, EventArgs e)
        {
            //Creating DataTable
            DataTable dt = new DataTable();

            //Adding the Columns
            foreach (DataGridViewColumn column in dgv_zbirni.Columns)
            {

                // satnica se pokazuje samo za grupu 8 ili tehn. direktora (14)
                dt.Columns.Add(column.HeaderText, column.ValueType);


            }

            //Adding the Rows
            foreach (DataGridViewRow row in dgv_zbirni.Rows)
            {
                dt.Rows.Add();

                foreach (DataGridViewCell cell in row.Cells)
                {


                    if (!(cell.Value is DBNull))
                    {
                        if (cell.Value != null)

                        {
                            dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();

                        }
                    }
                }
            }

            //if ((idloged == "8") || (idadm == "14"))
            //{

            //}
            //else
            //{
            //    // satnica se pokazuje samo za grupu 8 ili tehn. direktora (14)
            //    dt.Columns.RemoveAt(9);
            //    dt.Columns.RemoveAt(9);
            //    dt.Columns.RemoveAt(9);
            //    dt.Columns.RemoveAt(9);
            //}
            //Exporting to Excel
            string folderPath = "C:\\KKS\\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "Kompetencije");
                wb.SaveAs(folderPath + "Kompetencije.xlsx");
            }
            btn_zbirni_export.Text = "Export done";

            FileInfo fi = new FileInfo("C:\\kks\\kompetencije.xlsx");
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(@"C:\kks\kompetencije.xlsx");
            }
            else
            {
                //file doesn't exist
            }

            //Object oMissing = System.Reflection.Missing.Value;
            //oMissing = System.Reflection.Missing.Value;
            //Object oTrue = true;
            //Object oFalse = false;
            //Microsoft.Office.Interop.Word.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Document oExcelDoc = new Microsoft.Office.Interop.Excel.Document();
            //oExcel.Visible = true;

            //Object oTemplatePath = "c:\\kks\\kompetencije.xlsx";
            //oExcelDoc = oExcelDoc.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);





        }
        #endregion

        #region ocisti, isprazni labals
        private void OcistiPanele()
        {
            pnl_odredeno.Visible = false;
            pnl_norme.Visible = false;
            dgv_norme.Visible = false;
            pnl_vjestine.Visible = false;
            pl_zbirni.Visible = false;
            pnl_aktivnost.Visible = false;
            pnl_vjestine.Visible = false;
            panel1.Visible = false;
            pl_psr.Visible = false;
            pnl_škola.Visible = false;
            dgv_psr.Visible = false;
            pnl_vjestine_spajanje.Visible = false;
            pnl_dnevizv_norme.Visible = false;
            dgv_dnev_izv_norme.Visible = false;
            pnlv_log.Visible = false;
            pnl_projektmtz.Visible = false;
            pnl_naodlasku.Visible = false;
            pnl_novi_djelatnici.Visible = false;
            pnl_preglSkolovanja.Visible = false;
            pl_tvporuke.Visible = false;
            pnl_psr_izmjena.Visible = false;
        }
        private void Ocisti()
        {
            label102.BackColor = System.Drawing.Color.LightSeaGreen;
            pnl_odredeno.Visible = false;
            label85.Text = "";
            label58.Text = "";
            label84.Text = "";
            label83.Text = "";
            label58.Text = "";
            label46.Text = "";
            label45.Text = "";
            label73.Text = "";
            label22.Text = "";

            label24.Text = "";
            label25.Text = "";
            label128.Text = "";
            label26.Text = "";
            //label27.Text = "";
            tb_vjestina.Text = "";
            tb_vjestina.Visible = false;
            label28.Text = "";
            label29.Text = "";
            label30.Text = "";
            label31.Text = "";
            label32.Text = "";
            label33.Text = "";
            label34.Text = "";

            // 1 mjesec
            label35.Text = "";  // nedjelja
            label95.Text = "";  // praznik
            label36.Text = "";  // nedolasci nedjeljom
            label37.Text = "";  // dolasci subotom
            label38.Text = "";  // nedolasci subotom
            label39.Text = "";
            label40.Text = "";
            label41.Text = "";
            label42.Text = "";
            label43.Text = "";
            label44.Text = "";

            // 3.mjesec
            label57.Text = "";
            label96.Text = "";  // praznik
            label56.Text = "";
            label55.Text = "";
            label54.Text = "";
            label53.Text = "";
            label52.Text = "";
            label51.Text = "";
            label50.Text = "";
            label49.Text = "";
            label48.Text = "";


            // 6.mjesec
            label68.Text = "";
            label97.Text = "";  // praznik
            label67.Text = "";
            label66.Text = "";
            label65.Text = "";
            label64.Text = "";
            label63.Text = "";
            label62.Text = "";
            label61.Text = "";
            label60.Text = "";
            label59.Text = "";


            // 12.mjesec
            label83.Text = "";
            label98.Text = "";  // praznik
            label82.Text = "";
            label81.Text = "";
            label80.Text = "";
            label79.Text = "";
            label78.Text = "";
            label77.Text = "";
            label76.Text = "";
            label75.Text = "";
            label74.Text = "";

        }
        #endregion

        #region izlaz
        private void izlazToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SqlConnection cnn1 = new SqlConnection(connectionString);
            cnn1.Open();
            SqlCommand cmd1 = new SqlCommand("insert into kks_log (datum,korisnik,idprijave,opis) values  ( getdate(),'" + korisnik + "','" + idprijave + "','Odjava')", cnn1);
            SqlDataReader reader1 = cmd1.ExecuteReader();
            cnn1.Close();
            Application.Exit();
        }
        #endregion 

        private void dgv_psr_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // klik sa npoamenama radi sadmo u 1. koloni

            int row1 = e.RowIndex;
            int col1 = e.ColumnIndex;

            if (col1 == 0)
            {
                pnl_napomene.Visible = true;

                DataGridViewRow selectedRow = dgv_psr.Rows[row1];
                //            ssindex1 
                g1 = (int.Parse)(Convert.ToString(selectedRow.Cells["Godina"].Value));
                m1 = (int.Parse)(Convert.ToString(selectedRow.Cells["Mjesec"].Value));
                //int d1 = col1 - 7;
                //DateTime dat1 = new DateTime(g1, m1, d1);

                string sql11 = "select n.DanNapomene,n.napomena from fxsap.dbo.plansatirada p left join fxsap.dbo.plansatiradanapomene n on p.psrid=n.psrid where year(n.dannapomene)=" + g1.ToString() + " and month(n.dannapomene)=" + m1.ToString() + " and radnikid=" + ssindex1.ToString();


                SqlConnection connection = new SqlConnection(connectionString);
                SqlDataAdapter dataadapter = new SqlDataAdapter(sql11, connection);
                DataSet ds = new DataSet();
                connection.Open();
                dataadapter.Fill(ds, "event3");
                connection.Close();

                dgv_napomene.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dgv_napomene.DataSource = ds;
                dgv_napomene.DataMember = "event3";

                dgv_napomene.AutoResizeColumns();
                dgv_napomene.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

            }
            if ((col1 >7) && ( idloged=="8"))
            {
                pnl_psr_izmjena.Visible = true;

                DataGridViewRow selectedRow = dgv_psr.Rows[row1];
                //            ssindex1 
                 g1 = (int.Parse)(Convert.ToString(selectedRow.Cells["Godina"].Value));
                 m1 = (int.Parse)(Convert.ToString(selectedRow.Cells["Mjesec"].Value));
                //string danx;
                if ((col1 - 7) < 10)
                {
                    danx = "dan0" + (col1 - 7).ToString();
                }
                else
                {
                    danx = "dan" + (col1 - 7).ToString();
                }
                


                //int d1 = col1 - 7;
                //DateTime dat1 = new DateTime(g1, m1, d1);

                string sql11 = "select "+danx +" from fxsap.dbo.plansatirada p where godina=" + g1.ToString() + " and mjesec=" + m1.ToString() + " and radnikid=" + ssindex1.ToString();
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();                    
                    string sql1 = "select " + danx + " from fxsap.dbo.plansatirada p where godina=" + g1.ToString() + " and mjesec=" + m1.ToString() + " and radnikid=" + ssindex1.ToString();

                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string oldvalue = "";
                    while (reader.Read())
                    {
                        oldvalue = reader[danx].ToString();
                    }
                    cn.Close();
                    txtbox_sv.Text = oldvalue;
                    txtbox_nv.Text = oldvalue;
                }


               

            }

        }

        private void button2_Click(object sender, EventArgs e)
        {
            pnl_napomene.Visible = false;

        }

        private void dgv_napomene_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void cbx_mjestotroska_SelectedIndexChanged(object sender, EventArgs e)
        {
            panel1.Visible = true;
            pl_zbirni.Visible = false;
            pnl_norme.Visible = true;
            dgv_norme.Visible = true;

            using (SqlConnection cnn1 = new SqlConnection(connectionString))
            {
                if (cbx_mjestotroska.SelectedIndex <= 0)
                {
                    return;
                }
                int ssindex = int.Parse(cbx_mjestotroska.SelectedValue.ToString());
                ssindex1 = ssindex;
            }

            using (SqlConnection cnn1 = new SqlConnection(connectionString))
            {

                var dataSource = new List<radnici>();

                dataSource.Add(new radnici() { prezime = " - ", id = "0" });

                foreach (var radnikk in radnicii)
                {
                    if (radnikk.mt == ssindex1.ToString())
                    {
                        dataSource.Add(new radnici() { prezime = radnikk.prezime + " " + radnikk.ime + " - " + radnikk.id.ToString() + " ", id = radnikk.id });
                    }
                }

                combo_listadjelatnika.MaxDropDownItems = 60;
                this.combo_listadjelatnika.DataSource = dataSource;
                this.combo_listadjelatnika.DisplayMember = "prezime";
                this.combo_listadjelatnika.ValueMember = "id";
            }
            //

        }

        private void vjestineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (idadm != "8")
            {
                return;
            }
            
            pnl_vjestine_spajanje.Visible = false;
            panel1.Visible = false;
            pnl_vjestine.Visible = true;
            panel1.Visible = false;
            pl_zbirni.Visible = false;
            pl_psr.Visible = false;
            dgv_psr.Visible = false;

            string sql11 = "select * from vjestine";

            SqlConnection connection = new SqlConnection(connectionString);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql11, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event3");
            connection.Close();

            dgv_vjestine.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgv_vjestine.DataSource = ds;
            dgv_vjestine.DataMember = "event3";

            dgv_vjestine.AutoResizeColumns();
            dgv_vjestine.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

        }

        private void administracijaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Ocisti();
            OcistiPanele();

        }

        private void label104_Click(object sender, EventArgs e)
        {

        }

        //tablica sa normama
        private void dgv_norme_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // klik sa napomenama radi sadmo u 1. koloni

            int row1 = e.RowIndex;
            int col1 = e.ColumnIndex;

                pnl_napomene.Visible = true;

                DataGridViewRow selectedRow = dgv_norme.Rows[row1];
                //            ssindex1 
                string sdat1   = (Convert.ToString(selectedRow.Cells["Datum"].Value))  ;
                string slinija = (Convert.ToString(selectedRow.Cells["Linija"].Value)) ;
                string shala   = (Convert.ToString(selectedRow.Cells["Hala"].Value))   ;
                string vrijemeod = (Convert.ToString(selectedRow.Cells["VrijemeOd"].Value));
                string vrijemedo = (Convert.ToString(selectedRow.Cells["VrijemeDo"].Value));
            DateTime dat1 = DateTime.Parse(sdat1);
              int z = 0;
            //if ( 1 ==2 )
            //{ 

            //    m1 = (int.Parse)(Convert.ToString(selectedRow.Cells["Mjesec"].Value));
            //    //int d1 = col1 - 7;
            //    //DateTime dat1 = new DateTime(g1, m1, d1);

            //    string sql11 = "select n.DanNapomene,n.napomena from fxsap.dbo.plansatirada p left join fxsap.dbo.plansatiradanapomene n on p.psrid=n.psrid where year(n.dannapomene)=" + g1.ToString() + " and month(n.dannapomene)=" + m1.ToString() + " and radnikid=" + ssindex1.ToString();


            //    SqlConnection connection = new SqlConnection(connectionString);
            //    SqlDataAdapter dataadapter = new SqlDataAdapter(sql11, connection);
            //    DataSet ds = new DataSet();
            //    connection.Open();
            //    dataadapter.Fill(ds, "event3");
            //    connection.Close();

            //    dgv_napomene.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            //    dgv_napomene.DataSource = ds;
            //    dgv_napomene.DataMember = "event3";

            //    dgv_napomene.AutoResizeColumns();
            //    dgv_napomene.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);


            //    if ((col1 > 7) && (idloged == "8"))
            //    {
            //        pnl_psr_izmjena.Visible = true;

            //        DataGridViewRow selectedRow = dgv_psr.Rows[row1];
            //        //            ssindex1 
            //        g1 = (int.Parse)(Convert.ToString(selectedRow.Cells["Godina"].Value));
            //        m1 = (int.Parse)(Convert.ToString(selectedRow.Cells["Mjesec"].Value));
            //        //string danx;
            //        if ((col1 - 7) < 10)
            //        {
            //            danx = "dan0" + (col1 - 7).ToString();
            //        }
            //        else
            //        {
            //            danx = "dan" + (col1 - 7).ToString();
            //        }



            //        //int d1 = col1 - 7;
            //        //DateTime dat1 = new DateTime(g1, m1, d1);

            //        string sql11 = "select " + danx + " from fxsap.dbo.plansatirada p where godina=" + g1.ToString() + " and mjesec=" + m1.ToString() + " and radnikid=" + ssindex1.ToString();
            //        using (SqlConnection cn = new SqlConnection(connectionString))
            //        {
            //            cn.Open();
            //            string sql1 = "select " + danx + " from fxsap.dbo.plansatirada p where godina=" + g1.ToString() + " and mjesec=" + m1.ToString() + " and radnikid=" + ssindex1.ToString();

            //            SqlCommand sqlCommand = new SqlCommand(sql1, cn);
            //            SqlDataReader reader = sqlCommand.ExecuteReader();
            //            string oldvalue = "";
            //            while (reader.Read())
            //            {
            //                oldvalue = reader[danx].ToString();
            //            }
            //            cn.Close();
            //            txtbox_sv.Text = oldvalue;
            //            txtbox_nv.Text = oldvalue;
            //        }

            //    }


            //}

        }

        private void mjesečniPregledAktivnostiToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OcistiPanele();
            pnl_aktivnost.Visible = true;
            pnl_vjestine.Visible = false;
            pl_psr.Visible = false;
            pnl_vjestine_spajanje.Visible = false;
            pl_zbirni.Visible = false;

        }

        #region mjesečnaaktivnost

        private void cbx_godina_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void btn_a_search_Click(object sender, EventArgs e)
        {

            btn_a_export.Visible = true;

            string mjesec1 = (cbx_mjesec.SelectedIndex + 1).ToString();
            int godina1 = cbx_godina.SelectedIndex;

            int m1 = (int.Parse)(mjesec1);
            int g1 = 2016 + godina1;

            string godinas = (2016 + godina1).ToString();
            string poduzece1 = "1"; // ????????????????????
            string hala1 = "1";

            if (idadm == "1")
            {
                hala1 = "1";
            }
            else if (idadm == "3")
            {
                hala1 = "3";
            }

            string sql = "";
            string trazi = "";

            if (prikazibolovanje == 1)
                trazi = "b";

            if (prikazigodisnji == 1)
                trazi = "g";

            if (prikaziizostanak == 1)
                trazi = "e";

            string sql2 = " 1=1";

            if (trazi.Length > 0)
            {
                sql2 = " charindex('" + trazi + "',isnull(dan01,'')+isnull(dan02,'')+isnull(dan03,'')+isnull(dan04,'')+isnull(dan05,'')+isnull(dan06,'')+isnull(dan07,'')+isnull(dan08,'')+isnull(dan09,'')+isnull(dan10,'')+isnull(dan11,'')+isnull(dan12,'')+isnull(dan13,'')+isnull(dan14,'')+isnull(dan15,'')+isnull(dan16,'')+isnull(dan17,'')+isnull(dan18,'')+isnull(dan19,'')+isnull(dan20,'')+isnull(dan21,'')+isnull(dan22,'')+isnull(dan23,'')+isnull(dan24,'')+isnull(dan25,'')+isnull(dan26,'')+isnull(dan27,'')+isnull(dan28,'')+isnull(dan29,'')+isnull(dan30,'')+isnull(dan31,'')) >0 ";
            }

            if ((idloged == "8") || (idloged == "7"))   // oni koji mogu vidjeti sve i tehnički direktor mogu vidjeti satnice
            {

                if (idloged == "14")  // tehnički
                {
                    sql = "select Firma, Radnikid, p.Ime,Dan01,dan02,dan03,dan04,dan05,dan06,dan07,dan08,dan09,dan10,dan11,dan12,dan13,dan14,dan15,dan16,dan17,dan18,dan19,dan20,dan21,dan22,dan23,dan24,dan25,dan26,dan27,dan28,dan29,dan30,dan31,0 podbacaja,0 prebacaja,stimulacija1,stimulacija2,stimulacija3,stimulacija4,stimulacija5 from fxsap.dbo.plansatirada p left join radnici_ r on r.id = p.radnikid left join mjestotroska mt on mt.id = r.mt where mjesec =" + mjesec1 + " and godina = " + godinas + " and mt.grupa1 in (1,2,3,5,13)   "+sql2+"  order by ime";
                }
                else   // svi
                {
                    sql = "select Firma,r.mt, Radnikid, p.Ime,Dan01,dan02,dan03,dan04,dan05,dan06,dan07,dan08,dan09,dan10,dan11,dan12,dan13,dan14,dan15,dan16,dan17,dan18,dan19,dan20,dan21,dan22,dan23,dan24,dan25,dan26,dan27,dan28,dan29,dan30,dan31,0 podbacaja,0 prebacaja,stimulacija1,stimulacija2,stimulacija3,stimulacija4,stimulacija5 from fxsap.dbo.plansatirada p left join radnici_ r on r.id = p.radnikid left join mjestotroska mt on mt.id = r.mt where mjesec =" + mjesec1 + " and godina = " + godinas + " and "+sql2+" order by ime";
                }

                //sql = "select Firma,Radnikid,Ime,vrstaRM,Satnicakn,Dandolaska,Dan01,dan02,dan03,dan04,dan05,dan06,dan07,dan08,dan09,dan10,dan11,dan12,dan13,dan14,dan15,dan16,dan17,dan18,dan19,dan20,dan21,dan22,dan23,dan24,dan25,dan26,dan27,dan28,dan29,dan30,dan31,stimulacija1,stimulacija2,stimulacija3,stimulacija4,stimulacija5,rfid from fxsap.dbo.plansatirada where mjesec=" + mjesec1 + " and godina=" + godinas + " order by ime ";
                //sql = "select Firma,Radnikid,Ime,vrstaRM,Satnicakn,Dandolaska,Dan01,dan02,dan03,dan04,dan05,dan06,dan07,dan08,dan09,dan10,dan11,dan12,dan13,dan14,dan15,dan16,dan17,dan18,dan19,dan20,dan21,dan22,dan23,dan24,dan25,dan26,dan27,dan28,dan29,dan30,dan31,stimulacija1,stimulacija2,stimulacija3,stimulacija4,stimulacija5,rfid from fxsap.dbo.plansatirada where mjesec=" + mjesec1 + " and godina=" + godinas + " and firma= " + poduzece1 + " order by ime desc";
                //sql = "select e.datum,e.kolicinaok,e.norma,e.linija,e.smjena,Firma,Radnikid,p.Ime,vrstaRM,godina,mjesec,Satnicakn,Dandolaska,Dan01,dan02,dan03,dan04,dan05,dan06,dan07,dan08,dan09,dan10,dan11,dan12,dan13,dan14,dan15,dan16,dan17,dan18,dan19,dan20,dan21,dan22,dan23,dan24,dan25,dan26,dan27,dan28,dan29,dan30,dan31,stimulacija1,stimulacija2,stimulacija3,stimulacija4,stimulacija5,rfid from feroapp.dbo.evidencijanormiview e left join feroapp.dbo.radnici r on r.id_radnika = e.id_radnika left join fxsap.dbo.plansatirada p on p.radnikid = r.ID_Fink and mjesec = month(e.datum) and godina = year(e.datum) where mjesec = " + mjesec1 + " and godina ="+godina1+" and firma = 1 and e.hala = "+ hala1 + "and p.radnikid = r.id_fink and r.id_radnika = e.id_radnika order by p.ime,e.datum";

            }
            else   // Kičin,gradiški...
            {
                {
                    sql = "select Firma, r.mt,Radnikid, p.Ime,Dan01,dan02,dan03,dan04,dan05,dan06,dan07,dan08,dan09,dan10,dan11,dan12,dan13,dan14,dan15,dan16,dan17,dan18,dan19,dan20,dan21,dan22,dan23,dan24,dan25,dan26,dan27,dan28,dan29,dan30,dan31,0 podbacaja,0 prebacaja,stimulacija1,stimulacija2,stimulacija3,stimulacija4,stimulacija5 from fxsap.dbo.plansatirada p left join radnici_ r on r.id = p.radnikid left join mjestotroska mt on mt.id = r.mt where mjesec =" +
                         mjesec1 + " and godina = " + godinas + " and mt.grupa1='" + idloged + "' and p.vrstarm='Proizvodnja' and r.neradi=0  and  " + sql2 + " order by ime";
                }

                // sql = "select Firma,Radnikid,Ime,vrstaRM,Satnicakn,Dandolaska,Dan01,dan02,dan03,dan04,dan05,dan06,dan07,dan08,dan09,dan10,dan11,dan12,dan13,dan14,dan15,dan16,dan17,dan18,dan19,dan20,dan21,dan22,dan23,dan24,dan25,dan26,dan27,dan28,dan29,dan30,dan31,stimulacija1,stimulacija2,stimulacija3,stimulacija4,stimulacija5,rfid from fxsap.dbo.plansatirada where mjesec=" + mjesec1 + " and godina=" + godinas + " order by ime";
                //sql = "select e.datum,e.kolicinaok,e.norma,e.linija,e.smjena,Firma,Radnikid,p.Ime,vrstaRM,godina,mjesec,Satnicakn,Dandolaska,Dan01,dan02,dan03,dan04,dan05,dan06,dan07,dan08,dan09,dan10,dan11,dan12,dan13,dan14,dan15,dan16,dan17,dan18,dan19,dan20,dan21,dan22,dan23,dan24,dan25,dan26,dan27,dan28,dan29,dan30,dan31,stimulacija1,stimulacija2,stimulacija3,stimulacija4,stimulacija5,rfid from feroapp.dbo.evidencijanormiview e left join feroapp.dbo.radnici r on r.id_radnika = e.id_radnika left join fxsap.dbo.plansatirada p on p.radnikid = r.ID_Fink and mjesec = month(e.datum) and godina = year(e.datum) where mjesec = " + mjesec1 + " and godina =" + godina1 + " and firma = 1 and e.hala = " + hala1 + "and p.radnikid = r.id_fink and r.id_radnika = e.id_radnika order by p.ime,e.datum";
                //sql = "select Firma,vrstaRM,MT,Dandolaska,Danodlaska,Godina,mjesec,Dan01,dan02,dan03,dan04,dan05,dan06,dan07,dan08,dan09,dan10,dan11,dan12,dan13,dan14,dan15,dan16,dan17,dan18,dan19,dan20,dan21,dan22,dan23,dan24,dan25,dan26,dan27,dan28,dan29,dan30,dan31,stimulacija1,stimulacija2,stimulacija3,stimulacija4,stimulacija5,rfid from fxsap.dbo.plansatirada where radnikid=" + ssindex1 + " and firma= " + poduzece1 + " order by godina desc";
            }

            SqlConnection connection = new SqlConnection(connectionString);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql, connection);
            DataSet ds2 = new DataSet();
            connection.Open();
            dataadapter.Fill(ds2, "event2");
            connection.Close();

            dgv_aktivnost.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgv_aktivnost.DataSource = ds2;
            dgv_aktivnost.DataMember = "event2";

            //FreezeBand(dgv_aktivnost.Columns[1]);
           // dgv_aktivnost.Columns[2].Frozen = true;
            dgv_aktivnost.AutoResizeColumns();
            dgv_aktivnost.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);            
            dgv_aktivnost.ReadOnly = true;


            int bc = dgv_aktivnost.ColumnCount;
            int z = 0;

            foreach (DataGridViewRow row1 in dgv_aktivnost.Rows)
            {

                z++;
                lbl_postoci.Text = z.ToString();
                if (row1.Cells[1].Value == null)
                { }
                else
                {
                    int idradnik1 = int.Parse(row1.Cells[2].Value.ToString());
                    int bpod = 0,bpre=0;

                    for (int i = 4; i < bc; i++)
                    {

                        DataGridViewColumn column = dgv_aktivnost.Columns[i];
                        column.Width = 40;
                        if (i > 3)
                        {
                            int d1 = i - 3;
                            DateTime dat0 = new DateTime();

                            if (m1 < 12)
                            {
                                dat0 = new DateTime(g1, m1 + 1, 1);
                            }
                            else
                            {
                                dat0 = new DateTime(g1 + 1, 1, 1);
                            }

                            if (dat0.AddDays(-1).Day >= d1)
                            {
                                DateTime dat11 = new DateTime(g1, m1, d1);

                                DayOfWeek dat1 = dat11.DayOfWeek;

                                if (dat1 == DayOfWeek.Saturday)
                                {
                                    row1.Cells[i].Style.BackColor = System.Drawing.Color.LightYellow;

                                }
                                if (dat1 == DayOfWeek.Sunday)
                                {
                                    row1.Cells[i].Style.BackColor = System.Drawing.Color.LightGreen;
                                }

                                if (prikazinormu == 1)
                                {
                                    string norma1 = NormaRadnika(idradnik1, dat11);
                                    row1.Cells[i].Value = norma1;
                                    if  (norma1.Contains("Podba"))
                                        {
                                        bpod++;
                                        row1.Cells[35].Value = bpod;
                                    };
                                    if (norma1.Contains("Preba"))
                                    {
                                        bpre++;
                                        row1.Cells[36].Value = bpre;
                                    };

                                    if ((bpre+bpod)>0)
                                    {
                                        if (bpre >= bpod )
                                        {
                                            row1.Cells[36].Style.BackColor = System.Drawing.Color.Green;
                                            row1.Cells[35].Style.BackColor = System.Drawing.Color.White;
                                        }
                                        else 
                                        {
                                            row1.Cells[35].Style.BackColor = System.Drawing.Color.Red;
                                            row1.Cells[36].Style.BackColor = System.Drawing.Color.White;
                                        }
                                    }

                                }

                                if (prikazibolovanje == 1)      // bolovanje
                                {
                                    string v1 = row1.Cells[i].Value.ToString();
                                    if (v1.Contains("b"))
                                    {

                                    }
                                    else
                                    {
                                        row1.Cells[i].Value = "";
                                    }
                                }



                                if (prikazigodisnji == 1)      // godišnji
                                {

                                    string v1 = row1.Cells[i].Value.ToString();
                                    if (v1.Contains("g"))
                                    {

                                    }
                                    else
                                    {
                                        row1.Cells[i].Value = "";
                                    }
                                }

                                if (prikaziizostanak == 1)      // izostanci
                                {

                                    string v1 = row1.Cells[i].Value.ToString();
                                    if (v1.Contains("e"))
                                    {

                                    }
                                    else
                                    {
                                        row1.Cells[i].Value = "";
                                    }
                                }



                            }
                        }

                        if (row1.Cells[i].Value.ToString().Contains("g"))
                            row1.Cells[i].Style.BackColor = System.Drawing.Color.LawnGreen;

                        if (row1.Cells[i].Value.ToString().Contains("b"))
                            row1.Cells[i].Style.BackColor = System.Drawing.Color.LightCyan;

                        if (row1.Cells[i].Value.ToString().Contains("e"))
                            row1.Cells[i].Style.BackColor = System.Drawing.Color.LightPink;


                        if ((idloged == "8") || (idadm == "14"))
                        {

                        }
                        else
                        {
                            //row1.Cells[3].Value = 0.0;
                        }

                        if (i == 6)
                        {
                            if (row1.Cells[5].Value.ToString().Contains("."))
                            {
                                row1.Cells[5].Style.BackColor = System.Drawing.Color.Red;
                                label30.Text = label30.Text + " >> Prestanak radnog odnosa:" + row1.Cells[5].Value.ToString();
                            }
                        }
                    }

                }
            }


        }

        private void btn_vjest_novi_Click(object sender, EventArgs e)
        {
            pnl_v_details.Visible = true;
            btn_v_spremi.Visible = true;
            btn_v_spremiizmjene.Visible = false;

        }

        private void btn_v_spremi_Click(object sender, EventArgs e)
        {


            SqlConnection connection = new SqlConnection(connectionString);
            connection.Open();
            SqlCommand sqlCommand = new SqlCommand("SELECT max(id) id1 from vjestine", connection);
            SqlDataReader reader = sqlCommand.ExecuteReader();
            int newID = 1;
            while (reader.Read())
            {
                if (reader.HasRows)
                {
                    if (reader["id1"] is DBNull)
                    {

                    }
                    else
                    {

                        newID = ((int.Parse)(reader["id1"].ToString())) + 1;
                    }
                }
                else
                {
                    newID = 1;
                }

            }
            connection.Close();

            connection = new SqlConnection(connectionString);
            connection.Open();
            sqlCommand = new SqlCommand("insert into vjestine(id,naziv,grupa) values(" + newID.ToString() + ",'" + txb_v_naziv.Text + "','" + txb_v_grupa.Text + "')", connection);
            reader = sqlCommand.ExecuteReader();
            connection.Close();

            pnl_v_details.Visible = false;

            vjestineToolStripMenuItem_Click(sender, e);


        }

        private void bt_vjest_izmjeni_Click(object sender, EventArgs e)
        {
            pnl_v_details.Visible = true;
            btn_v_spremi.Visible = false;
            btn_v_spremiizmjene.Visible = true;

        }

        private void button3_Click(object sender, EventArgs e)   // Izmjeni vještinu
        {

            if (dgv_vjestine.SelectedCells.Count > 0)
            {
                int selectedrowindex = dgv_vjestine.SelectedCells[0].RowIndex;

                DataGridViewRow selectedRow = dgv_vjestine.Rows[selectedrowindex];

                string id1 = Convert.ToString(selectedRow.Cells["id"].Value);


                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand sqlCommand = new SqlCommand("update vjestine set naziv='" + txb_v_naziv.Text + "',grupa='" + txb_v_grupa.Text + "' where id =" + id1, connection);
                SqlDataReader reader = sqlCommand.ExecuteReader();
                connection.Close();

            }
            else
            {
                MessageBox.Show("Izaberite vještinu koju želite mjenjati !");
            }

            pnl_v_details.Visible = false;

            vjestineToolStripMenuItem_Click(sender, e);
        }

        private void dgv_vjestine_SelectionChanged(object sender, EventArgs e)
        {

        }

        private void btn_vjest_brisi_Click(object sender, EventArgs e)
        {
            pnl_v_dialogbrisi.Visible = true;
        }

        private void btn_v_b_nastavibrisi_Click(object sender, EventArgs e)
        {
            if (dgv_vjestine.SelectedCells.Count > 0)
            {
                int selectedrowindex = dgv_vjestine.SelectedCells[0].RowIndex;

                DataGridViewRow selectedRow = dgv_vjestine.Rows[selectedrowindex];

                string id1 = Convert.ToString(selectedRow.Cells["id"].Value);


                SqlConnection connection = new SqlConnection(connectionString);
                connection.Open();
                SqlCommand sqlCommand = new SqlCommand("delete from vjestine where id =" + id1, connection);
                SqlDataReader reader = sqlCommand.ExecuteReader();
                connection.Close();

            }
            else
            {
                MessageBox.Show("Izaberite vještinu koju želite brisati !");
            }

            pnl_v_details.Visible = false;
            pnl_v_dialogbrisi.Visible = false;
            vjestineToolStripMenuItem_Click(sender, e);
        }

        private void btn_a_javiose_Click(object sender, EventArgs e)
        {
            prikaziizostanak = 1;
            prikazinormu = 0;
            prikazigodisnji = 0;
            prikazibolovanje = 0;
            lbl_izabrano.Text = "Prikaz izostanka";
        }

        private void btn_a_nijesejavio_Click(object sender, EventArgs e)
        {
            prikaziizostanak = 0;
            lbl_izabrano.Text = "Prikaz bolovanja";
            prikazinormu = 0;
            prikazigodisnji = 0;
            prikazibolovanje = 1;
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            prikaziizostanak = 0;
            prikazinormu = 0;
            prikazigodisnji = 1;
            prikazibolovanje = 0;
            lbl_izabrano.Text = "Prikaz iskorištenih godišnjih";
        }

        private void vještineToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OcistiPanele();
            pnl_vjestine_spajanje.Visible = true;
            var dataSource = new List<radnici>();
            dataSource.Add(new radnici() { prezime = " - ", id = "0" });

            foreach (var radnikk in radnicii)
            {
             //   if (radnikk.mt=="700" || radnikk.mt == "716" )
                dataSource.Add(new radnici() { prezime = radnikk.prezime + " " + radnikk.ime + " - " + radnikk.id.ToString() + " ", id = radnikk.id });
            }

            cbx_vj_djelatnici.MaxDropDownItems = 60;
            this.cbx_vj_djelatnici.DataSource = dataSource;
            this.cbx_vj_djelatnici.DisplayMember = "prezime";
            this.cbx_vj_djelatnici.ValueMember = "id";

            cbxl_vjestine.Items.Clear();
            cboxlista.Clear();
            cboxlistaskol.Clear();


            this.combx_funkcije.DataSource = lista_funkcija;
            this.combx_funkcije.DisplayMember = "funkcija";
            this.combx_funkcije.ValueMember = "id";

            this.combx_projekti.DataSource = lista_projects;
            this.combx_projekti.DisplayMember = "naziv";
            this.combx_projekti.ValueMember = "id";

            this.combx_linije.DataSource = lista_linija;
            this.combx_linije.DisplayMember = "naziv";
            this.combx_linije.ValueMember = "id";

            this.combx_mt.DataSource = lista_mt;
            this.combx_mt.DisplayMember = "naziv";
            this.combx_mt.ValueMember = "id";

            string gv = "";
            if (idloged == "5")
            {
                gv = "T";
                tb_hala.Text = "3";
            }

            if (idloged == "1")
                gv = "D";

            if (idloged == "2")
                gv = "A";

            if (idloged == "9")
                gv = "O";           


            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                SqlCommand sqlCommand = new SqlCommand("SELECT * from vjestine order by naziv", cn);

                if (gv != "")
                {
                    sqlCommand = new SqlCommand("SELECT * from vjestine where grupa='" + gv + "' order by naziv ", cn);
                }

                SqlDataReader reader = sqlCommand.ExecuteReader();

                while (reader.Read())
                {
                    // reader["Datum"].ToString();
                    CheckboxItem item = new CheckboxItem();
                    item.Text = reader["naziv"].ToString();
                    item.Value = reader["id"].ToString();
                    cboxlista.Add(item);
                    cboxlistaskol.Add(item);
                }
                cn.Close();
            }
            cbxl_vjestine.Items.AddRange(cboxlista.ToArray());
            cbxlista_skolovanje.Items.AddRange(cboxlistaskol.ToArray());

            string idradnika1="0";

            if (cbx_vj_djelatnici.SelectedIndex.IsNumber())
            {

                if (cbx_vj_djelatnici.SelectedIndex > 0)
                {

                    idradnika1 = cbx_vj_djelatnici.SelectedValue.ToString();
                }
                else
                {
                    return;
                }
            }
        }

        // lista djelatnika
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string idradnika1;

            int ind1 = cbx_vj_djelatnici.SelectedIndex;
            if (ind1 == 0)
                return;

            if (idloged == "5")
            {
                tb_hala.Text = "3";
            }
            int idskolovanja = 0;
            
            string gv = "";
            if (idloged == "5")
            {
                gv = "T";
                tb_hala.Text = "3";
            }

            if (idloged == "1")
                gv = "D";

            if (idloged == "2")
                gv = "A";

            if (idloged == "9")
                gv = "O";

            cbxl_vjestine.Items.Clear();
            cbxlista_skolovanje.Items.Clear();
            cboxlista.Clear();
            cboxlistaskol.Clear();
            lista_skolovanja.Clear();
            
            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                SqlCommand sqlCommand = new SqlCommand("SELECT * from vjestine order by naziv", cn);

                if (gv != "")
                {
                    sqlCommand = new SqlCommand("SELECT * from vjestine where grupa='" + gv + "' order by naziv ", cn);
                }

                SqlDataReader reader = sqlCommand.ExecuteReader();

                while (reader.Read())
                {
                    // reader["Datum"].ToString();
                    CheckboxItem item = new CheckboxItem();
                    item.Text = reader["naziv"].ToString();
                    item.Value = reader["id"].ToString();
                    cboxlista.Add(item);
                    cboxlistaskol.Add(item);
                }
                cn.Close();
            }
       
            cbxlista_skolovanje.Items.AddRange(cboxlistaskol.ToArray());
            cbxl_vjestine.Items.AddRange(cboxlista.ToArray());   // napuni sve vjestine
            
            if (cbx_vj_djelatnici.SelectedIndex.IsNumber())
            {
                if (cbx_vj_djelatnici.SelectedIndex > 0)
                {
                    idradnika1 = cbx_vj_djelatnici.SelectedValue.ToString();
                    pnl_detail1.Visible = true;
                    textBox4.Text = "";
                    tb_s_napomena.Text = "";
                    checkBox14.Checked = false;

                    skolovanje1 sk1 = new skolovanje1();
                    sk1.idskolovanje = "0";
                    sk1.opis = "--- Odaberi program školovanja";

                    lista_skolovanja.Add(sk1);

                    // traži u listi školovanja
                    string sql11 = "select * from skolovanje where idradnika=" + idradnika1+ " order by izmjenaod desc";
                    SqlConnection cn2 = new SqlConnection(connectionString);
                    cn2.Open();
                    SqlCommand sqlcommand1 = new SqlCommand(sql11, cn2);
                    SqlDataReader reader1;
                    reader1 = sqlcommand1.ExecuteReader();
                    while (reader1.Read())
                    {
                        pnl_škola.Visible = true;
                        checkBox14.Checked = true;
                        textBox4.Text = reader1["mentor"].ToString();
                        if (reader1["idskolovanja"]==DBNull.Value)
                        {
                            idskolovanja = 0;
                        }
                        else
                        {
                            idskolovanja = (int.Parse)(reader1["idskolovanja"].ToString());
                        }
                        
                        tb_s_napomena.Text = reader1["napomena"].ToString();

                        sk1 = new skolovanje1();
                        sk1.idskolovanje = idskolovanja.ToString();                                                                
                        sk1.opis = tb_s_napomena.Text;

                        lista_skolovanja.Add(sk1);

                        string dat0 = reader1["oddatuma"].ToString();
                        DateTime dt = Convert.ToDateTime(dat0);
                        //DateTime dat1 = new DateTime(2012, 05, 28);
                        dateTimePicker2.Value = dt;
                        dat0 = reader1["dodatuma"].ToString();
                        dt = Convert.ToDateTime(dat0);
                        dateTimePicker3.Value = dt;

                    }
                    reader1.Close();
                    cn2.Close();
                    cn2.Open();
                    // popuni datum zapošljavanja
                    sql11 = "select * from radnici_ where id=" + idradnika1;
                    cn2 = new SqlConnection(connectionString);
                    cn2.Open();
                    sqlcommand1 = new SqlCommand(sql11, cn2);                    
                    reader1 = sqlcommand1.ExecuteReader();
                    while (reader1.Read())
                    {
                        label142.Text = "Datum zaposlenja: "+ reader1["datumzaposlenja"].ToString();
                    }
                    reader1.Close();
                    cn2.Close();
                }
                else
                {
                    return;
                }
            }
            else
            { return; }

            //// napuni combobox_skolovanja
            //this.combox_skolovanja.DataSource = lista_skolovanja;
            //this.combox_skolovanja.DisplayMember = "opis";
            //this.combox_skolovanja.ValueMember = "idskolovanje";


            cboxlista.Clear();
            cboxlistaskol.Clear();

            // povijest rada po halama, linijama

            string sql1 = "select e.hala,e.linija,p.kratkinazivikupca,count(*) kolikoputa,sum(e.kolicinaok) komada,sum(e.otpadobrada) skarta " +
                              "from feroapp.dbo.evidencijanormiview e " +
                              "left join feroapp.dbo.radnici r on r.id_radnika = e.id_radnika " +
                              "left join feroapp.dbo.partneri p on p.id_par = e.id_par " +
                              "where r.id_fink = " + idradnika1 + " and e.kolicinaok > 0 " +
                              "group by e.hala,e.linija,p.kratkinazivikupca " +
                              "order by count(*) desc";
                                   
                SqlConnection connection = new SqlConnection(connectionString);
                SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
                DataSet ds = new DataSet();
                connection.Open();
                dataadapter.Fill(ds, "event");
                connection.Close();

                // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
                dgv_vjest_history.DataSource = ds;
                dgv_vjest_history.DataMember = "event";
                dgv_vjest_history.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            
            // lista vjestina za pojedinog radnika, označeno da ili ne

            foreach (CheckboxItem item in cbxl_vjestine.Items)
            {

                item.Text = item.Text;
                item.Value = item.Value;
                item.checked1 = false;

                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    SqlCommand sqlCommand = new SqlCommand("SELECT * from radnicivjestine where idradnika=" + idradnika1, cn);
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    while (reader.Read())
                    {
                        // reader["Datum"].ToString();
                        string vje1 = reader["idvjestine"].ToString();
                        if (item.Value.ToString() == vje1)
                        {
                            item.checked1 = true;
                        }
                    }
                }

                if (idloged == "5")
                {
                    if (item.checked1)
                       cboxlista.Add(item);
                }
                else
                       cboxlista.Add(item);

            }

            cbxl_vjestine.Items.Clear();
            cbxl_vjestine.Items.AddRange(cboxlista.ToArray());
            
            for ( int z1 =0;z1<cboxlista.Count;z1++)
            {
                var a1 = cbxl_vjestine.Items[z1];
                CheckboxItem o1 = (CheckboxItem)(a1);
                if (o1.checked1)
                    cbxl_vjestine.SetItemChecked(z1, true);
            }
            
            cboxlistaskol.Clear();
            #region brisi
            //for (int ii = 0; ii < cboxlista.Count; ii++)
            //{


            //    for ( int j = 0; j< cbxl_vjestine.Items.Count;j++)
            //    {
            //        var v1 = cbxl_vjestine.Items[ii];
            //        CheckboxItem cv1 = (CheckboxItem)(v1);
            //        string vj1 = cv1.Value.ToString();
            //        if (cboxlista.Exists(i => i.Value == vj1))
            //        {

            //        }
            //        else
            //        {

            //        }

            //    }





            //    {


            //    }


            //    if (cboxlista.Contains(vj1))
            //        {

            //    }

            //    if (cboxlista[ii].checked1)
            //    {
            //        cbxl_vjestine.SetItemCheckState(ii, CheckState.Checked);
            //    }
            //    else
            //    {
            //        cbxl_vjestine.SetItemCheckState(ii, CheckState.Unchecked);
            //        cbxl_vjestine.Items.RemoveAt(ii);
            //    }
            //}
            // lista vjestina na skolovanju
            #endregion
            cbxlista_skolovanje.Visible = false ;
            cbx_skolo_marked.Visible = false;
            cboxlistaskol.Clear();

            if ((pnl_škola.Visible) &&  (idloged == "5"))
                {
                cbxlista_skolovanje.Visible = true;
                cbx_skolo_marked.Visible = true;
                string vje1 = "",vrije1="";

                foreach (CheckboxItem item in cbxlista_skolovanje.Items)
                {
                    item.Text = item.Text;
                    item.Value = item.Value;
                    item.checked1 = false;

                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        SqlCommand sqlCommand = new SqlCommand("SELECT * from radnicivjestines where idradnika=" + idradnika1, cn);
                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        while (reader.Read())
                        {
                            // reader["Datum"].ToString();
                            vje1 = reader["idvjestine"].ToString();
                            vrije1 = reader["vrijednost"].ToString();
                            if ((item.Value.ToString() == vje1) && (vrije1 == "1"))
                            {
                                item.checked1 = true;
                            }
                            if (item.Value.ToString() == vje1)
                            {
                                cboxlistaskol.Add(item);
                            }
                        }
                    }
                    //if (cbx_skolo_marked.Visible && cbx_skolo_marked.Checked)   // dodaj samo označene
                    //{
                    //    if (item.checked1)
                    //    {
                    //        cboxlistaskol.Add(item);
                    //    }

                    //}
                    //else
                    //{
                    //    cboxlistaskol.Add(item);
                    //}

                }

                cbxlista_skolovanje.Items.Clear();
                cbxlista_skolovanje.Items.AddRange(cboxlistaskol.ToArray());
                var p1 = cbxlista_skolovanje;
                
                for (int ii = 0; ii < cboxlistaskol.Count; ii++)
                {
                    var o1 = cbxlista_skolovanje.Items[ii];
                    CheckboxItem cb1 = (CheckboxItem)(o1);
                    if (cb1.checked1)
                        cbxlista_skolovanje.SetItemCheckState(ii, CheckState.Checked); 
                }                               
            }
            //
            this.combox_skolovanja.DataSource = null;
            this.combox_skolovanja.DataSource = lista_skolovanja;
            this.combox_skolovanja.DisplayMember = "opis";
            this.combox_skolovanja.ValueMember = "idskolovanje";
            
            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                SqlCommand sqlCommand = new SqlCommand("SELECT * from kompetencije where id=" + idradnika1, cn);
                SqlDataReader reader = sqlCommand.ExecuteReader();
                string projekt1 = "", hala1 = "", linija1 = "", funkcija1 = "", napomena1 = "",idprojekt="",mt1="",idmt1="",idfunkcija1="",idlinija1="";

                while (reader.Read())
                {
                    projekt1 = reader["projekt"].ToString().Trim();
                    mt1 = reader["mjesto_troska"].ToString();
                    funkcija1 = reader["funkcija"].ToString();
                    linija1 = reader["linija"].ToString();
                    idprojekt = "";idmt1 = "";idfunkcija1 = ""; idlinija1="";
                    using (SqlConnection cnp = new SqlConnection(connectionString))
                    {
                        cnp.Open();
                        SqlCommand sqlCommand2 = new SqlCommand("SELECT * from projekti where naziv ='" + projekt1 + "'", cnp);
                        SqlDataReader reader2 = sqlCommand2.ExecuteReader();
                        while (reader2.Read())
                        {
                            idprojekt = reader2["id"].ToString();
                        }
                        cnp.Close();
                        cnp.Open();
                        sqlCommand2 = new SqlCommand("SELECT * from funkcije where funkcija ='" + funkcija1 + "'", cnp);
                        reader2 = sqlCommand2.ExecuteReader();
                        while (reader2.Read())
                        {
                            idfunkcija1 = reader2["id"].ToString();
                        }
                        cnp.Close();
                        cnp.Open();
                        sqlCommand2 = new SqlCommand("SELECT * from linije where naziv ='" + linija1 + "'", cnp);
                        reader2 = sqlCommand2.ExecuteReader();
                        while (reader2.Read())
                        {
                            idlinija1 = reader2["id"].ToString();
                        }
                        cnp.Close();
                        cnp.Open();
                        sqlCommand2 = new SqlCommand("SELECT * from mjestotroska where naziv ='" + mt1 + "'", cnp);
                        reader2 = sqlCommand2.ExecuteReader();
                        while (reader2.Read())
                           {
                                idmt1 = reader2["id"].ToString();
                            }                       
                        
                        cnp.Close();
                    }

                    tb_hala.Text = reader["hala"].ToString();                    
                    //tb_funkcija.Text = reader["funkcija"].ToString();
                    tb_skolvposto.Text = reader["Školovanje_posto"].ToString();
                    tb_napomena.Text = reader["napomena"].ToString();                    
                    combx_projekti.SelectedValue = idprojekt;
                    combx_funkcije.SelectedValue = idfunkcija1;
                    combx_linije.SelectedValue   = idlinija1;
                    combx_mt.SelectedValue       = idmt1;
                }
                cn.Close();
            }

        }

        private void btn_vj_spremi_Click(object sender, EventArgs e)
        {
            // ovo radi update samo vještina
            string idradnika1 = cbx_vj_djelatnici.SelectedValue.ToString();
            string firma = "1";
            string idvjestine;
            int noviidskolovanja = 0;

            int idskolo=(int.Parse)(combox_skolovanja.SelectedValue.ToString());

            //if ((!cbs_novo.Checked)  && (idskolo>=0))
            //{
            //    using (SqlConnection cn = new SqlConnection(connectionString))
            //    {
            //        cn.Open();
            //        string sql1 = "delete from RadniciVjestine where idradnika= " + idradnika1 + " and idskolovanja = " + idskolo.ToString();
            //        SqlCommand sqlCommand = new SqlCommand(sql1, cn);
            //        SqlDataReader reader = sqlCommand.ExecuteReader();
            //        cn.Close();
            //    }
            //}

            if (cbs_novo.Checked)
            {
                foreach (CheckboxItem item in cbxl_vjestine.CheckedItems)
                {
                    //if item.  Novo ili izmjena

                    idvjestine = item.Value.ToString();

                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        string sql1 = "insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestine + "," + firma + ")";
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        cn.Close();
                    }

                }
            }
            int uk = cbxl_vjestine.Items.Count;
            for (int ii = 0; ii < uk; ii++)
            {
                cbxl_vjestine.SetItemCheckState(ii, CheckState.Unchecked);
            }
            cbx_vj_djelatnici.SelectedIndex = 0;
            noviidskolovanja = 0;

            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                string sql1 = "select(nvl( max(idskolovanja) ids from radnicivjestines";
                SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                SqlDataReader reader = sqlCommand.ExecuteReader();

                reader.Read();
                noviidskolovanja = (int.Parse)(reader["ids"].ToString()) +1 ;
                               
                cn.Close();
            }



            foreach (CheckboxItem item in cbxlista_skolovanje.CheckedItems)
            {
                //if item.

                idvjestine = item.Value.ToString();

                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    string sql1 = "insert into RadniciVjestineS (idradnika,idvjestine,firma,idskolovanje) values (" + idradnika1 + "," + idvjestine + "," + firma + "," + noviidskolovanja.ToString() + ")";
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                    SqlDataReader reader = sqlCommand.ExecuteReader();

                    cn.Close();
                }

            }
            int uks = cbxlista_skolovanje.Items.Count;
            for (int ii = 0; ii < uk; ii++)
            {
                cbxlista_skolovanje.SetItemCheckState(ii, CheckState.Unchecked);
            }
            cbx_vj_djelatnici.SelectedIndex = 0;


            // test popunjavanje radnicivjestine
            //if item.
            #region popunjavanje vjestina if (1 == 2)
            {
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    string sql10 = "select * from kompetencije0310$";
                    SqlCommand sqlCommand0 = new SqlCommand(sql10, cn);
                    SqlDataReader reader0 = sqlCommand0.ExecuteReader();

                    //2   AR
                    //11  AR, IR, Turm
                    //12  Porshe
                    //14  IR
                    //15  VC Hurco
                    //16  Sherer
                    //17  Okuma
                    //18  PUMA tvrdo tokarenje
                    //19  IV
                    //20  EMAG
                    //21  Valjci
                    //22  Pile
                    //23  Wupertal prsten
                    //24  IWK
                    //25  Pakiranje IWK
                    //26  SKF
                    //27  Pakiranje SKF
                    //28  Glodanje SKF
                    //29  Više vretanasta bušilica
                    //30  Stupna bušilica
                    //31  Narezivanje navoja
                    //32  SONA
                    //33  3494
                    //34  GT600
                    //35  Kontrola prstenova 100 posto
                    //36  Kontrola poroznosti
                    //37  Kontrola SONA 100 posto
                    //38  Membrane
                    //AR
                    //IR
                    //Turm
                    //VC_HURCO
                    //Porche
                    //Sherer
                    //Okuma
                    //Puma_TT
                    //IV
                    //Emag
                    //Valjci
                    //Pile
                    //[WUP_ prsten]
                    //IWK
                    //PakiranjeIWK
                    //Membrane
                    //SKF
                    //PakiranjeSKF
                    //GlodanjeSKF
                    //ViseVretena
                    //Stupna
                    //[Narezivanje_ Navoja]
                    //Sona
                    //[3494]
                    //GT600
                    //Kontrola100PostoPrsteni
                    //Kontrola_Poroznosti
                    //Kontrola_100PostoSona

                    string sql1 = "", idvjestina = "";
                    while (reader0.Read())
                    {
                        sql1 = "";
                        idradnika1 = reader0["ID"].ToString();
                        string vjestina2 = reader0["AR"].ToString();
                        string vjestina11 = reader0["IR"].ToString();
                        string vjestinaTurm = reader0["Turm"].ToString();
                        string vjestina15 = reader0["VC_HURCO"].ToString();
                        string vjestina12 = reader0["Porche"].ToString();
                        string vjestina16 = reader0["Sherer"].ToString();
                        string vjestina17 = reader0["Okuma"].ToString();
                        string vjestina18 = reader0["Puma_TT"].ToString();
                        string vjestina19 = reader0["IV"].ToString();
                        string vjestina20 = reader0["Emag"].ToString();
                        string vjestina21 = reader0["Valjci"].ToString();
                        string vjestina22 = reader0["Pile"].ToString();
                        string vjestina23 = "0";// reader0["WUPprsten]"].ToString();
                        string vjestina24 = reader0["IWK"].ToString();
                        string vjestina25 = reader0["PakiranjeIWK"].ToString();
                        string vjestina38 = reader0["Membrane"].ToString();
                        string vjestina26 = reader0["SKF"].ToString();
                        string vjestina27 = reader0["PakiranjeSKF"].ToString();
                        string vjestina28 = reader0["GlodanjeSKF"].ToString();
                        string vjestina29 = reader0["ViseVretena"].ToString();
                        string vjestina30 = reader0["Stupna"].ToString();
                        string vjestina31 = reader0["Narezivanje_ Navoja"].ToString();
                        string vjestina32 = reader0["Sona"].ToString();
                        string vjestina33 = reader0["3494"].ToString();
                        string vjestina34 = reader0["GT600"].ToString();
                        string vjestina35 = reader0["Kontrola100PostoPrsteni"].ToString();
                        string vjestina36 = reader0["Kontrola_Poroznosti"].ToString();
                        string vjestina37 = reader0["Kontrola_100PostoSona"].ToString();
                        firma = reader0["Poduzece"].ToString();

                        if (firma == "FX")
                            firma = "1";
                        else
                            firma = "3";

                        if (idradnika1 == "651")
                        {
                            idradnika1 = idradnika1;
                        }


                        if ((vjestina2 == "1") && (vjestina11 == "1") && (vjestinaTurm == "1"))  // AR,IR, Turm
                        {
                            idvjestina = "11";
                            sql1 = "insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina2 == "1") && (vjestina11 == "0") && (vjestinaTurm == ""))   // AR
                        {
                            idvjestina = "2";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina11 == "1") && (vjestina2 == "0") && (vjestinaTurm == ""))   // IR
                        {
                            idvjestina = "14";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina11 == "1") && (vjestina2 == "1") && (vjestinaTurm == ""))   // AR,IR
                        {
                            idvjestina = "2";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                            idvjestina = "14";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }


                        if ((vjestina15 == "1"))   // VC_HURCO
                        {
                            idvjestina = "15";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina12 == "1"))   // Porche
                        {
                            idvjestina = "12";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina16 == "1"))   // Sherer
                        {
                            idvjestina = "16";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina17 == "1"))   // Okuma
                        {
                            idvjestina = "17";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina18 == "1"))   // Puma_TT
                        {
                            idvjestina = "18";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina19 == "1"))   // IV
                        {
                            idvjestina = "19";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina20 == "1"))   // Emag
                        {
                            idvjestina = "20";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina21 == "1"))   // Valjci
                        {
                            idvjestina = "21";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina22 == "1"))   // Pile
                        {
                            idvjestina = "22";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina23 == "1"))   // Wupertal prsteni
                        {
                            idvjestina = "23";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina24 == "1"))   // IWK
                        {
                            idvjestina = "24";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina25 == "1"))   // Pakiranje IWK
                        {
                            idvjestina = "25";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina38 == "1"))   // Membrane
                        {
                            idvjestina = "26";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina26 == "1"))   // SKF
                        {
                            idvjestina = "26";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina27 == "1"))   // Pakiranje SKF
                        {
                            idvjestina = "27";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina28 == "1"))   // GlodanjeSKF
                        {
                            idvjestina = "28";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina29 == "1"))   // Visevretena
                        {
                            idvjestina = "29";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina30 == "1"))   // Stupna
                        {
                            idvjestina = "30";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina31 == "1"))   // Narezivanje_navoja
                        {
                            idvjestina = "31";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina32 == "1"))   // Sona
                        {
                            idvjestina = "32";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina33 == "1"))   // 3494
                        {
                            idvjestina = "33";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina34 == "1"))   //  GT600
                        {
                            idvjestina = "34";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina35 == "1"))   // Kontrola100Prsteni
                        {
                            idvjestina = "35";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina36 == "1"))   //  Kontrola_Poroznosti
                        {
                            idvjestina = "36";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina37 == "1"))   //  Kontrola_100PostoSona
                        {
                            idvjestina = "37";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }

                        if (sql1.Length > 0)
                        {
                            using (SqlConnection cn1 = new SqlConnection(connectionString))
                            {
                                cn1.Open();
                                //sql1 = "insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                                SqlCommand sqlCommand = new SqlCommand(sql1, cn1);
                                SqlDataReader reader = sqlCommand.ExecuteReader();
                                cn1.Close();
                            }
                        }

                    }

                    cn.Close();
                }


                // end test
            }
            #endregion

            
        }

        private void cbxl_vjestine_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        // kad se odabere opcija izvještaj određeno
        private void popisDjelatnikaNaOdređenomToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pnl_odredeno.Visible = true;


            using (SqlConnection cnn1 = new SqlConnection(connectionString))
            {

                var dataSource = new List<mjesto_troska>();

                dataSource.Add(new mjesto_troska() { naziv = " - ", id = "0" });

                foreach (var mt1 in lista_mt)
                {
                    dataSource.Add(new mjesto_troska() { naziv = mt1.naziv, id = mt1.id });
                }

                cbx_mjestotroska.MaxDropDownItems = 60;
                this.combo_odred.DataSource = dataSource;
                this.combo_odred.DisplayMember = "Naziv";
                this.combo_odred.ValueMember = "id";

            }
        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btn_odred_Click(object sender, EventArgs e)
        {

            int firmaid = comb_poduzece.SelectedIndex;
            string mtid = combo_odred.SelectedValue.ToString();

            DateTime dat1 = datep_od_ov.Value;
            DateTime dat2 = datp_do_odr.Value;

            string dat10 = dat1.Year.ToString() + '-' + dat1.Month.ToString() + '-' + dat1.Day.ToString();
            string dat20 = dat2.Year.ToString() + '-' + dat2.Month.ToString() + '-' + dat2.Day.ToString();

            string sql1 = "select p.acregno ID,p.acworker PrezimeIme,p.addatebirth DatumRođenja,j.acCostDrv MT,j.adEmployedTo OdredenoDo from tHR_Prsn p left join thr_prsnjob j on p.acworker = j.acworker where j.adEmployedTo >= '" + dat10 + "' and j.adEmployedTo <= '" + dat20 + "'";
            sql1 = "select p.acregno ID,p.acworker PrezimeIme,p.addatebirth DatumRođenja,j.acCostDrv MT,j.adEmployedTo OdredenoDo from tHR_Prsn p left join thr_prsnjob j on p.acworker = j.acworker where j.acemploymenttype like '%- odre%' and ( (j.adEmployedTo>=" + dat10 + " and j.adEmployedTo<='" + dat20 + "' ) and j.addateend is null) ";
            if (mtid != "0")
            {
                sql1 = "select p.acregno ID,p.acworker PrezimIme,p.addatebirth DatumRođenja,j.acCostDrv MT,j.adEmployedTo OdredenoDo from tHR_Prsn p left join thr_prsnjob j on p.acworker = j.acworker where j.adEmployedTo >= '" + dat10 + "' and j.adEmployedTo <= '" + dat20 + "' and j.accostdrv=" + mtid;
                sql1 = "select p.acregno ID,p.acworker PrezimeIme,p.addatebirth DatumRođenja,j.acCostDrv MT,j.adEmployedTo OdredenoDo from tHR_Prsn p left join thr_prsnjob j on p.acworker = j.acworker where j.acemploymenttype like '%- odre%' and ( (j.adEmployedTo>=" + dat10 + " and j.adEmployedTo<='" + dat20 + "' ) and j.addateend is null) and j.accostdrv=" + mtid;
            }

            string connecStrin = connectionStringpa;

            if (firmaid == 1)
            {
                connecStrin = connectionStringpb;
            }


            SqlConnection connection = new SqlConnection(connecStrin);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            dgv_odred.DataSource = ds;
            dgv_odred.DataMember = "event";
            dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

        }
        // export to excell

        private void button4_Click(object sender, EventArgs e)
        {
            //Creating DataTable
            DataTable dt = new DataTable();

            //Adding the Columns
            foreach (DataGridViewColumn column in dgv_odred.Columns)
            {

                // satnica se pokazuje samo za grupu 8 ili tehn. direktora (14)
                dt.Columns.Add(column.HeaderText, column.ValueType);


            }

            //Adding the Rows
            foreach (DataGridViewRow row in dgv_odred.Rows)
            {
                dt.Rows.Add();

                foreach (DataGridViewCell cell in row.Cells)
                {


                    if (!(cell.Value is DBNull))
                    {
                        if (cell.Value != null)

                        {
                            dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();

                        }
                    }
                }
            }

            string folderPath = "C:\\KKS\\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "IzvjestajOdredjeno");
                wb.SaveAs(folderPath + "IzvjestajOdredjeno.xlsx");
            }
            btn_zbirni_export.Text = "Export done";

            FileInfo fi = new FileInfo("C:\\kks\\IzvjestajOdredjeno.xlsx");
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(@"C:\kks\IzvjestajOdredjeno.xlsx");
            }
            else
            {
                //file doesn't exist
            }

            //Object oMissing = System.Reflection.Missing.Value;
            //oMissing = System.Reflection.Missing.Value;
            //Object oTrue = true;
            //Object oFalse = false;
            //Microsoft.Office.Interop.Word.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Document oExcelDoc = new Microsoft.Office.Interop.Excel.Document();
            //oExcel.Visible = true;

            //Object oTemplatePath = "c:\\kks\\kompetencije.xlsx";
            //oExcelDoc = oExcelDoc.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);
            


        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            OcistiPanele();
        }

        private void dnevniIzvjestajiAktivnostiToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void btn_exportNorm_Click(object sender, EventArgs e)
        {

            //Creating DataTable
            DataTable dt = new DataTable();

            //Adding the Columns
            foreach (DataGridViewColumn column in dgv_norme.Columns)
            {
                dt.Columns.Add(column.HeaderText, column.ValueType);
            }

            //Adding the Rows
            foreach (DataGridViewRow row in dgv_norme.Rows)
            {
                dt.Rows.Add();
                foreach (DataGridViewCell cell in row.Cells)
                {

                    if (cell.Value != null)
                    {
                        if (cell.Value is DBNull)
                        {
                            dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = new TimeSpan(0, 0, 0);
                        }
                        else
                        {
                            dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();
                        }
                    }
                }
            }

            //Exporting to Excel
            string folderPath = "C:\\KKS\\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "Norme");
                wb.SaveAs(folderPath + "Norma1.xlsx");
            }
            button2.Text = "Export done";


            FileInfo fi = new FileInfo("C:\\kks\\Norma1.xlsx");
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(@"C:\\kks\\Norma1.xlsx");
            }
            else
            {
                //file doesn't exist
            }

        }

        private void izračunNormeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (idadm != "8")
            {
                return;
            }



            int m1 = DateTime.Now.Month - 1;

            int g1 = DateTime.Now.Year;
            DateTime datl = new DateTime(g1, m1 + 1, 1).AddDays(-1);
            if (m1 < 0)
            {
                m1 = 12;
                g1 = DateTime.Now.Year - 1;
            }

            DateTime dat1 = new DateTime(g1, m1, 1);
            DateTime dat2 = new DateTime(g1, m1, datl.Day);

            string d11 = dat1.Year.ToString() + '-' + dat1.Month.ToString() + '-' + dat1.Day.ToString();
            string d22 = dat2.Year.ToString() + '-' + dat2.Month.ToString() + '-' + dat2.Day.ToString();

            string sql11 = "delete from mjesecniizv where mjesec="+m1.ToString()+" and godina="+g1.ToString();
            SqlConnection cn2 = new SqlConnection(connectionString);
            cn2.Open();
            SqlCommand sqlCommand1 = new SqlCommand(sql11, cn2);
            SqlDataReader reader1 = sqlCommand1.ExecuteReader();
            reader1.Close();
            cn2.Close();


            string sql1 = "select * from rfind.dbo.evidnormiradad('" + d11 + "','" + d22 + "') order by radnik"; //" + smj; bez računanja efektivnog radnog vremena  

            // sql1 = "rfind.dbo.ldp_recalc '" + dat13 + "','" + dat23 + "',41";
            SqlConnection cn = new SqlConnection(connectionString);
            cn.Open();
            SqlCommand sqlCommand = new SqlCommand(sql1, cn);
            
            SqlDataReader reader = sqlCommand.ExecuteReader();
            string pec;
            int i = 0;
            int ukupn = 0, ukupp = 0, ukupkol = 0, zastoj = 0, norma = 0, k1, n1;
            int ukuppodbacaj = 0, ukupprebacaja = 0,ukupskart=0;

            while (reader.Read())
            {

                string radnikid = reader["id_radnika"].ToString();
                string imeradnika = reader["radnik"].ToString();
                string radnik1 = radnikid;
                int id_fink = 0;
                if (radnik1 != "")
                {
                    sql11 = "select isnull(max(id_fink),0) id_fink from feroapp.dbo.radnici where id_radnika=" + radnik1;
                    cn2 = new SqlConnection(connectionString);
                    cn2.Open();
                    sqlCommand1 = new SqlCommand(sql11, cn2);
                    //reader1 = sqlCommand1.ExecuteReader();
                    id_fink = (Int32)sqlCommand1.ExecuteScalar();
                    reader1.Close();
                    cn2.Close();
                }
                ukupprebacaja = 0;
                ukuppodbacaj = 0;
                ukupskart = 0;
                int jos = 1;
                if (radnik1=="1691")
                {
                    jos=1;
                }
                while (radnikid == radnik1 && jos==1 )
                {
                    if (radnik1 == "1691")
                    {
                        jos = 1;
                    }


                    string sk1 = reader["kolicinaok"].ToString();
                    if (reader["kolicinaok"] == DBNull.Value)
                    {
                        k1 = 0;
                    }
                    else
                    {
                        k1 = (int.Parse)(reader["kolicinaok"].ToString());
                    }

                    if (reader["norma"] == DBNull.Value)
                    {
                        n1 = 0;
                    }
                    else
                    {
                        n1 = (int.Parse)(reader["norma"].ToString());
                    }
                    double post1 = 0.0;
                    if (k1 > 0)
                    {
                        post1 = (k1 + 0.00001) / (n1 + 0.0001);
                    }
                    else
                    {
                        post1 = 0.0;
                    }

                    double satirada = (double.Parse)((reader["minutaradaradnika"].ToString()).Replace(',', '.')) / 100;
                    if (reader["otpadobrada"]== DBNull.Value)
                    { }
                    else
                    {
                        ukupskart = ukupskart + (int.Parse)(reader["otpadobrada"].ToString());
                    }
                    

                    string brisi = reader["norma"].ToString();
                    if (brisi.Length > 0)
                        norma = (int.Parse)(reader["norma"].ToString());
                    else
                        norma = 0;
                    
                    double planirano = norma * (satirada / 480);
                    
                    double posto1 = 0.00;
                    n1 = (int)(planirano);

                    if (n1 > 0)
                    {
                        if (k1 == 0)
                        {
                            posto1 = 0.00;
                        }
                        else
                        {
                            posto1 = posto1 + k1 / (n1 );
                        }

                        if (posto1 >= 1)
                        {
                            ukupprebacaja = ukupprebacaja + 1;

                        }
                        if (posto1 < 1)
                        {
                            ukuppodbacaj = ukuppodbacaj + 1;

                        }

                    }

                    ukupn = ukupn + n1;

                    if (reader["norma"] != DBNull.Value)
                    {
                        ukupp = ukupp + n1;
                        ukupkol = ukupkol + k1;
                    }

                    i++;
                    
                    if (reader.Read())
                    {
                        radnikid = reader["id_radnika"].ToString();
                    }
                    else
                    {

                        radnikid = "-1";
                        jos = 0;
                    }

                }

                if (radnik1 == "1691")
                {
                    jos = 1;
                }

                sql11 = "select count(*) broj,radnomjesto from pregledvremena where idradnika="+id_fink.ToString()+ " and month(datum)=" + m1.ToString() + " and year(datum)=" + g1.ToString() +  " group by radnomjesto" ;
                cn2 = new SqlConnection(connectionString);
                cn2.Open();
                sqlCommand1 = new SqlCommand(sql11, cn2);
                reader1 = sqlCommand1.ExecuteReader();
                int godisnji = 0,bolovanje=0;
                while (reader1.Read())
                {
                    string rm = reader1["radnomjesto"].ToString().Trim();
                    if (rm=="GO" || rm=="G.O.")
                    {
                        godisnji = (int.Parse)(reader1["broj"].ToString());
                    }
                    if (rm == "BO" || rm == "B.O.")
                    {
                        bolovanje = (int.Parse)(reader1["broj"].ToString());
                    }
                }

                sql11 = "rfind.dbo.ldp_recalc '"+m1.ToString()+"."+g1.ToString()+"','"+id_fink.ToString()+"',100";
                cn2 = new SqlConnection(connectionString);
                cn2.Open();
                sqlCommand1 = new SqlCommand(sql11, cn2);
                reader1 = sqlCommand1.ExecuteReader();
                int godisnji2 = 0, bolovanje2 = 0, neopravdano=0;
                while (reader1.Read())
                {


                    string dansve = reader1["Dan01"].ToString()+","+ reader1["Dan02"].ToString()+ ","+reader1["Dan03"].ToString()+","+ reader1["Dan04"].ToString() + "," + reader1["Dan05"].ToString() + "," + reader1["Dan06"].ToString() + "," + reader1["Dan07"].ToString() + "," + reader1["Dan08"].ToString() + "," + reader1["Dan09"].ToString() + "," + reader1["Dan10"].ToString();
                    dansve = dansve+","+reader1["Dan11"].ToString() + "," + reader1["Dan12"].ToString() + "," + reader1["Dan13"].ToString() + "," + reader1["Dan14"].ToString() + "," + reader1["Dan15"].ToString() + "," + reader1["Dan16"].ToString() + "," + reader1["Dan17"].ToString() + "," + reader1["Dan18"].ToString() + "," + reader1["Dan19"].ToString() + "," + reader1["Dan20"].ToString();
                    dansve = dansve + "," + reader1["Dan21"].ToString() + "," + reader1["Dan22"].ToString() + "," + reader1["Dan23"].ToString() + "," + reader1["Dan24"].ToString() + "," + reader1["Dan25"].ToString() + "," + reader1["Dan26"].ToString() + "," + reader1["Dan27"].ToString() + "," + reader1["Dan28"].ToString() + "," + reader1["Dan29"].ToString() + "," + reader1["Dan30"].ToString()+"," + reader1["Dan31"].ToString();

                    string regexPattern = @"\b7g\b";
                    godisnji2 = Regex.Matches(dansve, regexPattern).Count;

                    regexPattern = @"\b5g\b";
                    godisnji2 = godisnji2 + Regex.Matches(dansve, regexPattern).Count;

                    regexPattern = @"\b8b\b";
                    bolovanje2 = Regex.Matches(dansve, regexPattern).Count;

                    regexPattern = @"\b0e\b";
                    neopravdano = Regex.Matches(dansve, regexPattern).Count;

                    
                }



                reader1.Close();
                cn2.Close();
                if (radnik1 != "")
                {
                    sql11 = "insert into mjesecniizv (mjesec,godina,idradnika,imeradnika,prebacaja,podbacaja,otpadobrada,godisnji,bolovanje,neopravdano) values ( " + m1.ToString() + "," + g1.ToString() + ",'" + id_fink.ToString() + "','" + imeradnika + "'," + ukupprebacaja.ToString() + "," + ukuppodbacaj.ToString() + "," + ukupskart.ToString() + "," + godisnji2.ToString() + "," + bolovanje2.ToString()  +"," + neopravdano.ToString() + ")";
                    cn2 = new SqlConnection(connectionString);
                    cn2.Open();
                    sqlCommand1 = new SqlCommand(sql11, cn2);
                    reader1 = sqlCommand1.ExecuteReader();
                    reader1.Close();
                    cn2.Close();
                }


            }
            cn.Close();
            MessageBox.Show( "Gotovo !");
        }

        private void pregledIzvršenjaPoRadnicimaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Ocisti();
            OcistiPanele();

            pnl_dnevizv_norme.Visible = true;
            dgv_dnev_izv_norme.Visible = true;
            dgv_dnev_izv_norme.Height = 800;
            pnl_dnevizv_norme.Height = 1000;

            dgv_dnev_izv_norme.Width = 1100;
            pnl_dnevizv_norme.Width = 1200;

            string sql = "select * from mjesecniizv order by imeradnika";
            SqlConnection connection = new SqlConnection(connectionString);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql, connection);
            DataSet ds2 = new DataSet();
            connection.Open();
            dataadapter.Fill(ds2, "event2");
            connection.Close();

            dgv_dnev_izv_norme.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgv_dnev_izv_norme.DataSource = ds2;
            dgv_dnev_izv_norme.DataMember = "event2";

            //FreezeBand(dgv_aktivnost.Columns[1]);
            dgv_dnev_izv_norme.Columns[2].Frozen = true;
            dgv_dnev_izv_norme.AutoResizeColumns();
            dgv_dnev_izv_norme.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgv_dnev_izv_norme.ReadOnly = true;

        }

        private void menuStrip1_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void tb_vjestina_TextChanged(object sender, EventArgs e)
        {

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void combx_projekti_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (combx_projekti.Text.Contains("Š_"))
            {
                label127.Visible = true;
                tb_skolvposto.Visible = true;
            }
            else
            {
                label127.Visible = false;
                tb_skolvposto.Visible = false;
            }
        }

        private void checkBox14_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox14.Checked)
            {
                pnl_škola.Visible = true;
                textBox4.Text = "";

                if (cbx_vj_djelatnici.SelectedIndex.IsNumber())
                {

                    if (cbx_vj_djelatnici.SelectedIndex > 0)
                    {

                        string idradnika1 = cbx_vj_djelatnici.SelectedValue.ToString();
                        // traži u listi školovanja
                        string sql11 = "select * from skolovanje where idradnika=" + idradnika1 + " order by izmjenaod" ;
                        SqlConnection cn2 = new SqlConnection(connectionString);
                        cn2.Open();
                        SqlCommand sqlcommand1 = new SqlCommand(sql11, cn2);
                        SqlDataReader reader1;
                        reader1 = sqlcommand1.ExecuteReader();
                        while (reader1.Read())
                                {

                            textBox4.Text = reader1["mentor"].ToString();
                            tb_s_napomena.Text= reader1["napomena"].ToString();
                            string dat0 = reader1["oddatuma"].ToString();
                            DateTime dt = Convert.ToDateTime(dat0);
                            //DateTime dat1 = new DateTime(2012, 05, 28);
                            dateTimePicker2.Value = dt;
                            dat0 = reader1["dodatuma"].ToString();
                            dt = Convert.ToDateTime(dat0);
                            dateTimePicker3.Value = dt;

                        }
                        reader1.Close();
                        cn2.Close();





                    }
                    else
                    {
                        return;
                    }


                }
                else
                { return; }


            }
            else
                pnl_škola.Visible = false;
        }

        private void cbx_odlazak_CheckedChanged(object sender, EventArgs e)
        {
            if (cbx_odlazak.Checked)
            {
                lbl_najava.Visible = true;
                datetp1.Visible = true;
            }
                
            else
            {
                lbl_najava.Visible = false;
                datetp1.Visible = false;
            }
        }

        // spremi sve

        private void btn_spremiv_Click(object sender, EventArgs e)
        {
            pnl_škola.Visible = false;
            // izmjena školovanje ????, mjenaj sve
            
            /////////////////// begin loading
            ///
            if ((textBox4.Text.Contains("TTTT")) && ( 1==2))
                {

                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    string sql1 = "select * from kks_to$ order by idradnika,idvjestine";
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    reader.Read();

                    string idradnika0= reader["idradnika"].ToString();
                    string idvjestine0=reader["idvjestine"].ToString();

                    string poduzece0 = "", ime0 = "", prezime0 = "",firma0="1";
                    int idskolovanja0 = 0;
                    string danas = DateTime.Now.Year.ToString() + "-" + DateTime.Now.Month.ToString() + "-" + DateTime.Now.Day.ToString();

                    using (SqlConnection cns = new SqlConnection(connectionString))
                    {
                        cns.Open();
                        string sqls = "select max(idskolovanja) ids from skolovanje" ;
                        SqlCommand sqlCommand2 = new SqlCommand(sqls, cns);
                        SqlDataReader readers = sqlCommand2.ExecuteReader();
                        idskolovanja0=5;
                        if (readers.HasRows)
                            //idskolovanja0 = (int.Parse)(readerr["ids"].ToString())+1;   // novi id skolovanja
                        cns.Close();
                    }

                    using (SqlConnection cnr = new SqlConnection(connectionString))
                    {
                        cnr.Open();
                        string sqlr = "select * from radnici_ where id=" + idradnika0;
                        SqlCommand sqlCommand2r = new SqlCommand(sqlr, cnr);
                        SqlDataReader readerr = sqlCommand2r.ExecuteReader();
                        readerr.Read();
                        poduzece0 = readerr["poduzece"].ToString();
                        firma0 = "3";

                        if (poduzece0.Contains("Feroim"))
                        {
                            firma0 = "1";
                            poduzece0 = "FX";
                        }
                        ime0 =  readerr["prezime"].ToString().TrimEnd()+ " "+readerr["ime"].ToString().TrimEnd() + " - " + readerr["id"].ToString().TrimEnd() ;
                        cnr.Close();
                    }
                    
                    // unesi novo školovanje
                    string sqlrvs = "";
                    using (SqlConnection cns = new SqlConnection(connectionString))
                    {
                        cns.Open();
                        sqlrvs = idskolovanja0 + "," + idradnika0 + ",'" + ime0 + "','2019-03-15','2019-04-01','','3','','Dalibor Jančić','Školovanje_','" + poduzece0 + "','" + danas + "',1)";
                        sqlrvs = " insert into skolovanje(idskolovanja,idradnika,imeradnika,oddatuma,dodatuma,projekt,hala,linija,mentor,napomena,poduzece,izmjenaod,gotovo) values ( " + sqlrvs;
                        SqlCommand sqlCommand2s = new SqlCommand(sqlrvs, cns);
                        SqlDataReader reader2s = sqlCommand2s.ExecuteReader();
                        cns.Close();
                    }
                                       
                    int jos = 1;
                    
                    while (jos==1)
                    {
                        idvjestine0 = reader["idvjestine"].ToString();  // uzmi vjestinu iz kks_to$

                        using (SqlConnection cn2 = new SqlConnection(connectionString))
                        {
                            cn2.Open();
                            sqlrvs = idradnika0 + "," + idvjestine0 + ",1," + firma0 + "," + idskolovanja0.ToString() + ",'" + danas + "',7)";
                            sqlrvs= " insert into radnicivjestines(idradnika,idvjestine,vrijednost,firma,idskolovanja,datumocjenjivanja,ocjenjivac) values("+sqlrvs;
                            SqlCommand sqlCommand2rs = new SqlCommand(sqlrvs, cn2);
                            SqlDataReader reader2rs = sqlCommand2rs.ExecuteReader();
                            cn2.Close();
                        }

                        using (SqlConnection cn2 = new SqlConnection(connectionString))
                        {
                            cn2.Open();
                            sqlrvs = idradnika0 + "," + idvjestine0 + ",1," + firma0+" )" ;
                            sqlrvs = " insert into radnicivjestine(idradnika,idvjestine,vrijednost,firma) values(" + sqlrvs;
                            SqlCommand sqlCommand2rs2 = new SqlCommand(sqlrvs, cn2);
                            SqlDataReader reader2rs2 = sqlCommand2rs2.ExecuteReader();
                            cn2.Close();
                        }

                        reader.Read();
                        if (reader["idradnika"]==DBNull.Value)
                        {
                            jos = 0;
                            continue;
                        }


                            if (idradnika0 == reader["idradnika"].ToString())
                        {
                            continue;
                        }
                        else
                        {

                            idskolovanja0 = idskolovanja0 + 1;   // novi id skolovanja
                            idradnika0 = reader["idradnika"].ToString();
                            
                            using (SqlConnection cnr = new SqlConnection(connectionString))
                            {
                                cnr.Open();

                                string sqlr = "select * from radnici_ where id=" + idradnika0;
                                SqlCommand sqlCommand21 = new SqlCommand(sqlr, cnr);
                                SqlDataReader readerr21 = sqlCommand21.ExecuteReader();
                                readerr21.Read();
                                poduzece0 = readerr21["poduzece"].ToString();
                                firma0 = "3";
                                if (poduzece0.Contains("Feroim"))
                                {
                                    firma0 = "1";
                                    poduzece0 = "FX";
                                }
                                ime0 = readerr21["ime"].ToString().TrimEnd() + " " + readerr21["prezime"].ToString().TrimEnd()+" - "+ readerr21["id"].ToString().TrimEnd();

                                sqlrvs = "";
                                using (SqlConnection cns = new SqlConnection(connectionString))
                                {
                                    cns.Open();
                                    sqlrvs = idskolovanja0 + "," + idradnika0 + ",'" + ime0 + "','2019-03-15','2019-04-01','','3','','Dalibor Jančić','Školovanje_','" + poduzece0 + "','" + danas + "',1)";
                                    sqlrvs = " insert into skolovanje(idskolovanja,idradnika,imeradnika,oddatuma,dodatuma,projekt,hala,linija,mentor,napomena,poduzece,izmjenaod,gotovo) values ( " + sqlrvs;
                                    SqlCommand sqlCommand22 = new SqlCommand(sqlrvs, cns);
                                    SqlDataReader reader22s = sqlCommand22.ExecuteReader();
                                    cns.Close();
                                }

                            }
                        }                                                                     

                    }                                                                          
                    cn.Close();
                }

                return;
                
            }
                       
         /////////////////// end loading
            string idradnika1 = cbx_vj_djelatnici.SelectedValue.ToString();
            string imeradnika1 = cbx_vj_djelatnici.Text.ToString();
            string poduzece1 = "";
            if (idradnika1 == "" || idradnika1=="0")
                    return;

            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                string sql1 = "select poduzece from kompetencije where id=" + idradnika1;
                SqlCommand sqlCommand = new SqlCommand(sql1, cn);                
                SqlDataReader reader = sqlCommand.ExecuteReader();
                reader.Read();
                poduzece1=reader["Poduzece"].ToString();
                cn.Close();
            }

            //if (cb_pod1.Checked)
            //    poduzece1 = "FX";
            //if (cb_pod2.Checked)
            //    poduzece1 = "Tokabu";

            //string projekt1 = combx_projekti.Text.ToString();
            //string mt = combx_mt.Text.ToString();


            string firma = "1";    //?????
            string idvjestine;
            int i1 = combox_skolovanja.SelectedIndex;
            if (!cbs_novo.Checked)
            {
                if (i1<=0)
                   {
                      MessageBox.Show("Odaberite školovanje !");
                      return;
                   }
            }
            
            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                string sql1 = "delete from RadniciVjestine where idradnika=" + idradnika1;
                SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                SqlDataReader reader = sqlCommand.ExecuteReader();
                cn.Close();
            }

            foreach (CheckboxItem item in cbxl_vjestine.CheckedItems)
            {
                //if item.
                string vrijednost1 = "0" ;
                idvjestine = item.Value.ToString();
              //  if (!cbs_novo.Checked)
                {
                    vrijednost1 = "0";
                    if (item.checked1)
                        vrijednost1 = "1";
                }
                if (idloged=="5")
                {

                }
                //if (vrijednost1 == "1")
                {
                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        string sql1 = "insert into RadniciVjestine (idradnika,idvjestine,vrijednost,firma) values (" + idradnika1 + "," + idvjestine + "," + vrijednost1 + "," + firma + ")";
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        cn.Close();
                    }
                }
            }            

            // test popunjavanje radnicivjestine
            //if item.
            #region popunjavanje vjestina if (1 == 2)
            if (1==2)
            {
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    string sql10 = "select * from kompetencije0310$";
                    SqlCommand sqlCommand0 = new SqlCommand(sql10, cn);
                    SqlDataReader reader0 = sqlCommand0.ExecuteReader();

                    //2   AR
                    //11  AR, IR, Turm
                    //12  Porshe
                    //14  IR
                    //15  VC Hurco
                    //16  Sherer
                    //17  Okuma
                    //18  PUMA tvrdo tokarenje
                    //19  IV
                    //20  EMAG
                    //21  Valjci
                    //22  Pile
                    //23  Wupertal prsten
                    //24  IWK
                    //25  Pakiranje IWK
                    //26  SKF
                    //27  Pakiranje SKF
                    //28  Glodanje SKF
                    //29  Više vretanasta bušilica
                    //30  Stupna bušilica
                    //31  Narezivanje navoja
                    //32  SONA
                    //33  3494
                    //34  GT600
                    //35  Kontrola prstenova 100 posto
                    //36  Kontrola poroznosti
                    //37  Kontrola SONA 100 posto
                    //38  Membrane
                    //AR
                    //IR
                    //Turm
                    //VC_HURCO
                    //Porche
                    //Sherer
                    //Okuma
                    //Puma_TT
                    //IV
                    //Emag
                    //Valjci
                    //Pile
                    //[WUP_ prsten]
                    //IWK
                    //PakiranjeIWK
                    //Membrane
                    //SKF
                    //PakiranjeSKF
                    //GlodanjeSKF
                    //ViseVretena
                    //Stupna
                    //[Narezivanje_ Navoja]
                    //Sona
                    //[3494]
                    //GT600
                    //Kontrola100PostoPrsteni
                    //Kontrola_Poroznosti
                    //Kontrola_100PostoSona

                    string sql1 = "", idvjestina = "";
                    while (reader0.Read())
                    {
                        sql1 = "";
                        idradnika1 = reader0["ID"].ToString();
                        string vjestina2 = reader0["AR"].ToString();
                        string vjestina11 = reader0["IR"].ToString();
                        string vjestinaTurm = reader0["Turm"].ToString();
                        string vjestina15 = reader0["VC_HURCO"].ToString();
                        string vjestina12 = reader0["Porche"].ToString();
                        string vjestina16 = reader0["Sherer"].ToString();
                        string vjestina17 = reader0["Okuma"].ToString();
                        string vjestina18 = reader0["Puma_TT"].ToString();
                        string vjestina19 = reader0["IV"].ToString();
                        string vjestina20 = reader0["Emag"].ToString();
                        string vjestina21 = reader0["Valjci"].ToString();
                        string vjestina22 = reader0["Pile"].ToString();
                        string vjestina23 = "0";// reader0["WUPprsten]"].ToString();
                        string vjestina24 = reader0["IWK"].ToString();
                        string vjestina25 = reader0["PakiranjeIWK"].ToString();
                        string vjestina38 = reader0["Membrane"].ToString();
                        string vjestina26 = reader0["SKF"].ToString();
                        string vjestina27 = reader0["PakiranjeSKF"].ToString();
                        string vjestina28 = reader0["GlodanjeSKF"].ToString();
                        string vjestina29 = reader0["ViseVretena"].ToString();
                        string vjestina30 = reader0["Stupna"].ToString();
                        string vjestina31 = reader0["Narezivanje_Navoja"].ToString();
                        string vjestina32 = reader0["Sona"].ToString();
                        string vjestina33 = reader0["3494"].ToString();
                        string vjestina34 = reader0["GT600"].ToString();
                        string vjestina35 = reader0["Kontrola100PostoPrsteni"].ToString();
                        string vjestina36 = reader0["Kontrola_Poroznosti"].ToString();
                        string vjestina37 = reader0["Kontrola_100PostoSona"].ToString();
                        firma = reader0["Poduzece"].ToString();

                        if (firma == "FX")
                            firma = "1";
                        else
                            firma = "3";

                        if (idradnika1 == "651")
                        {
                            idradnika1 = idradnika1;
                        }


                        if ((vjestina2 == "1") && (vjestina11 == "1") && (vjestinaTurm == "1"))  // AR,IR, Turm
                        {
                            idvjestina = "11";
                            sql1 = "insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina2 == "1") && (vjestina11 == "0") && (vjestinaTurm == ""))   // AR
                        {
                            idvjestina = "2";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina11 == "1") && (vjestina2 == "0") && (vjestinaTurm == ""))   // IR
                        {
                            idvjestina = "14";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina11 == "1") && (vjestina2 == "1") && (vjestinaTurm == ""))   // AR,IR
                        {
                            idvjestina = "2";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                            idvjestina = "14";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }


                        if ((vjestina15 == "1"))   // VC_HURCO
                        {
                            idvjestina = "15";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina12 == "1"))   // Porche
                        {
                            idvjestina = "12";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina16 == "1"))   // Sherer
                        {
                            idvjestina = "16";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina17 == "1"))   // Okuma
                        {
                            idvjestina = "17";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina18 == "1"))   // Puma_TT
                        {
                            idvjestina = "18";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina19 == "1"))   // IV
                        {
                            idvjestina = "19";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina20 == "1"))   // Emag
                        {
                            idvjestina = "20";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina21 == "1"))   // Valjci
                        {
                            idvjestina = "21";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina22 == "1"))   // Pile
                        {
                            idvjestina = "22";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina23 == "1"))   // Wupertal prsteni
                        {
                            idvjestina = "23";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina24 == "1"))   // IWK
                        {
                            idvjestina = "24";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina25 == "1"))   // Pakiranje IWK
                        {
                            idvjestina = "25";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina38 == "1"))   // Membrane
                        {
                            idvjestina = "26";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina26 == "1"))   // SKF
                        {
                            idvjestina = "26";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina27 == "1"))   // Pakiranje SKF
                        {
                            idvjestina = "27";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina28 == "1"))   // GlodanjeSKF
                        {
                            idvjestina = "28";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina29 == "1"))   // Visevretena
                        {
                            idvjestina = "29";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina30 == "1"))   // Stupna
                        {
                            idvjestina = "30";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina31 == "1"))   // Narezivanje_navoja
                        {
                            idvjestina = "31";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina32 == "1"))   // Sona
                        {
                            idvjestina = "32";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina33 == "1"))   // 3494
                        {
                            idvjestina = "33";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina34 == "1"))   //  GT600
                        {
                            idvjestina = "34";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina35 == "1"))   // Kontrola100Prsteni
                        {
                            idvjestina = "35";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina36 == "1"))   //  Kontrola_Poroznosti
                        {
                            idvjestina = "36";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }
                        if ((vjestina37 == "1"))   //  Kontrola_100PostoSona
                        {
                            idvjestina = "37";
                            sql1 = sql1 + ";insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                        }

                        if (sql1.Length > 0)
                        {
                            using (SqlConnection cn1 = new SqlConnection(connectionString))
                            {
                                cn1.Open();
                                //sql1 = "insert into RadniciVjestine (idradnika,idvjestine,firma) values (" + idradnika1 + "," + idvjestina + "," + firma + ")";
                                SqlCommand sqlCommand = new SqlCommand(sql1, cn1);
                                SqlDataReader reader = sqlCommand.ExecuteReader();
                                cn1.Close();
                            }
                        }

                    }

                    cn.Close();
                }


                // end test
            }
            #endregion
            
            // 
            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();// napravi kopiju starog sloga prije promjene
                string vrijeme = DateTime.Now.Year.ToString();
                string sql10 = "insert into kompetencije_log( id , prezimeime , napomena , projekt,hala,linija,funkcija,mjesto_troska,korisnik,datumpromjene)  select id,prezimeime,napomena,projekt,hala,linija,funkcija,mjesto_troska,'"+korisnik+"' as korisnik,getdate() from kompetencije where id=" + idradnika1;
                SqlCommand sqlCommand0 = new SqlCommand(sql10, cn);
                SqlDataReader reader0 = sqlCommand0.ExecuteReader();
                cn.Close();
            }
            // update kompetencije
            string projekt1 = combx_projekti.Text.ToString();
            string linija1 = combx_linije.Text.ToString(); ;

            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                string idradnik1 = idradnika1;
                string napo1 = tb_napomena.Text.ToString();
                string hala1 = tb_hala.Text;
                string funkc1 = combx_funkcije.Text.ToString();
                string skolovposto = tb_skolvposto.Text.ToString();
                string mt1 = combx_mt.Text.ToString();
                string sql10 = "update kompetencije set napomena='"+napo1+"',projekt='"+projekt1+"', Hala='"+tb_hala.Text.ToString()+ "',Školovanje_posto='"+skolovposto+"', linija='" + linija1+ "', funkcija='" + funkc1+ "', mjesto_troska='" + mt1+"' where id=" + idradnik1;
                SqlCommand sqlCommand0 = new SqlCommand(sql10, cn);
                SqlDataReader reader0 = sqlCommand0.ExecuteReader();
                cn.Close();
            }
            int noviidskolovanja = 0;
            string sql3 = "";
            
            // školovanje
            using (SqlConnection cnn3 = new SqlConnection(connectionString))
            {
                cnn3.Open();

                DateTime sod = dateTimePicker2.Value;
                DateTime sdo = dateTimePicker3.Value;
                string d1 = sod.Year + "-" + sod.Month + "-" + sod.Day;
                string d2 = sdo.Year + "-" + sdo.Month + "-" + sdo.Day;
                string hala2 = textBox6.Text.ToString();
                string linija2 = textBox7.Text.ToString();
                string mentor2 = textBox4.Text.ToString();
                string projekt2 = textBox5.Text.ToString();
                string napomena2 = tb_s_napomena.Text.ToString();

                int gotovo = 0;
                if (checkBox13.Checked)
                {
                    gotovo = 1;
                }
                string poduzece2 =  poduzece1;

                sql3 = "select * from skolovanje where gotovo=0 and idradnika="+idradnika1;
                SqlCommand cmd3 = new SqlCommand(sql3, cnn3);
                SqlDataReader rdr3 = null;
                rdr3 = cmd3.ExecuteReader();
                int imaskolovanje = 0,idskolovanje1=-1;
                while ( rdr3.Read() )
                {
                    imaskolovanje = 1;
                    idskolovanje1 = (int.Parse)(rdr3["idskolovanja"].ToString());
                }
                var o10 = combox_skolovanja.SelectedItem;
                skolovanje1 sk10 = (skolovanje1)(o10);

                idskolovanje1 = (int.Parse)(sk10.idskolovanje.ToString());  // izabrani id skolovanja, za update
                //int i10 = combox_skolovanja.SelectedIndex;                
                //var 
                //idskolovanje1 = (int.Parse)(combox_skolovanja.SelectedValue.ToString());
                                
                cnn3.Close();
                sql3 = "";

                if (cbs_novo.Checked)  // ako je novo odredi novi id skolovanja
                {
                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        string sql1 = "select max(idskolovanja) ids from radnicivjestines";
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        reader.Read();

                        if (reader["ids"]==DBNull.Value)
                            noviidskolovanja = 1;
                        else
                            noviidskolovanja = (int.Parse)(reader["ids"].ToString()) + 1;
                    
                        cn.Close();

                    }
                    sql3 = "insert into skolovanje(idskolovanja,idradnika,imeradnika,oddatuma,dodatuma,projekt,hala,linija,mentor,napomena,gotovo,poduzece,izmjenaod) values(" + noviidskolovanja+","+idradnika1 + ",'" + imeradnika1 + "','" + d1 + "','" + d2 + "','" + projekt1 + "','" + tb_hala.Text + "','" + linija1 + "','" + mentor2 + "','" + napomena2 + "'," + gotovo.ToString() + ",'" + poduzece2 + "',getdate() )";

                 }

                //if ( (cbs_izmjena.Checked) && ( ( imaskolovanje==1)&&((!cbs_novo.Checked)) ) )        // update

            if (((imaskolovanje == 1) && ((!cbs_novo.Checked))))        // update
            {
                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        string sql1 = "delete from RadniciVjestineS where idskolovanja=" + idskolovanje1.ToString() + " and idradnika="+idradnika1;
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        cn.Close();
                        noviidskolovanja = idskolovanje1;
                    }

                    sql3 = "update skolovanje set oddatuma='"+d1+"',dodatuma='"+d2+"',projekt='"+projekt1+"',hala='"+tb_hala.Text+"',linija='"+linija1+"',mentor='"+mentor2+"',napomena='"+napomena2+"',gotovo="+gotovo.ToString()+",poduzece='"+poduzece2 + "',izmjenaod=getdate() where idradnika=" + idradnika1+ " and idskolovanja="+idskolovanje1.ToString() ;

                 }
                 cnn3.Open();
                 cmd3 = new SqlCommand(sql3, cnn3);
                 rdr3 = null;
                 rdr3 = cmd3.ExecuteReader();
                 cnn3.Close();

            }

            if ((sql3 != "") && (idloged == "5"))   // samo za TO
            {
                int uks = 0;

                uks = cbxlista_skolovanje.Items.Count;
                var ch1 = cbxlista_skolovanje.Items;

                string check1 = "0";
                idvjestine = "";

                for (int ii = 0; ii < uks; ii++)
                {

                    var o1 = ((CheckboxItem)ch1[ii]);
                    if (cbxlista_skolovanje.GetItemCheckState(ii) == CheckState.Checked)
                    {
                        check1 = "1";
                    }
                    else
                    {
                        if (cbs_novo.Checked)   // ako je novo i nije označen , preskoči
                        {
                            continue;
                        }
                        else   // ako onije novo
                        {
                            check1 = "0";
                        }

                    }

                    idvjestine = o1.Value.ToString();
                    if (cbs_novo.Checked)   // ako je novo , stavi check1=0
                    {
                        check1 = "0";
                    }

                    //CheckboxItem checked10 = ((CheckboxItem)ch1[ii]);
                    //check1 = "0";

                    //if (checked10.checked1)
                    //        check1 = "1";

                    DateTime date1 = DateTime.Now;
                    string danas = date1.Year.ToString() + "-" + date1.Month.ToString() + "-" + date1.Day.ToString();
                    int r= 0;
                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        string sql1 = "insert into RadniciVjestineS (idradnika,vrijednost,idvjestine,firma,idskolovanja,datumocjenjivanja,ocjenjivac) values (" + idradnika1 + "," + check1 + "," + idvjestine + "," + firma + "," + noviidskolovanja.ToString() + ",'"+ danas + "'," + idadm +")";
                        SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                        SqlDataReader reader = sqlCommand.ExecuteReader();
                        cn.Close();
                        r = 1;
                        
                    }

                    if ((check1 == "1") && (cbs_izmjena.Checked == true)) // ako je vještina na školovanju označena, ubaci je u radne vještine
                    {
                        using (SqlConnection cn = new SqlConnection(connectionString))
                        {
                            cn.Open();
                            string sql1 = "insert into RadniciVjestine(idradnika,vrijednost,idvjestine,firma) values (" + idradnika1 + "," + check1 + "," + idvjestine + "," + firma + ")";
                            SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                            SqlDataReader reader = sqlCommand.ExecuteReader();
                            cn.Close();
                        }

                    }
                    if ((r==0) && (check1=="1"))
                    {
                        MessageBox.Show("test1","ok");
                    }



                }

                for (int ii = 0; ii < uks; ii++)
                {
                    cbxlista_skolovanje.SetItemCheckState(ii, CheckState.Unchecked);
                }
            }
            cbx_vj_djelatnici.SelectedIndex = 0;

            MessageBox.Show("Podaci su spremljeni !");
            //vještineToolStripMenuItem_Click.Click();
            pnl_detail1.Visible = false;
            cbs_novo.Checked = false;
            cbs_izmjena.Checked = false;
            int uk = cbxl_vjestine.Items.Count;
            for (int ii = 0; ii < uk; ii++)
            {
                cbxl_vjestine.SetItemCheckState(ii, CheckState.Unchecked);
            }
            cbx_vj_djelatnici.SelectedIndex = 0;

        }

        private void cbs_novo_CheckedChanged(object sender, EventArgs e)
        {

            if (cbs_novo.Checked)
            {

                string gv = "";
                if (idloged == "5")
                {
                    gv = "T";
                    tb_hala.Text = "3";
                }

                if (idloged == "1")
                    gv = "D";

                if (idloged == "2")
                    gv = "A";

                if (idloged == "9")
                    gv = "O";

                cbs_izmjena.Checked = false;
                cbxlista_skolovanje.Items.Clear();
                cboxlistaskol.Clear();

                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    SqlCommand sqlCommand = new SqlCommand("SELECT * from vjestine order by naziv", cn);

                    if (gv != "")
                    {
                        sqlCommand = new SqlCommand("SELECT * from vjestine where grupa='" + gv + "' order by naziv ", cn);
                    }

                    SqlDataReader reader = sqlCommand.ExecuteReader();

                    while (reader.Read())
                    {
                        // reader["Datum"].ToString();
                        CheckboxItem item = new CheckboxItem();
                        item.Text = reader["naziv"].ToString();
                        item.Value = reader["id"].ToString();
                        cboxlista.Add(item);
                        cboxlistaskol.Add(item);
                    }
                    cn.Close();
                }
                cbxl_vjestine.Items.AddRange(cboxlista.ToArray());
                cbxlista_skolovanje.Items.AddRange(cboxlistaskol.ToArray());
                cbxlista_skolovanje.Visible = true;
                cbx_skolo_marked.Visible = true;


            }

        }

        private void cbs_izmjena_CheckedChanged(object sender, EventArgs e)
        {
            cbs_novo.Checked = false;
        }

        private void cb_pod1_CheckedChanged(object sender, EventArgs e)
        {
            cb_pod2.Checked= false;
        }

        private void cb_pod2_CheckedChanged(object sender, EventArgs e)
        {
            cb_pod1.Checked = false;
        }


        // vještine
        private void button5_Click(object sender, EventArgs e)
        {
            string listavjestina1 = "", listavjestina2 = "", listavjestina = "";
            //foreach (CheckboxItem item in cbxl_vjestine.Items)
            //{
            //    //if item  is checked
            //    listavjestina1 = listavjestina1 + item.Value.ToString() + ",";
            //}

            foreach (CheckboxItem item in cbxl_vjestine.CheckedItems)
            {
                //if item  is checked
                listavjestina = listavjestina + item.Value.ToString() + ",";
            }
            if (listavjestina != "")   // obrisi zarez na kraju
                listavjestina = listavjestina.Substring(0, listavjestina.Length - 1);

            string sql0 = "";

            if (idloged == "1" || idloged == "8" || idloged == "14")
            {

                sql0 = "select x1.* " +
                            "from( select distinct k.id, k.prezimeime, k.funkcija, k.projekt, k.hala, k.linija, k.mjesto_troska,mt.grupa1 gmt," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 2) ar," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 14) ir," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 39) turm," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 15) vc_hurco," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 12) porche," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 16) sherer," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 17) okuma," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 18) puma_tt," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 19) iv," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 20) emag," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 21) valjci," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 22) pile," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 23) wupertal_prsten," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 24) iwk," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 25) pakiranje_iwk," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 38) membrane," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 26) skf," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 27) pakiranje_skf," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 28) glodanje_skf," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 29) vise_vretenasta," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 30) stupna," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 31) narezivanje_navoja," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 32) sona," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 33) [3494]," +
                            "(select count(*) from radnicivjestine v where v.idradnika=k.id and idvjestine= 34) gt600," +
                            "(select count(*) from radnicivjestine v where v.idradnika=k.id and idvjestine= 35) kontrolaprstena100posto," +
                            "(select count(*) from radnicivjestine v where v.idradnika=k.id and idvjestine= 36) kontrolaporoznosti," +
                            "(select count(*) from radnicivjestine v where v.idradnika=k.id and idvjestine= 37) kontrolasona100posto,";
                //        string sql10="k.DatumZaposlenja,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt) x1 where id in (select idradnika from radnicivjestine where idvjestine in (" + listavjestina + ")) and x1.gmt='" + idloged + "' order by prezimeime";
            }
            else if (idloged=="5")
            {
                sql0 = "select x1.* " +
                          "from( select distinct k.id, k.prezimeime, k.funkcija, k.projekt, k.hala, k.linija, k.mjesto_troska,mt.grupa1 gmt," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 61) BDF_AR_IR," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 62) CementacijaDebrecen," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 63) CementacijaUsluga," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 64) KaljenjeUsluga," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 65) Sona," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 66) NSK," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 67) PakiranjeBDF," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 68) VozačViličara," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 69) DnevniPlanKaljenja," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 70) DnevniPlanPakiranja," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 71) PlanSlaganja," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 72) SlaganjeŠarže," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 73) PredPranje," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 74) PlanŠaržeProces," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 75) DaljnjaObrada," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 76) RaskapanjeŠarže," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 77) ProgListProizvKalj ," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 78) ProgListProizvPaki, " +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 79) UmjeravanjeTvrdomjera, " +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 80) MjerTvrdoćeRockwell, " +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 81) ProgTvrdoće, " +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 82) KvalitŠaržMaterOkretSelekti, " +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 83) RadNaUređajuPakiranje, " +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 84) OcjenaKvaliteteKaljMikros," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 85) KontrolaParametaraPeći," +                          
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 86) PriprUzorZaMetalogr," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 87) ŠkolovNovogDjelat," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 88) RučnoVođenjeProcesaUklanjanjeZastoja," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 89) SamostalnoVođenjeSmjeneBezTestovaINovihPproizvoda," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 90) SamostalnoRiješavanjeZastojaIManjihKvarova," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 91) SamostalanUProvođenjuPlanaKaljenja," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 92) SamostalanUDokazivanjuKvaliteteTO," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 93) SamostalanUNadzoruIIzmjenamaParametara," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 94) SamostalanDimenzionalnaKontrolaSetUPMjernogStola," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 95) RasporedLjudstva," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 96) IzradaPlanovaŠaržiranjaIReceptura," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 97) MehaničkaPoštelavanja," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 98) RiješavanjeGrešakaNaUređaju," +
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 99) PlaniranjeProizvodnjePremaPlanuKupca,"+
                          "(select isnull(vrijednost,0) from radnicivjestine v where v.idradnika = k.id and idvjestine = 100) IzradaDokumentacijeERSTEMUSTERA,";

                sql0 = " select select r.prezime+' 'r.ime prezimeime,j.Naziv,k.funkcija,k.projekt,k.hala,k.linija,k.mjesto_troska,k.RadnoMjesto,k.DatumZaposlenja from radnicivjestine v " +
                       "left join vjestine j on j.id = v.idvjestine left join radnici_ r on r.id = v.idradnika left join kompetencije k on k.id = r.id where v.idvjestine in (" + listavjestina + ") order by prezime";

                sql0 = "select distinct k.id, k.prezimeime, k.funkcija, k.projekt, k.hala, k.linija, k.mjesto_troska,mt.grupa1 gmt," ;
            }
            else
            {
                sql0 = "select x1.* " +
                        "from( select distinct k.id, k.prezimeime, k.funkcija, k.projekt, k.hala, k.linija, k.mjesto_troska,mt.grupa1 gmt,";


            }

            string sql1 = "select distinct k.* from kompetencije k left join radnicivjestine v on k.id=v.idradnika  order by prezimeime";
            string hala1 = tb_hala.Text.ToString();
            if (hala1 == "")
                hala1 = "1,3";

            if (idloged == "8")   // ako je grupa 8 ili tehn. direktor, može vidjeti sve, može printati
            {           
                
                if (listavjestina == "")
                {
                    sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 order by prezimeime";
                }
                else
                {
                    sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where id in (select idradnika from radnicivjestine where idvjestine in (" + listavjestina + "))  order by prezimeime";
                }

            }
            else if (idloged=="5")
            {
                if (listavjestina == "")
                {

                    sql1 = " select r.prezime+' '+r.ime prezimeime,j.Naziv,k.funkcija,k.projekt,k.hala,k.linija,k.mjesto_troska,k.RadnoMjesto,k.DatumZaposlenja from radnicivjestine v " +
                           "left join vjestine j on j.id = v.idvjestine left join radnici_ r on r.id = v.idradnika left join kompetencije k on k.id = r.id  order by prezime";
                }
                else
                    sql1 = " select r.prezime+' '+r.ime prezimeime,j.Naziv,k.funkcija,k.projekt,k.hala,k.linija,k.mjesto_troska,k.RadnoMjesto,k.DatumZaposlenja from radnicivjestine v " +
                       "left join vjestine j on j.id = v.idvjestine left join radnici_ r on r.id = v.idradnika left join kompetencije k on k.id = r.id where v.idvjestine in (" + listavjestina + ") order by prezime";
                }
            else
            {
                if (idadm == "14")  // tehnički direktor može vidjeti određene grupe  1- tokarija,steleri, 2 održavanje,,alatnica,brusenje 5 kaljenje 13 tehnologija
                {

                    if (listavjestina == "")
                    {
                        sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where x1.gmt in ( 1,2,5,13) order by prezimeime";
                    }
                    else
                    {
                        sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where id in (select idradnika from radnicivjestine where idvjestine in (" + listavjestina + ") and x1.gmt in ( 1,2,5,13) order by prezimeime";
                    }
                }
                else   // ostali korisnici, kicin etc.
                {

                    if (listavjestina == "")
                    {
                        sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where x1.hala in ("+hala1+") and x1.gmt='" + idloged + "' order by prezimeime";
                    }
                    else
                    {
                        sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where id in (select idradnika from radnicivjestine where idvjestine in (" + listavjestina + ")) and x1.hala in ("+hala1+") and x1.gmt='" + idloged + "' order by prezimeime";
                    }
                }               


            }


            // sql1 = "select distinct k.* from kompetencije k left join radnicivjestine v on k.id=v.idradnika  order by prezimeime";

            //            if (listavjestina!="")
            //          {
            //    listavjestina = listavjestina.Substring(0, listavjestina.Length - 1);
            //   sql1 = "select distinct k.* from kompetencije k left join radnicivjestine v on k.id=v.idradnika  where v.idvjestine in (" + listavjestina + ") order by prezimeime";
            //          }

            lblvj_stat.Text = "Broj redova: 0" ;
            SqlConnection connection = new SqlConnection(connectionString);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            // Pregled vjestina za djelatnike
            dgv_vjest_history.DataSource = ds;
            dgv_vjest_history.DataMember = "event";
            dgv_vjest_history.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgv_vjest_history.ReadOnly = false;            
            this.dgv_vjest_history.Columns["prezimeime"].Frozen = true;
            for (int ii1=0;ii1<8;ii1++)  // prvih 7 kolona su readonly
            {
                dgv_vjest_history.Columns[ii1].ReadOnly = true;
            }
            

            sql1 = null; sql0=null;
            lblvj_stat.Text="Broj redova:"+(dgv_vjest_history.RowCount-1).ToString();

        }

        //export vjestine
        private void button6_Click(object sender, EventArgs e)
        {
            //Creating DataTable
            DataTable dt = new DataTable();

            //Adding the Columns
            foreach (DataGridViewColumn column in dgv_vjest_history.Columns)
            {
                dt.Columns.Add(column.HeaderText, column.ValueType);
            }

            //Adding the Rows
            foreach (DataGridViewRow row in dgv_vjest_history.Rows)
            {
                dt.Rows.Add();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if ( (cell.Value != DBNull.Value) &&  (cell.Value!=null))
                    {
                        dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();
                    }
                }
            }

            //Exporting to Excel
            string folderPath = "C:\\KKS\\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "Djelatnici");
                wb.SaveAs(folderPath + "VjestineAktivnost.xlsx");
            }
            button6.Text = "Export done";


            FileInfo fi = new FileInfo("C:\\kks\\VjestineAktivnost.xlsx");
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(@"C:\\kks\\VjestineAktivnost.xlsx");
            }
            else
            {
                //file doesn't exist
            }

        }
        // filter na školovanju
        private void button7_Click(object sender, EventArgs e)
        {
            string listavjestina1 = "", listavjestina2 = "", listavjestina = "";
            //foreach (CheckboxItem item in cbxl_vjestine.Items)
            //{
            //    //if item  is checked
            //    listavjestina1 = listavjestina1 + item.Value.ToString() + ",";
            //}
            string idradnika1s = "";
            if (cbx_vj_djelatnici.SelectedIndex.IsNumber())
            {

                if (cbx_vj_djelatnici.SelectedIndex > 0)
                {

                 idradnika1s = cbx_vj_djelatnici.SelectedValue.ToString();
                }
                else
                {
                    return;
                }
            }
            
            DataSet ds = new DataSet();
            string sql1 = "", sql0="";
            if (idloged == "5")
            {
                    //SqlCommand cmd = new SqlCommand("dbo.KKS_skolovanje1 "+, cnn1);
                    //cmd.CommandType = CommandType.StoredProcedure;
                    lblvj_stat.Text = "Broj redova: 0";
                    sql1 = "dbo.kks_skolovanje1 "+idradnika1s;
                    SqlConnection connection = new SqlConnection(connectionString);
                    SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);

                    connection.Open();
                    dataadapter.Fill(ds, "event");
                    connection.Close();
                    SqlDataReader rdr = null;
                    SqlConnection cnn = null;
                    //rdr = cmd.ExecuteReader();               
            }
            else
            {
                sql0 = "select x1.* " +
                            "from( select distinct k.id, mt.grupa1 gmt, k.prezimeime, k.funkcija, k.projekt,k.školovanje_posto, k.hala, k.linija, k.mjesto_troska," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 2) ar," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 14) ir," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 39) turm," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 15) vc_hurco," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 12) porche," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 16) sherer," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 17) okuma," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 18) puma_tt," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 19) iv," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 20) emag," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 21) valjci," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 22) pile," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 23) wupertal_prsten," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 24) iwk," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 25) pakiranje_iwk," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 38) membrane," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 26) skf," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 27) pakiranje_skf," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 28) glodanje_skf," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 29) vise_vretenasta," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 30) stupna," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 31) narezivanje_navoja," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 32) sona," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 33) [3494]," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine= 34) gt600," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine= 35) kontrolaprstena100posto," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine= 36) kontrolaporoznosti," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine= 37) kontrolasona100posto,";
                //        string sql10="k.DatumZaposlenja,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt) x1 where id in (select idradnika from radnicivjestine where idvjestine in (" + listavjestina + ")) and x1.gmt='" + idloged + "' order by prezimeime";

                sql1 = "select distinct k.* from kompetencije k left join radnicivjestine v on k.id=v.idradnika  order by prezimeime";
                string hala1 = tb_hala.Text.ToString();
                if (hala1 == "")
                    hala1 = "1,3";

                if (idloged == "8")   // ako je grupa 8 ili tehn. direktor, može vidjeti sve, može printati
                {
                    if (listavjestina == "")
                    {
                        sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where x1.projekt like 'š_%' order by prezimeime";
                    }
                }
                else
                {
                    if (idadm == "14")  // tehnički direktor može vidjeti određene grupe  1- tokarija,steleri, 2 održavanje,,alatnica,brusenje 5 kaljenje 13 tehnologija
                    {

                        if (listavjestina == "")
                        {
                            sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where x1.gmt in ( 1,2,5,13) order by prezimeime";
                        }
                        else
                        {
                            sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where id in (select idradnika from radnicivjestine where idvjestine in (" + listavjestina + ") and x1.gmt in ( 1,2,5,13) order by prezimeime";
                        }
                    }
                    else   // ostali korisnici, kicin etc.
                    {
                        sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where  ((x1.projekt like 'Š_%') or (x1.projekt like 'š_%')) and x1.hala in (" + hala1 + ") and x1.gmt='" + idloged + "' order by prezimeime";
                    }


                }

                lblvj_stat.Text = "Broj redova: 0";
                SqlConnection connection = new SqlConnection(connectionString);
                SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
             
                connection.Open();
                dataadapter.Fill(ds, "event");
                connection.Close();

            }

            // sql1 = "select distinct k.* from kompetencije k left join radnicivjestine v on k.id=v.idradnika  order by prezimeime";

            //            if (listavjestina!="")
            //          {
            //    listavjestina = listavjestina.Substring(0, listavjestina.Length - 1);
            //   sql1 = "select distinct k.* from kompetencije k left join radnicivjestine v on k.id=v.idradnika  where v.idvjestine in (" + listavjestina + ") order by prezimeime";
            //          }

            

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            dgv_vjest_history.DataSource = null;
            dgv_vjest_history.DataSource = ds;
            dgv_vjest_history.DataMember = "event";
            dgv_vjest_history.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgv_vjest_history.ReadOnly = true;
            sql1 = null; sql0 = null;
            lblvj_stat.Text = "Broj redova:" + (dgv_vjest_history.RowCount - 1).ToString();
            
        }


        // filter početnici
        private void button8_Click(object sender, EventArgs e)
        {
            // filter na školovanju
                string listavjestina1 = "", listavjestina2 = "", listavjestina = "";
                //foreach (CheckboxItem item in cbxl_vjestine.Items)
                //{
                //    //if item  is checked
                //    listavjestina1 = listavjestina1 + item.Value.ToString() + ",";
                //}


                string sql0 = "select x1.* " +
                            "from( select distinct k.id, mt.grupa1 gmt, k.prezimeime, k.funkcija, k.projekt, k.hala, k.linija, k.mjesto_troska," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 2) ar," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 14) ir," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 39) turm," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 15) vc_hurco," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 12) porche," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 16) sherer," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 17) okuma," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 18) puma_tt," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 19) iv," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 20) emag," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 21) valjci," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 22) pile," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 23) wupertal_prsten," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 24) iwk," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 25) pakiranje_iwk," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 38) membrane," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 26) skf," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 27) pakiranje_skf," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 28) glodanje_skf," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 29) vise_vretenasta," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 30) stupna," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 31) narezivanje_navoja," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 32) sona," +
                            "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 33) [3494]," +
                            "(select count(*) from radnicivjestine v where v.idradnika=k.id and idvjestine= 34) gt600," +
                            "(select count(*) from radnicivjestine v where v.idradnika=k.id and idvjestine= 35) kontrolaprstena100posto," +
                            "(select count(*) from radnicivjestine v where v.idradnika=k.id and idvjestine= 36) kontrolaporoznosti," +
                            "(select count(*) from radnicivjestine v where v.idradnika=k.id and idvjestine= 37) kontrolasona100posto,";
                //        string sql10="k.DatumZaposlenja,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt) x1 where id in (select idradnika from radnicivjestine where idvjestine in (" + listavjestina + ")) and x1.gmt='" + idloged + "' order by prezimeime";

                string sql1 = "select distinct k.* from kompetencije k left join radnicivjestine v on k.id=v.idradnika  order by prezimeime";
                DateTime dat1 = DateTime.Now.AddDays(-30);
                string datz = dat1.Year.ToString() + '-' + dat1.Month.ToString() + '-' + dat1.Day.ToString();
                string hala1 = tb_hala.Text.ToString();
            if ((hala1 == "") || (hala1 is null))
                hala1 = "'1','3','',NULL";

            if (idloged == "8")   // ako je grupa 8 ili tehn. direktor, može vidjeti sve, može printati
                {

                    if (listavjestina == "")
                    {
                        sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where ( x1.projekt='' or x1.projekt is null) and x1.datumzaposlenja>='" + datz+"' order by prezimeime";
                    }

                }
                else
                {
                    if (idloged== "7")  // tehnički direktor može vidjeti određene grupe  1- tokarija,steleri, 2 održavanje,,alatnica,brusenje 5 kaljenje 13 tehnologija
                    {


                    if (listavjestina == "")
                        {
                            sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where x1.gmt in ( 1,2,5,13) order by prezimeime";
                        }
                        else
                        {
                            sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where id in (select idradnika from radnicivjestine where idvjestine in (" + listavjestina + ") and x1.gmt in ( 1,2,5,13) order by prezimeime";
                        }

                    sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where  (x1.projekt='' or x1.projekt is null)  and ( x1.hala in (" + hala1 + ") or x1.hala is null) and x1.datumzaposlenja>='" + datz + "' and x1.gmt in (1,2,5,3,13) order by prezimeime";

                }
                    else   // ostali korisnici, kicin etc.
                    {
                        sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where  (x1.projekt='' or x1.projekt is null)  and ( x1.hala in ("+hala1+") or x1.hala is null) and x1.datumzaposlenja>='" + datz + "' and x1.gmt='" + idloged + "'order by prezimeime"; 
                    }


                }


                // sql1 = "select distinct k.* from kompetencije k left join radnicivjestine v on k.id=v.idradnika  order by prezimeime";

                //            if (listavjestina!="")
                //          {
                //    listavjestina = listavjestina.Substring(0, listavjestina.Length - 1);
                //   sql1 = "select distinct k.* from kompetencije k left join radnicivjestine v on k.id=v.idradnika  where v.idvjestine in (" + listavjestina + ") order by prezimeime";
                //          }

                lblvj_stat.Text = "Broj redova: 0" ;
                SqlConnection connection = new SqlConnection(connectionString);
                SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
                DataSet ds = new DataSet();
                connection.Open();
                dataadapter.Fill(ds, "event");
                connection.Close();

                // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
                dgv_vjest_history.DataSource = ds;
                dgv_vjest_history.DataMember = "event";
                dgv_vjest_history.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dgv_vjest_history.ReadOnly = true;
                sql1 = null; sql0 = null;
                lblvj_stat.Text = "Broj redova:" + (dgv_vjest_history.RowCount - 1).ToString();
        }

        // filter projekti
        private void button9_Click(object sender, EventArgs e)
        {
            string projekt1 = combx_projekti.Text;
            // filter na školovanju
            string listavjestina1 = "", listavjestina2 = "", listavjestina = "";
            //foreach (CheckboxItem item in cbxl_vjestine.Items)
            //{
            //    //if item  is checked
            //    listavjestina1 = listavjestina1 + item.Value.ToString() + ",";
            //}


            string sql0 = "select x1.* " +
                        "from( select distinct k.id, mt.grupa1 gmt, k.prezimeime, k.funkcija, k.projekt,k.školovanje_posto, k.hala, k.linija, k.mjesto_troska," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 2) ar," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 14) ir," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 39) turm," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 15) vc_hurco," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 12) porche," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 16) sherer," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 17) okuma," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 18) puma_tt," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 19) iv," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 20) emag," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 21) valjci," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 22) pile," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 23) wupertal_prsten," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 24) iwk," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 25) pakiranje_iwk," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 38) membrane," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 26) skf," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 27) pakiranje_skf," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 28) glodanje_skf," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 29) vise_vretenasta," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 30) stupna," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 31) narezivanje_navoja," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 32) sona," +
                        "(select count(*) from radnicivjestine v where v.idradnika = k.id and idvjestine = 33) [3494]," +
                        "(select count(*) from radnicivjestine v where v.idradnika=k.id and idvjestine= 34) gt600," +
                        "(select count(*) from radnicivjestine v where v.idradnika=k.id and idvjestine= 35) kontrolaprstena100posto," +
                        "(select count(*) from radnicivjestine v where v.idradnika=k.id and idvjestine= 36) kontrolaporoznosti," +
                        "(select count(*) from radnicivjestine v where v.idradnika=k.id and idvjestine= 37) kontrolasona100posto,";
            //        string sql10="k.DatumZaposlenja,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt) x1 where id in (select idradnika from radnicivjestine where idvjestine in (" + listavjestina + ")) and x1.gmt='" + idloged + "' order by prezimeime";

            string sql1 = "select distinct k.* from kompetencije k left join radnicivjestine v on k.id=v.idradnika  order by prezimeime";
            string hala1 = tb_hala.Text.ToString();

            if (hala1 == "")
                hala1 = "'1','3',NULL,''";

            string sqllp = "x1.projekt='"+projekt1+"'";

            if (projekt1.Contains("Svi proj" ))
            {
               sqllp = "1=1";
            }

            if (projekt1.Contains("Nema proj"))
            {
                sqllp = "x1.projekt='' or x1.projekt is null";
            }


            if (idloged == "8")   // ako je grupa 8 ili tehn. direktor, može vidjeti sve, može printati
            {

                if (listavjestina == "")
                {
                    sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where x1.projekt='" + projekt1 + "' order by prezimeime";
                }

            }
            else
            {
                if (idadm == "14")  // tehnički direktor može vidjeti određene grupe  1- tokarija,steleri, 2 održavanje,,alatnica,brusenje 5 kaljenje 13 tehnologija
                {

                    if (listavjestina == "")
                    {
                        sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where x1.gmt in ( 1,2,5,13) order by prezimeime";
                    }
                    else
                    {
                        sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where id in (select idradnika from radnicivjestine where idvjestine in (" + listavjestina + ") and x1.gmt in ( 1,2,5,13) order by prezimeime";
                    }
                }
                else   // ostali korisnici, kicin etc.
                {
                    sql1 = sql0 + " k.DatumZaposlenja,k.istek_ugovora,k.napomena from kompetencije k left join radnici_ r on r.id=k.id left join mjestotroska mt on mt.id= r.mt where r.neradi=0) x1 where  ((("+sqllp+" ) and ( x1.hala in ("+hala1+") or x1.hala is null)) and x1.gmt='" + idloged + "') order by prezimeime";
                }


            }


            // sql1 = "select distinct k.* from kompetencije k left join radnicivjestine v on k.id=v.idradnika  order by prezimeime";

            //            if (listavjestina!="")
            //          {
            //    listavjestina = listavjestina.Substring(0, listavjestina.Length - 1);
            //   sql1 = "select distinct k.* from kompetencije k left join radnicivjestine v on k.id=v.idradnika  where v.idvjestine in (" + listavjestina + ") order by prezimeime";
            //          }

            lblvj_stat.Text = "Broj redova: 0";
            SqlConnection connection = new SqlConnection(connectionString);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            dgv_vjest_history.DataSource = ds;
            dgv_vjest_history.DataMember = "event";
            dgv_vjest_history.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgv_vjest_history.ReadOnly = false;
            sql1 = null; sql0 = null;
            lblvj_stat.Text = "Broj redova:" + (dgv_vjest_history.RowCount - 1).ToString();
            }

        private void combx_funkcije_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void label127_Click(object sender, EventArgs e)
        {

        }

        private void dgv_vjest_history_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row1 = dgv_vjest_history.CurrentRow;
            int columnIndex = dgv_vjest_history.CurrentCell.ColumnIndex;
            string idradnika=row1.Cells[0].Value.ToString();
            string value0 = (dgv_vjest_history.CurrentCell.Value.ToString());
           
                try
                {
                int value1 = (int.Parse)(value0);
                izmjenivjestinu(idradnika, columnIndex, value1);
            }
                catch
                {

                }
                finally
                {
                    
                }
           
        }

        private void izmjenivjestinu(string idradnika, int columnIndex, int value1)
        {
            int idvjestine1 = 0;
            switch (columnIndex)
                {
                case 19:
                           idvjestine1 = 72;
                            break;
                default:
                    idvjestine1 = 0;
                    break;

            }
            using (SqlConnection cn2 = new SqlConnection(connectionString))
            {
                cn2.Open();
                string sql1 = "delete from radnicivjestine where idradnika=" + idradnika + " and idvjestine=" + idvjestine1.ToString()  ;
                SqlCommand sqlCommand2 = new SqlCommand(sql1, cn2);
                SqlDataReader reader2 = sqlCommand2.ExecuteReader();
                cn2.Close();

                cn2.Open();
                sql1 = "insert into radnicivjestine(idradnika,idvjestine,vrijednost) values(" + idradnika + "," + idvjestine1 + "," + value1.ToString()+")";
                sqlCommand2 = new SqlCommand(sql1, cn2);
                reader2 = sqlCommand2.ExecuteReader();
                cn2.Close();

            }
        }

        private void pregledPromjenaVještinaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (idadm != "8")
            {
                return;
            }

            OcistiPanele();
            pnlv_log.Visible = true;
        }

        private void btnv_pretrazi_Click(object sender, EventArgs e)
        {
            
            DateTime dat1 = datepv_log_.Value;
            string sql1 = "select * from kompetencije_log where datumpromjene>='" + dat1.Year.ToString() + "-" + dat1.Month.ToString() + "-" + dat1.Day.ToString()+"' order by datumpromjene";
            SqlConnection connection = new SqlConnection(connectionString);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            dgv_log.DataSource = ds;
            dgv_log.DataMember = "event";
            dgv_log.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgv_log.ReadOnly = true;

        }

        private void datepv_log__ValueChanged(object sender, EventArgs e)
        {

        }

        private void pregledVještinaPoProjketuToolStripMenuItem_Click(object sender, EventArgs e)
        {

            string sql1;
            OcistiPanele();
            pnl_projektmtz.Visible = true;
            using (SqlConnection cn2 = new SqlConnection(connectionString))
            {
                cn2.Open();
                sql1 = "update projectmtz set linija=0;";
                sql1 = sql1 + "update projectmtz set available=0;";
                sql1 = sql1 + "update projectmtz set education=0;";
                sql1 = sql1 + "update projectmtz set dodatneoperacije=0;";
                sql1 = sql1 + "update projectmtz set linija=0;";
                sql1 = sql1 + "update projectmtz set okuma=0;";
                sql1 = sql1 + "update projectmtz set sherer=0;";
                sql1 = sql1 + "update projectmtz set emag=0;";
                sql1 = sql1 + "update projectmtz set iv=0;";
                sql1 = sql1 + "update projectmtz set leaving=0";
                SqlCommand sqlCommand2 = new SqlCommand(sql1, cn2);
                SqlDataReader reader2 = sqlCommand2.ExecuteReader();
                cn2.Close();
            }




            DateTime dat1 = datepv_log_.Value;
            sql1 = "select projekt, count(*) broj  from kompetencije k left join radnici_ r on r.id = k.id where r.neradi = 0 and k.mjesto_troska = 'Tokarenje' group by projekt order by projekt";
            SqlConnection cn = new SqlConnection(connectionString);
            cn.Open();
            SqlCommand sqlCommand = new SqlCommand(sql1, cn);
            SqlDataReader reader = sqlCommand.ExecuteReader();
            string projekt1 = "";
            int broj = 0;
            while (reader.Read())
            {
                // reader["Datum"].ToString();

                projekt1 = reader["Projekt"].ToString().TrimEnd();

                broj = (int.Parse)(reader["broj"].ToString());

                if (projekt1 == "BDF")
                    sql1 = "update projectmtz set available=" + broj.ToString() + " where project like '%AUSTRIA%';";

                if (projekt1 == "SKF")
                    sql1 = "update projectmtz set available=" + broj.ToString() + " where project like '%SKF%';";

                if (projekt1 == "IWK")
                    sql1 = "update projectmtz set available=" + broj.ToString() + " where project like '%KYSUCE%';";

                if (projekt1 == "SONA")
                    sql1 = "update projectmtz set available=" + broj.ToString() + " where project like '%SONA%';";

                if (projekt1 == "DEB")
                    sql1 = "update projectmtz set available=" + broj.ToString() + " where project like '%FAG%';";

                if (projekt1 == "VALJCI")
                    sql1 = "update projectmtz set available=" + broj.ToString() + " where project like '%ROLLERS%';";

                if (projekt1 == "BMW")
                    sql1 = "update projectmtz set available=" + broj.ToString() + " where project like '%SCHWEINFURT%';";

                if (projekt1 == "WUP")
                    sql1 = "update projectmtz set available=" + broj.ToString() + " where project like '%RINGE%';";


                if (projekt1 == "Š_BDF")
                    sql1 = sql1 + "update projectmtz set education=" + broj.ToString() + " where project like '%AUSTRIA%';";

                if (projekt1 == "Š_SKF")
                    sql1 = sql1 + "update projectmtz set education=" + broj.ToString() + " where project like '%SKF%';";

                if (projekt1 == "Š_IWK")
                    sql1 = sql1 + "update projectmtz set education=" + broj.ToString() + " where project like '%KYSUCE%';";

                if (projekt1 == "Š_SONA")
                    sql1 = sql1 + "update projectmtz set education=" + broj.ToString() + " where project like '%SONA%';";

                if (projekt1 == "Š_DEB")
                    sql1 = sql1 + "update projectmtz set education=" + broj.ToString() + " where project like '%FAG%';";

                if (projekt1 == "Š_VALJCI")
                    sql1 = sql1 + "update projectmtz set education=" + broj.ToString() + " where project like '%ROLLERS%';";

                if (projekt1 == "Š_BMW")
                    sql1 = sql1 + "update projectmtz set education=" + broj.ToString() + " where project like '%SCHWEINFURT%';";

                if (projekt1 == "Š_WUP")
                    sql1 = sql1 + "update projectmtz set education=" + broj.ToString() + " where project like '%RINGE%';";


                using (SqlConnection cn2 = new SqlConnection(connectionString))
                {
                    cn2.Open();
                    SqlCommand sqlCommand2 = new SqlCommand(sql1, cn2);
                    SqlDataReader reader2 = sqlCommand2.ExecuteReader();
                    cn2.Close();
                }


            }
            cn.Close();

            // analiza po vjestinama

            sql1 = "select count(*) broj from (select distinct id from( select projekt, count(*) broj, k.id, v.Naziv " +
                "from kompetencije k " +
                "left join RadniciVjestine rv on rv.idradnika = k.id " +
                "left join vjestine v on v.id = rv.idvjestine " +
                "left join radnici_ r on r.id = k.id  where r.neradi = 0 and k.mjesto_troska = 'Tokarenje'  and k.projekt = 'BDF' AND v.naziv = 'Sherer'  " +
                " group by projekt,k.id,v.naziv )x1)x2";


            //"and k.id not in ( select idradnika from radnicivjestine where idvjestine in (16, 17)) group by projekt,k.id,v.naziv )x1";//
            cn = new SqlConnection(connectionString);
            cn.Open();
            sqlCommand = new SqlCommand(sql1, cn);
            int broj1 = (int.Parse)(sqlCommand.ExecuteScalar().ToString());
            reader = sqlCommand.ExecuteReader();

            sql1 = sql1 = "update projectmtz set sherer=" + broj1.ToString() + " where project like '%AUSTRIA%';";
            using (SqlConnection cn2 = new SqlConnection(connectionString))
            {
                cn2.Open();
                SqlCommand sqlCommand2 = new SqlCommand(sql1, cn2);
                SqlDataReader reader2 = sqlCommand2.ExecuteReader();
                cn2.Close();
            }

            sql1 = "select count(*) broj from (select distinct id from( select projekt, count(*) broj, k.id, v.Naziv " +
                "from kompetencije k " +
                "left join RadniciVjestine rv on rv.idradnika = k.id " +
                "left join vjestine v on v.id = rv.idvjestine " +
                "left join radnici_ r on r.id = k.id  where r.neradi = 0 and k.mjesto_troska = 'Tokarenje'  and k.projekt = 'BDF' AND v.naziv = 'Okuma'  " +
                "and k.id not in ( select idradnika from radnicivjestine where idvjestine in (16)) group by projekt,k.id,v.naziv )x1)x2";



            //"and k.id not in ( select idradnika from radnicivjestine where idvjestine in (16, 17)) group by projekt,k.id,v.naziv )x1";//
            cn = new SqlConnection(connectionString);
            cn.Open();
            sqlCommand = new SqlCommand(sql1, cn);
            broj1 = (int.Parse)(sqlCommand.ExecuteScalar().ToString());
            reader = sqlCommand.ExecuteReader();

            sql1 = sql1 = "update projectmtz set okuma=" + broj1.ToString() + " where project like '%AUSTRIA%';";
            using (SqlConnection cn2 = new SqlConnection(connectionString))
            {
                cn2.Open();
                SqlCommand sqlCommand2 = new SqlCommand(sql1, cn2);
                SqlDataReader reader2 = sqlCommand2.ExecuteReader();
                cn2.Close();
            }

            sql1 = "select count(*) broj from (select distinct id from( select projekt, count(*) broj, k.id, v.Naziv " +
                "from kompetencije k " +
                "left join RadniciVjestine rv on rv.idradnika = k.id " +
                "left join vjestine v on v.id = rv.idvjestine " +
                "left join radnici_ r on r.id = k.id  where r.neradi = 0 and k.mjesto_troska = 'Tokarenje'  and k.projekt = 'BDF' AND v.naziv = 'IV'  " +
                "and k.id not in ( select idradnika from radnicivjestine where idvjestine in (16,17)) group by projekt,k.id,v.naziv )x1)x2";



            //"and k.id not in ( select idradnika from radnicivjestine where idvjestine in (16, 17)) group by projekt,k.id,v.naziv )x1";//
            cn = new SqlConnection(connectionString);
            cn.Open();
            sqlCommand = new SqlCommand(sql1, cn);
            broj1 = (int.Parse)(sqlCommand.ExecuteScalar().ToString());
            reader = sqlCommand.ExecuteReader();

            sql1 = sql1 = "update projectmtz set iv=" + broj1.ToString() + " where project like '%AUSTRIA%';";
            using (SqlConnection cn2 = new SqlConnection(connectionString))
            {
                cn2.Open();
                SqlCommand sqlCommand2 = new SqlCommand(sql1, cn2);
                SqlDataReader reader2 = sqlCommand2.ExecuteReader();
                cn2.Close();
            }

            sql1 = "select count(*) broj from (select distinct id from( select projekt, count(*) broj, k.id, v.Naziv " +
                "from kompetencije k " +
                "left join RadniciVjestine rv on rv.idradnika = k.id " +
                "left join vjestine v on v.id = rv.idvjestine " +
                "left join radnici_ r on r.id = k.id  where r.neradi = 0 and k.mjesto_troska = 'Tokarenje'  and k.projekt = 'BDF' AND v.naziv = 'VC Hurco'  " +
                "and k.id not in ( select idradnika from radnicivjestine where idvjestine in (16,17,19)) group by projekt,k.id,v.naziv )x1)x2";



            //"and k.id not in ( select idradnika from radnicivjestine where idvjestine in (16, 17)) group by projekt,k.id,v.naziv )x1";//
            cn = new SqlConnection(connectionString);
            cn.Open();
            sqlCommand = new SqlCommand(sql1, cn);
            broj1 = (int.Parse)(sqlCommand.ExecuteScalar().ToString());
            reader = sqlCommand.ExecuteReader();

            sql1 = sql1 = "update projectmtz set dodatneoperacije=" + broj1.ToString() + " where project like '%AUSTRIA%';";
            using (SqlConnection cn2 = new SqlConnection(connectionString))
            {
                cn2.Open();
                SqlCommand sqlCommand2 = new SqlCommand(sql1, cn2);
                SqlDataReader reader2 = sqlCommand2.ExecuteReader();
                cn2.Close();
            }

            sql1 = "select count(*) broj from (select distinct id from( select projekt, count(*) broj, k.id, v.Naziv " +
                "from kompetencije k " +
                "left join RadniciVjestine rv on rv.idradnika = k.id " +
                "left join vjestine v on v.id = rv.idvjestine " +
                "left join radnici_ r on r.id = k.id  where r.neradi = 0 and k.mjesto_troska = 'Tokarenje'  and k.projekt = 'SKF' AND   " +
                "k.id in ( select idradnika from radnicivjestine where idvjestine in (28,27,29,31,30)) group by projekt,k.id,v.naziv )x1)x2";

            //"and k.id not in ( select idradnika from radnicivjestine where idvjestine in (16, 17)) group by projekt,k.id,v.naziv )x1";//
            cn = new SqlConnection(connectionString);
            cn.Open();
            sqlCommand = new SqlCommand(sql1, cn);
            broj1 = (int.Parse)(sqlCommand.ExecuteScalar().ToString());
            reader = sqlCommand.ExecuteReader();

            sql1 = sql1 = "update projectmtz set dodatneoperacije=" + broj1.ToString() + " where project like '%SKF%';";
            using (SqlConnection cn2 = new SqlConnection(connectionString))
            {
                cn2.Open();
                SqlCommand sqlCommand2 = new SqlCommand(sql1, cn2);
                SqlDataReader reader2 = sqlCommand2.ExecuteReader();
                cn2.Close();
            }

            sql1 = "select count(*) broj from (select distinct id from( select projekt, count(*) broj, k.id, v.Naziv " +
                "from kompetencije k " +
                "left join RadniciVjestine rv on rv.idradnika = k.id " +
                "left join vjestine v on v.id = rv.idvjestine " +
                "left join radnici_ r on r.id = k.id  where r.neradi = 0 and k.mjesto_troska = 'Tokarenje'  and k.projekt = 'IWK' AND   " +
                "k.id in ( select idradnika from radnicivjestine where idvjestine in (36,25)) group by projekt,k.id,v.naziv )x1)x2";

            //"and k.id not in ( select idradnika from radnicivjestine where idvjestine in (16, 17)) group by projekt,k.id,v.naziv )x1";//
            cn = new SqlConnection(connectionString);
            cn.Open();
            sqlCommand = new SqlCommand(sql1, cn);
            broj1 = (int.Parse)(sqlCommand.ExecuteScalar().ToString());
            reader = sqlCommand.ExecuteReader();

            sql1 = sql1 = "update projectmtz set dodatneoperacije=" + broj1.ToString() + " where project like '%KYSUCE%';";
            using (SqlConnection cn2 = new SqlConnection(connectionString))
            {
                cn2.Open();
                SqlCommand sqlCommand2 = new SqlCommand(sql1, cn2);
                SqlDataReader reader2 = sqlCommand2.ExecuteReader();
                cn2.Close();
            }

            sql1 = "select count(*) broj from (select distinct id from( select projekt, count(*) broj, k.id, v.Naziv " +
               "from kompetencije k " +
               "left join RadniciVjestine rv on rv.idradnika = k.id " +
               "left join vjestine v on v.id = rv.idvjestine " +
               "left join radnici_ r on r.id = k.id  where r.neradi = 0 and k.mjesto_troska = 'Tokarenje'  and k.projekt = 'BMW' AND   " +
               "k.id in ( select idradnika from radnicivjestine where idvjestine in (20)) group by projekt,k.id,v.naziv )x1)x2";

            //"and k.id not in ( select idradnika from radnicivjestine where idvjestine in (16, 17)) group by projekt,k.id,v.naziv )x1";//
            cn = new SqlConnection(connectionString);
            cn.Open();
            sqlCommand = new SqlCommand(sql1, cn);
            broj1 = (int.Parse)(sqlCommand.ExecuteScalar().ToString());
            reader = sqlCommand.ExecuteReader();

            sql1 = sql1 = "update projectmtz set emag=" + broj1.ToString() + " where project like '%SCHWEINFURT%';";
            using (SqlConnection cn2 = new SqlConnection(connectionString))
            {
                cn2.Open();
                SqlCommand sqlCommand2 = new SqlCommand(sql1, cn2);
                SqlDataReader reader2 = sqlCommand2.ExecuteReader();
                cn2.Close();
            }

            sql1 = "select count(*) broj from (select distinct id from( select projekt, count(*) broj, k.id, v.Naziv " +
              "from kompetencije k " +
              "left join RadniciVjestine rv on rv.idradnika = k.id " +
              "left join vjestine v on v.id = rv.idvjestine " +
              "left join radnici_ r on r.id = k.id  where r.neradi = 0 and k.mjesto_troska = 'Tokarenje'  and k.projekt = 'SONA' AND   " +
              "k.id in ( select idradnika from radnicivjestine where idvjestine in (34,37)) group by projekt,k.id,v.naziv )x1)x2";

            //"and k.id not in ( select idradnika from radnicivjestine where idvjestine in (16, 17)) group by projekt,k.id,v.naziv )x1";//
            cn = new SqlConnection(connectionString);
            cn.Open();
            sqlCommand = new SqlCommand(sql1, cn);
            broj1 = (int.Parse)(sqlCommand.ExecuteScalar().ToString());
            reader = sqlCommand.ExecuteReader();

            sql1 = sql1 = "update projectmtz set dodatneoperacije=" + broj1.ToString() + " where project like '%SONA%';";
            using (SqlConnection cn2 = new SqlConnection(connectionString))
            {
                cn2.Open();
                SqlCommand sqlCommand2 = new SqlCommand(sql1, cn2);
                SqlDataReader reader2 = sqlCommand2.ExecuteReader();
                cn2.Close();
            }

            using (SqlConnection cn2 = new SqlConnection(connectionString))
            {
                cn2.Open();
                sql1 = "update projectmtz set linija=(isnull(available,0) - isnull(okuma,0)- isnull(sherer,0)-isnull(iv,0)-isnull(emag,0)-isnull(dodatneoperacije,0));";
                sql1 = sql1 + "update projectmtz set available=j.suma from ( select sum(available) suma from projectmtz where project!='Z  UKUPNO') j where project='Z  UKUPNO';";
                sql1 = sql1 + "update projectmtz set education=j.suma from ( select sum(education) suma from projectmtz where project!='Z  UKUPNO') j where project='Z  UKUPNO';";
                sql1 = sql1 + "update projectmtz set dodatneoperacije=j.suma from ( select sum(dodatneoperacije) suma from projectmtz where project!='Z  UKUPNO') j where project='Z  UKUPNO';";
                sql1 = sql1 + "update projectmtz set linija=j.suma from ( select sum(linija) suma from projectmtz where project!='Z  UKUPNO') j where project='Z  UKUPNO';";
                sql1 = sql1 + "update projectmtz set okuma=j.suma from ( select sum(okuma) suma from projectmtz where project!='Z  UKUPNO') j where project='Z  UKUPNO';";
                sql1 = sql1 + "update projectmtz set sherer=j.suma from ( select sum(sherer) suma from projectmtz where project!='Z  UKUPNO') j where project='Z  UKUPNO';";
                sql1 = sql1 + "update projectmtz set emag=j.suma from ( select sum(emag) suma from projectmtz where project!='Z  UKUPNO') j where project='Z  UKUPNO';";
                sql1 = sql1 + "update projectmtz set iv=j.suma from ( select sum(iv) suma from projectmtz where project!='Z  UKUPNO') j where project='Z  UKUPNO';";
                sql1 = sql1 + "update projectmtz set leaving=j.suma from ( select count(*) suma from kompetencije k left join radnici_ r on k.id=r.id where r.neradi=0 and projekt='Na odlasku' and k.mjesto_troska='Tokarenje') j where project='Z  UKUPNO'";
                SqlCommand sqlCommand2 = new SqlCommand(sql1, cn2);
                SqlDataReader reader2 = sqlCommand2.ExecuteReader();
                cn2.Close();
            }

            using (SqlConnection cn2 = new SqlConnection(connectionString))
            {
                cn2.Open();
                sql1 = "update projectmtz set P_linija=j.suma from ( select sum(p_linija) suma from projectmtz where project!='Z  UKUPNO') j where project='Z  UKUPNO';";
                sql1 = sql1 + "update projectmtz set p_okuma=j.suma from ( select sum(p_okuma) suma from projectmtz where project!='Z  UKUPNO') j where project='Z  UKUPNO';";
                sql1 = sql1 + "update projectmtz set p_emag=j.suma from ( select sum(p_emag) suma from projectmtz where project!='Z  UKUPNO') j where project='Z  UKUPNO';";
                sql1 = sql1 + "update projectmtz set p_sherer=j.suma from ( select sum(p_sherer) suma from projectmtz where project!='Z  UKUPNO') j where project='Z  UKUPNO';";
                sql1 = sql1 + "update projectmtz set p_iv=j.suma from ( select sum(p_iv) suma from projectmtz where project!='Z  UKUPNO') j where project='Z  UKUPNO';";
                sql1 = sql1 + "update projectmtz set p_dodatneoperacije=j.suma from ( select sum(p_dodatneoperacije) suma from projectmtz where project!='Z  UKUPNO') j where project='Z  UKUPNO';";
                sql1 = sql1 + "update projectmtz set difference=(isnull(available,0) -  isnull(P_okuma,0)- isnull(P_sherer,0)-isnull(P_iv,0)-isnull(P_emag,0)-isnull(P_linija,0) -isnull(P_dodatneoperacije,0) )";
                SqlCommand sqlCommand2 = new SqlCommand(sql1, cn2);
                SqlDataReader reader2 = sqlCommand2.ExecuteReader();
                cn2.Close();
            }


            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);

            sql1 = "select * from  projectmtz order by project";
            SqlConnection connection = new SqlConnection(connectionString);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            dgv_projektmtz.DataSource = ds;
            dgv_projektmtz.DataMember = "event";
            dgv_projektmtz.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgv_projektmtz.ReadOnly = true;

            

            sql1 = "select available,education from projectmtz where project='Z  UKUPNO'";

            SqlConnection conection = new SqlConnection(connectionString);
            connection.Open();
            sqlCommand = new SqlCommand(sql1, connection);
            reader = sqlCommand.ExecuteReader();

            while (reader.Read())
            {
                label135.Text = reader["available"].ToString().TrimEnd();
                label136.Text = reader["education"].ToString().TrimEnd();                
            }
            connection.Close();

            sql1 = "select sum(x1.broj) brojn from (select count(*) broj from kompetencije k left join radnici_ r on r.id = k.id where r.neradi = 0 and k.mjesto_troska = 'Tokarenje' and r.mt = 700 AND(PROJEKT IN('Dugotrajno bolovanje', 'Na odlasku', 'Nema projekt', 'NERADI', null, '') or projekt is null) group by projekt) x1";

            conection = new SqlConnection(connectionString);
            connection.Open();
            sqlCommand = new SqlCommand(sql1, connection);
            reader = sqlCommand.ExecuteReader();
            // nis u projektu
            while (reader.Read())
            {
                label137.Text = reader["brojn"].ToString().TrimEnd();
            }
            connection.Close();

            sql1 = "select count(*) broj from radnici_  where mt=716";

            conection = new SqlConnection(connectionString);
            connection.Open();
            sqlCommand = new SqlCommand(sql1, connection);
            reader = sqlCommand.ExecuteReader();
            // Šteleri
            while (reader.Read())
            {
                label138.Text = reader["broj"].ToString().TrimEnd();
            }
            label140.Text = ((int.Parse)(label135.Text) + (int.Parse)(label136.Text) + (int.Parse)(label137.Text) ).ToString();
            connection.Close();


        }
        private void btn_mtzexcel_Click(object sender, EventArgs e)
        {
            //Creating DataTable
            DataTable dt = new DataTable();

            //Adding the Columns
            foreach (DataGridViewColumn column in dgv_projektmtz.Columns)
            {

                // satnica se pokazuje samo za grupu 8 ili tehn. direktora (14)
                dt.Columns.Add(column.HeaderText, column.ValueType);


            }

            //Adding the Rows
            foreach (DataGridViewRow row in dgv_projektmtz.Rows)
            {
                dt.Rows.Add();

                foreach (DataGridViewCell cell in row.Cells)
                {


                    if (!(cell.Value is DBNull))
                    {
                        if (cell.Value != null)

                        {
                            dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();

                        }
                    }
                }
            }

            //if ((idloged == "8") || (idadm == "14"))
            //{

            //}
            //else
            //{
            //    // satnica se pokazuje samo za grupu 8 ili tehn. direktora (14)
            //    dt.Columns.RemoveAt(9);
            //    dt.Columns.RemoveAt(9);
            //    dt.Columns.RemoveAt(9);
            //    dt.Columns.RemoveAt(9);
            //}
            //Exporting to Excel
            string folderPath = "C:\\KKS\\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "Kompetencije");
                wb.SaveAs(folderPath + "projektmtz.xlsx");
            }
            btn_mtzexcel.Text = "Export done";

            FileInfo fi = new FileInfo("C:\\kks\\projektmtz.xlsx");
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(@"C:\kks\projektmtz.xlsx");
            }
            else
            {
                //file doesn't exist
            }

            //Object oMissing = System.Reflection.Missing.Value;
            //oMissing = System.Reflection.Missing.Value;
            //Object oTrue = true;
            //Object oFalse = false;
            //Microsoft.Office.Interop.Word.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Document oExcelDoc = new Microsoft.Office.Interop.Excel.Document();
            //oExcel.Visible = true;

            //Object oTemplatePath = "c:\\kks\\kompetencije.xlsx";
            //oExcelDoc = oExcelDoc.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);





        }

        private void dgv_projektmtz_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgv_vjest_history_SelectionChanged(object sender, EventArgs e)
        {
            
        }

        private void dgv_vjest_history_MouseLeave(object sender, EventArgs e)
        {
            lblv_odabrano.Text = " Odabrano " + dgv_vjest_history.SelectedCells.Count.ToString();
        }

        private void dgv_vjest_history_MouseEnter(object sender, EventArgs e)
        {
            
        }

        private void dgv_vjest_history_CellMouseUp(object sender, DataGridViewCellMouseEventArgs e)
        {
            lblv_odabrano.Text = " Odabrano " + dgv_vjest_history.SelectedCells.Count.ToString();
        }

        private void pregledVještinaPoProjektulinijamaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OcistiPanele();
            pnl_projektmtz.Visible = true;

            string sql1 = "select Projekt, k.Linija,v.Naziv Naziv_vještine, count(*) Broj  from kompetencije k left join RadniciVjestine rv on rv.idradnika = k.id left join vjestine v on v.id = rv.idvjestine left join radnici_ r on r.id = k.id where r.neradi = 0 and k.mjesto_troska = 'Tokarenje' group by projekt,k.linija,v.naziv order by projekt";
            SqlConnection connection = new SqlConnection(connectionString);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();
            dgv_projektmtz.Width = 500;
            dgv_projektmtz.Height = 800;
            pnl_projektmtz.Height = 1200;

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            dgv_projektmtz.DataSource = ds;
            dgv_projektmtz.DataMember = "event";
            dgv_projektmtz.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgv_projektmtz.ReadOnly = true;

        }

        private void pregledDjslartnToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void pregledDjelatnikaNaOdlaskuToolStripMenuItem_Click(object sender, EventArgs e)
        {
            //int firmaid = comb_poduzece.SelectedIndex;
            //string mtid = combo_odred.SelectedValue.ToString();
            pnl_naodlasku.Visible = true;
            DateTime dat1 = datep_od_ov.Value;
            DateTime dat2 = datp_do_odr.Value;

            string dat10 = dat1.Year.ToString() + '-' + dat1.Month.ToString() + '-' + dat1.Day.ToString();
            string dat20 = dat2.Year.ToString() + '-' + dat2.Month.ToString() + '-' + dat2.Day.ToString();

            string sql1 = "select k.id,k.prezimeime,k.funkcija,k.projekt,k.hala,k.linija,k.mjesto_troska,k.vještine,k.školovanje_posto,k.radnomjesto,k.datumzaposlenja,k.istek_ugovora,k.staz,j.danodlaska,k.godisnji_ostalo from kompetencije k left join ( select * from fxsap.dbo.plansatirada where danodlaska != '' and danodlaska<(dateadd(day,30,getdate())) and danodlaska> (dateadd(day, -30, getdate()))  and danodlaska is not NULL" +
                          ") j on j.radnikid = k.id where projekt not in ('xNERADI', 'xNa odlasku') and j.danodlaska is not null order by PrezimeIme desc ";
            
            string connecStrin = connectionString;

            //if (firmaid == 1)
            {
            //    connecStrin = connectionStringpb;
            }

            SqlConnection connection = new SqlConnection(connecStrin);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            dgv_naodlasku.DataSource = ds;
            dgv_naodlasku.DataMember = "event";
            dgv_naodlasku.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void pregledNovihDjelatnikaToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void pregledNovihDjelatnikaToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            OcistiPanele();

            pnl_novi_djelatnici.Visible = true;
            string sql1 = "";
            if (idloged == "8")
                {
                sql1 = "select cast(acregno as int ) as Id,acname Ime,acsurname Prezime,j.acCostDrv Mjesto_troška,j.acDept Lokacija,d.acnumber RFID,j.acjob Radno_mjesto,'' Radni_staz,j.adDate DatumZaposlenja,'' sifra_rm,p.acstreet Ulica,acpost Pošta,'' vrijeme,accity Grad,j.acFieldSA Vrsta_isplate, adDateExit Datum_odlaska from thr_prsn p left join thr_prsnjob j on p.acworker = j.acworker left join thr_prsnadddoc d on d.acWorker = p.acworker and d.actype = 8 where d.actype=8 and j.addateend is null " +
                "order by cast(acregno as int) desc ";
            }
            else
            {
                sql1 = "select cast(acregno as int ) as Id,acname Ime,acsurname Prezime,j.acCostDrv Mjesto_troška,j.acDept Lokacija,d.acnumber RFID,j.acjob Radno_mjesto,'' Radni_staz,j.adDate DatumZaposlenja,'' sifra_rm,p.acstreet Ulica,acpost Pošta,'' vrijeme,accity Grad,j.acFieldSA Vrsta_isplate, adDateExit Datum_odlaska from thr_prsn p left join thr_prsnjob j on p.acworker = j.acworker left join thr_prsnadddoc d on d.acWorker = p.acworker and d.actype = 8 " +
                "where j.addateend is null and d.actype=8 and addateenter>=dateadd(d,-3,getdate()) order by cast(acregno as int) desc ";
            }
            string connecStrin = connectionStringpa;
            if (chbx_Feroimpex.Checked)
            {
                connecStrin = connectionStringpa;
                chbx_TKB.Checked = false;
            }
            if (chbx_TKB.Checked)
            {
                connecStrin = connectionStringpb;
                chbx_Feroimpex.Checked = false;
            }

            //if (firmaid == 1)
            {
                //    connecStrin = connectionStringpb;
            }

            SqlConnection connection = new SqlConnection(connecStrin);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            dgv_novi_djelatnici.DataSource = ds;
            dgv_novi_djelatnici.DataMember = "event";
            dgv_novi_djelatnici.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
        }

        private void chbx_Feroimpex_CheckedChanged(object sender, EventArgs e)
        {
            OcistiPanele();

            pnl_novi_djelatnici.Visible = true;
            string sql1 = "";
            if (idloged == "8")
            {
                sql1 = "select acregno  as Id,acname Ime,acsurname Prezime,j.acCostDrv Mjesto_troška,j.acDept Lokacija,d.acnumber RFID,j.acjob Radno_mjesto,'' Radni_staz,j.adDate DatumZaposlenja,'' sifra_rm,p.acstreet Ulica,acpost Pošta,'' vrijeme,accity Grad,j.acFieldSA Vrsta_isplate, adDateExit Datum_odlaska from thr_prsn p left join thr_prsnjob j on p.acworker = j.acworker left join thr_prsnadddoc d on d.acWorker = p.acworker and d.actype = 8 where d.actype=8 and j.addateend is null " +
                "order by cast(acregno as int)  desc ";
//                sql1 = "select cast(acregno as int ) as Id,acname Ime,acsurname Prezime,j.acCostDrv Mjesto_troška,j.acDept Lokacija,d.acnumber RFID,j.acjob Radno_mjesto,'' Radni_staz,p.adDateEnter DatumZaposlenja,'' sifra_rm,p.acstreet Ulica,acpost Pošta,'' vrijeme,accity Grad,j.acFieldSA Vrsta_isplate, adDateExit Datum_odlaska from thr_prsn p left join thr_prsnjob j on p.acworker = j.acworker left join thr_prsnadddoc d on d.acWorker = p.acworker and d.actype = 8 where d.actype=8 and j.addateend is null " +
//                "order by cast(acregno as int) desc ";
            }
            else
            {
                sql1 = "select cast(acregno as int ) as Id,acname Ime,acsurname Prezime,j.acCostDrv Mjesto_troška,j.acDept Lokacija,d.acnumber RFID,j.acjob Radno_mjesto,'' Radni_staz,j.adDate DatumZaposlenja,'' sifra_rm,p.acstreet Ulica,acpost Pošta,'' vrijeme,accity Grad,j.acFieldSA Vrsta_isplate, adDateExit Datum_odlaska from thr_prsn p left join thr_prsnjob j on p.acworker = j.acworker left join thr_prsnadddoc d on d.acWorker = p.acworker and d.actype = 8 " +
                "where j.addateend is null and d.actype=8 and p.addateenter>=dateadd(d,-3,getdate()) order by cast(acregno as int) desc ";
            }
            string connecStrin = connectionStringpa;
            if (chbx_Feroimpex.Checked)
            {
                connecStrin = connectionStringpa;
                chbx_TKB.Checked = false;
            }
            else
            {
                connecStrin = connectionStringpb;
                chbx_Feroimpex.Checked = false;
            }

            //if (firmaid == 1)
            {
                //    connecStrin = connectionStringpb;
            }

            SqlConnection connection = new SqlConnection(connecStrin);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            dgv_novi_djelatnici.DataSource = ds;
            dgv_novi_djelatnici.DataMember = "event";
            dgv_novi_djelatnici.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

        }
        // kadrovski izvještaj 
        private void button10_Click(object sender, EventArgs e)
        {
            //Creating DataTable
            DataTable dt = new DataTable();

            //Adding the Columns
            foreach (DataGridViewColumn column in dgv_novi_djelatnici.Columns)
            {
                // satnica se pokazuje samo za grupu 8 ili tehn. direktora (14)
                dt.Columns.Add(column.HeaderText, column.ValueType);
            }

            //Adding the Rows
            foreach (DataGridViewRow row in dgv_novi_djelatnici.Rows)
            {
                dt.Rows.Add();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (!(cell.Value is DBNull))
                    {
                        if (cell.Value != null)

                        {
                            dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();

                        }
                    }
                }
            }

            //if ((idloged == "8") || (idadm == "14"))
            //{

            //}
            //else
            //{
            //    // satnica se pokazuje samo za grupu 8 ili tehn. direktora (14)
            //    dt.Columns.RemoveAt(9);
            //    dt.Columns.RemoveAt(9);
            //    dt.Columns.RemoveAt(9);
            //    dt.Columns.RemoveAt(9);
            //}
            //Exporting to Excel
            string folderPath = "C:\\KKS\\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "Kompetencije");
                wb.SaveAs(folderPath + "Popisdjelatnika1.xlsx");
            }
            btn_zbirni_export.Text = "Export done";

            FileInfo fi = new FileInfo("C:\\kks\\Popisdjelatnika1.xlsx");
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(@"C:\kks\Popisdjelatnika1.xlsx");
            }
            else
            {
                //file doesn't exist
            }

            //Object oMissing = System.Reflection.Missing.Value;
            //oMissing = System.Reflection.Missing.Value;
            //Object oTrue = true;
            //Object oFalse = false;
            //Microsoft.Office.Interop.Word.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Document oExcelDoc = new Microsoft.Office.Interop.Excel.Document();
            //oExcel.Visible = true;

            //Object oTemplatePath = "c:\\kks\\kompetencije.xlsx";
            //oExcelDoc = oExcelDoc.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);
            
        }

        private void btn_pregledrfid_Click(object sender, EventArgs e)
        {
            dgv_rfid.Visible = true;
            pnl_rfid.Visible = true;
            btn_hide_rfid.Visible = true;

            //button2.Visible = true;
            //button2.Text = "Export to excell ";
            string connectionString = @"Data Source=192.168.0.3;Initial Catalog=RFIND;User ID=fx_public;Password=.";
            if (combo_listadjelatnika.SelectedIndex <= 0)
            {
                return;
            }

            int ssindex = int.Parse(combo_listadjelatnika.SelectedValue.ToString());
            ssindex1 = ssindex;

            

            //string ime1 = label24.Text.TrimEnd();
            string sql;
            //int lokacija1 = comboBox1.SelectedIndex;
            DateTime dat10 = DateTime.Now.AddDays(1);
            string dat1 = dat10.Year.ToString() + '-' + dat10.Month.ToString() + '-'+dat10.Day.ToString();


//            string dat1 = dateTimePicker1.Value.Year + "-" + dateTimePicker1.Value.Month + "-" + dateTimePicker1.Value.Day;
//            string dat2 = dateTimePicker2.Value.Year + "-" + dateTimePicker2.Value.Month + "-" + dateTimePicker2.Value.Day + " 23:59:59.00";
            string lokacija;

            lokacija = "";
            //switch (lokacija1 + 1)
            //{
            //    case 1:
            //        lokacija = "_";    // sve
            //        break;
            //    case 2:
            //        lokacija = "22293";    // uprava ulaz
            //        break;
            //    case 3:
            //        lokacija = "544666574";    // tehnologija
            //        break;
            //    case 4:
            //        lokacija = "544666577";   // garderoba p1
            //        break;
            //    case 5:
            //        lokacija = "544666590";   // hala3
            //        break;
            //    case 6:
            //        lokacija = "544666595";   // hala4
            //        break;
            //    case 7:
            //        lokacija = "544666584";   // zona
            //        break;
            //    default:
            //        break;
            //}


            //if (lokacija == "" || lokacija == "_")
                sql = "select dt Vrijeme,LastName Prezime,u.FirstName Ime,r.name Lokacija,b.extid FxId,e.[User],e.No2 Serial_number,e.no RFID_Hex from event e left join badge b on e.No= b.BadgeNo left join [dbo].[User] u on u.extid=b.extid left join eventtype t on e.EventType=t.Code left join reader r on r.id=e.device_id WHERE E.[USER] IS NOT NULL and e.eventtype!='SP23'  and ( e.dt<='" + dat1 + "') AND u.extid="+ssindex.ToString();
            //else
            //    sql = "select dt Vrijeme,LastName Prezime,u.FirstName Ime,r.name Lokacija,e.Device_ID Uredaj,e.EventType,t.CodeName,b.extid FxId,e.[User],e.No2 Serial_number,e.no RFID_Hex from event e left join badge b on e.No= b.BadgeNo left join [dbo].[User] u on u.extid=b.extid left join eventtype t on e.EventType=t.Code left join reader r on r.id=e.device_id WHERE E.[USER] IS NOT NULL and e.eventtype!='SP23'  and ( e.dt>='" + dat1 + "' and e.dt<='" + dat2 + "') AND  ( lastname+' '+firstname) like '%" + textBox1.Text + "%'  and  e.device_id='" + lokacija + "') ";

            //if (checkBox1.Checked)
            //{
            //    sql = sql + "and eventtype='SP40' order by dt desc";
            //}
            //else
            {
                sql = sql + "order by dt desc";
            }

            //sql = "select dt Vrijeme,LastName Prezime,u.FirstName Ime,r.name Lokacija,e.Device_ID Uredaj,e.EventType,t.CodeName,b.extid FxId,e.[User],e.No2 Serial_number,e.no RFID_Hex from event e left join badge b on e.No= b.BadgeNo left join [dbo].[User] u on u.extid=b.extid left join eventtype t on e.EventType=t.Code left join reader r on r.id=e.device_id WHERE E.[USER] IS NOT NULL and e.dt>='" + dat1+"'";

            SqlConnection connection = new SqlConnection(connectionString);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            dgv_rfid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            dgv_rfid.DataSource = ds;
            dgv_rfid.DataMember = "event";

            dgv_rfid.AutoResizeColumns();
            dgv_rfid.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

            //dataGridView1.Width = System.Windows.Forms.Screen.PrimaryScreen.Bounds.Width - 120;

            
        }

        private void btn_hide_rfid_Click(object sender, EventArgs e)
        {
            btn_hide_rfid.Visible = false;
            pnl_rfid.Visible = false;
            panel2.Visible = true;
        }

        private void pregledŠkolovanjaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pnl_preglSkolovanja.Visible = true;
            string sql1 = "select s.*,r.datumzaposlenja Datum_zaposlenja,k.školovanje_posto,k.funkcija Funkcija from skolovanje s left join radnici_ r left join kompetencije k on k.id=r.id on r.id=s.idradnika where gotovo=0 order by k.prezimeime,izmjenaod";
            string connecStrin = @"Data Source=192.168.0.3;Initial Catalog=RFIND;User ID=fx_public;Password=.";

            SqlConnection connection = new SqlConnection(connecStrin);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            dgv_naskolovanju.DataSource = ds;
            dgv_naskolovanju.DataMember = "event";
            dgv_naskolovanju.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

        }

        private void btn_export_excel_skolovanje_Click(object sender, EventArgs e)
        {

            //Creating DataTable
            DataTable dt = new DataTable();

            //Adding the Columns
            foreach (DataGridViewColumn column in dgv_naskolovanju.Columns)
            {

                // satnica se pokazuje samo za grupu 8 ili tehn. direktora (14)
                dt.Columns.Add(column.HeaderText, column.ValueType);


            }

            //Adding the Rows
            foreach (DataGridViewRow row in dgv_naskolovanju.Rows)
            {
                dt.Rows.Add();

                foreach (DataGridViewCell cell in row.Cells)
                {


                    if (!(cell.Value is DBNull))
                    {
                        if (cell.Value != null)

                        {
                            dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();

                        }
                    }
                }
            }

            //if ((idloged == "8") || (idadm == "14"))
            //{

            //}
            //else
            //{
            //    // satnica se pokazuje samo za grupu 8 ili tehn. direktora (14)
            //    dt.Columns.RemoveAt(9);
            //    dt.Columns.RemoveAt(9);
            //    dt.Columns.RemoveAt(9);
            //    dt.Columns.RemoveAt(9);
            //}
            //Exporting to Excel
            string folderPath = "C:\\KKS\\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "Kompetencije");
                wb.SaveAs(folderPath + "Skolovanje.xlsx");
            }
            btn_zbirni_export.Text = "Export done";

            FileInfo fi = new FileInfo("C:\\kks\\skolovanje.xlsx");
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(@"C:\kks\skolovanje.xlsx");
            }
            else
            {
                //file doesn't exist
            }

            //Object oMissing = System.Reflection.Missing.Value;
            //oMissing = System.Reflection.Missing.Value;
            //Object oTrue = true;
            //Object oFalse = false;
            //Microsoft.Office.Interop.Word.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Document oExcelDoc = new Microsoft.Office.Interop.Excel.Document();
            //oExcel.Visible = true;

            //Object oTemplatePath = "c:\\kks\\kompetencije.xlsx";
            //oExcelDoc = oExcelDoc.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

        }

        private void dgv_vjest_history_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgv_vjest_history_CellMouseLeave(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dgv_vjest_history_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row1 = dgv_vjest_history.CurrentRow;
            int columnIndex = dgv_vjest_history.CurrentCell.ColumnIndex;
            string idradnika = row1.Cells[0].Value.ToString();
        }

        private void dgv_vjest_history_CellMouseEnter(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row1 = dgv_vjest_history.CurrentRow;
            int columnIndex = dgv_vjest_history.CurrentCell.ColumnIndex;
            
            string idradnika = row1.Cells[0].Value.ToString();
        }

        private void dgv_vjest_history_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            DataGridViewRow row1 = dgv_vjest_history.CurrentRow;
            int columnIndex = dgv_vjest_history.CurrentCell.ColumnIndex;
            string idradnika = row1.Cells[0].Value.ToString();
            string value0 = (dgv_vjest_history.CurrentCell.Value.ToString());

            try
            {
                int value1 = (int.Parse)(value0);
                izmjenivjestinu(idradnika, columnIndex, value1);
            }
            catch
            {

            }
            finally
            {

            }
        }

        

        private void dgv_vjest_history_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dgv_vjest_history.CurrentCell.ColumnIndex >= 4)  //example-'Column index=4'
            {
                dgv_vjest_history.BeginEdit(true);
            }
        }

        private void tb_hala_TextChanged(object sender, EventArgs e)
        {

        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {



        }

        private void dataGridView1_DoubleClick(object sender, EventArgs e)
        {


        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            SqlConnection cnn1 = new SqlConnection(connectionString);
            cnn1.Open();
            SqlCommand cmd1 = new SqlCommand("insert into kks_log (datum,korisnik,idprijave,opis) values  ( getdate(),'" + korisnik + "','" + idprijave + "','Izlaz na x')", cnn1);
            SqlDataReader reader1 = cmd1.ExecuteReader();
            cnn1.Close();            
        }
        
        // pregled djaltanika sa RFID

        private void pregledDjelatnikaSaRFIDToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OcistiPanele();
            //int firmaid = comb_poduzece.SelectedIndex;
            //string mtid = combo_odred.SelectedValue.ToString();
            pnl_naodlasku.Visible = true;
            DateTime danas = DateTime.Now;
            int m1 = danas.Month - 1;
            int y1 = danas.Year;
            if (m1 == 0)
            {
                m1 = 12;
                y1 = danas.Year - 1;
            }

            DateTime dat1 = new DateTime(y1, m1, 1);
            label130.Text = "Pregled djelatnika sa RFID";

            string dat10 = dat1.Year.ToString() + '-' + dat1.Month.ToString() + '-' + dat1.Day.ToString();
            string dat20 = danas.Year.ToString() + '-' + danas.Month.ToString() + '-' + danas.Day.ToString();

            string sql1 = "select id,ime PrezimeIme,DatumZaposlenja,addateend DatumOdlaska,MT,Lokacija,RadnoMjesto,poduzece Poduzeće,acnumber RFID,acactive AktivnaKartica,adtimeins DatumUnosaKartice,adtimechg DatumPromjene from( select acregno id, p.acworker ime, j.adDate datumzaposlenja, j.addateend, j.acCostDrv MT, j.acdept lokacija, j.acjob RadnoMjesto, 'AT' poduzece, d.acactive, d.acnumber, d.adtimeins, d.adtimechg from PantheonFxAt.dbo.thr_prsn p left join .PantheonFxAt.dbo.thr_prsnjob j on p.acworker = j.acworker left join PantheonFxAt.dbo.thr_prsnadddoc d on d.acworker = P.acworker and d.actype = 8 where(j.adDateEnd is null ) and rtrim(p.acregno) not  in ('', '0000') union all " +
                 " select acregno id, p2.acworker ime, j2.adDate datumzaposlenja, j2.addateend, j2.acCostDrv mt, j2.acdept lokacija, j2.acjob RadnoMjesto, 'TKB' poduzece, d.acactive, d.acnumber, d.adtimeins, d.adtimechg from PantheonTKB.dbo.thr_prsn p2  left join PantheonTKB.dbo.thr_prsnjob j2 on p2.acworker = j2.acworker left join PantheonTKB.dbo.thr_prsnadddoc d on p2.acworker = d.acworker and d.actype = 8 where(j2.adDateEnd is null ) and rtrim(p2.acregno) not  in ('', '0000') )y1 order by id ";

            string connecStrin = connectionStringpa;

            //if (firmaid == 1)
            {
                //    connecStrin = connectionStringpb;
            }

            SqlConnection connection = new SqlConnection(connecStrin);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            dgv_naodlasku.DataSource = ds;
            dgv_naodlasku.DataMember = "event";
            dgv_naodlasku.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

        }
// export u excell

        private void btn_export_Click(object sender, EventArgs e)
        {
            //Creating DataTable
            DataTable dt = new DataTable();

            //Adding the Columns
            foreach (DataGridViewColumn column in dgv_naodlasku.Columns)
            {

                // satnica se pokazuje samo za grupu 8 ili tehn. direktora (14)
                dt.Columns.Add(column.HeaderText, column.ValueType);


            }

            //Adding the Rows
            foreach (DataGridViewRow row in dgv_naodlasku.Rows)
            {
                dt.Rows.Add();

                foreach (DataGridViewCell cell in row.Cells)
                {


                    if (!(cell.Value is DBNull))
                    {
                        if (cell.Value != null)

                        {
                            dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();

                        }
                    }
                }
            }

            //if ((idloged == "8") || (idadm == "14"))
            //{

            //}
            //else
            //{
            //    // satnica se pokazuje samo za grupu 8 ili tehn. direktora (14)
            //    dt.Columns.RemoveAt(9);
            //    dt.Columns.RemoveAt(9);
            //    dt.Columns.RemoveAt(9);
            //    dt.Columns.RemoveAt(9);
            //}
            //Exporting to Excel
            string folderPath = "C:\\KKS\\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "RFID");
                wb.SaveAs(folderPath + "RFID.xlsx");
            }
            btn_exportRFID.Text = "Export done";

            FileInfo fi = new FileInfo("C:\\kks\\RFID.xlsx");
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(@"C:\kks\rfid.xlsx");
            }
            else
            {
                //file doesn't exist
            }

            //Object oMissing = System.Reflection.Missing.Value;
            //oMissing = System.Reflection.Missing.Value;
            //Object oTrue = true;
            //Object oFalse = false;
            //Microsoft.Office.Interop.Word.Application oExcel = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Document oExcelDoc = new Microsoft.Office.Interop.Excel.Document();
            //oExcel.Visible = true;

            //Object oTemplatePath = "c:\\kks\\kompetencije.xlsx";
            //oExcelDoc = oExcelDoc.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);





        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {

        }

        private void porukaNaTVToolStripMenuItem_Click(object sender, EventArgs e)
        {
            

        }

        private void tVPorukeToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OcistiPanele();
            pl_tvporuke.Visible = true;

        }

        private void button11_Click(object sender, EventArgs e)
        {
            using (SqlConnection cn = new SqlConnection(connectionString))
            {

                cn.Open();
                if (chbox_tv1.Checked)
                {                    
                    string sql1 = "update tv_poruka set status='1',poruka='" + tv_poruka1.Text.TrimEnd() + "' where tip='1' ";
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                    SqlDataReader reader = sqlCommand.ExecuteReader();                    
                }
                else
                {                    
                    string sql1 = "update tv_poruka set status='0' where tip='1' ";
                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                    SqlDataReader reader = sqlCommand.ExecuteReader();                  
                }
                cn.Close();
                cn.Open();
                if (chbox_tv2.Checked)
                {                    
                    string sql1 = "update tv_poruka set status='1',poruka='" + tv_poruka2.Text.TrimEnd() + "' where tip='2' ";
                    SqlCommand sqlCommand2 = new SqlCommand(sql1, cn);
                    SqlDataReader reader2 = sqlCommand2.ExecuteReader();                    
                }
                else
                {                    
                    string sql1 = "update tv_poruka set status='0' where tip='2' ";
                    SqlCommand sqlCommand2 = new SqlCommand(sql1, cn);
                    SqlDataReader reader2= sqlCommand2.ExecuteReader();
                }
                cn.Close();
                cn.Open();
                if (chbox_tv3.Checked)
                {                    
                    string sql1 = "update tv_poruka set status='1',poruka='" + tv_poruka3.Text.TrimEnd() + "' where tip='3' ";
                    SqlCommand sqlCommand3 = new SqlCommand(sql1, cn);
                    SqlDataReader reader3 = sqlCommand3.ExecuteReader();
                }
                else
                {                    
                    string sql1 = "update tv_poruka set status='0' where tip='3' ";
                    SqlCommand sqlCommand3 = new SqlCommand(sql1, cn);
                    SqlDataReader reader3 = sqlCommand3.ExecuteReader();
                }
                
                cn.Close();
            }

        }

        private void razno1ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            
        }

        private void raznoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process pProcess = new System.Diagnostics.Process();
            pProcess.StartInfo.FileName = @"C:\utils\zvono\zvonolan.exe";
            pProcess.StartInfo.Arguments = "192.168.20.6 Pocetak"; //argument
            pProcess.StartInfo.UseShellExecute = false;
            pProcess.StartInfo.RedirectStandardOutput = true;
            pProcess.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
            pProcess.StartInfo.CreateNoWindow = true; //not diplay a windows
            pProcess.Start();
            string output = pProcess.StandardOutput.ReadToEnd(); //The output result
            pProcess.WaitForExit();
        }

        private void bt_poruka_Click(object sender, EventArgs e)
        {
            
            pnl_poruka.Visible = true;
        }

        private void btn_posalji_Click(object sender, EventArgs e)
        {
            SqlConnection cnn1 = new SqlConnection(connectionString);
            cnn1.Open();
            string korisnikK = LoginForm.korisnik.ToString().Substring(0,LoginForm.korisnik.Length);   // trenutni korisnik
            SqlCommand cmd1 = new SqlCommand("insert into poruke (userid,datumu,sadržaj,author) values  ( '"+idradnika1.ToString()+"',getdate(),'"+textBox1.Text.TrimEnd() +"','"+korisnikK+"')", cnn1);
            SqlDataReader reader1 = cmd1.ExecuteReader();
            cnn1.Close();
            pnl_poruka.Visible = false;
        }

        private void pregledRadaSaoznakamaZaPlansatiradaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OcistiPanele();
            //int firmaid = comb_poduzece.SelectedIndex;
            //string mtid = combo_odred.SelectedValue.ToString();
            pnl_naodlasku.Visible = true;
            DateTime danas = DateTime.Now;
            int m1 = danas.Month - 1;
            int y1 = danas.Year;
            if (m1 == 0)
            {
                m1 = 12;
                y1 = danas.Year - 1;
            }

            DateTime dat1 = new DateTime(y1, m1, 1);
            label130.Text = "Pregled rada djelatnika za jučer";

            string dat10 = dat1.Year.ToString() + '-' + dat1.Month.ToString() + '-' + dat1.Day.ToString();
            string dat20 = danas.AddDays(-1).Year.ToString() + '-' + danas.AddDays(-1).Month.ToString() + '-' + danas.AddDays(-1).Day.ToString();

            string sql1 = "select k.id,k.prezimeime,k.funkcija,k.projekt,k.hala,k.linija,k.mjesto_troska,k.vještine,k.školovanje_posto,k.radnomjesto,k.datumzaposlenja,k.istek_ugovora,k.staz,j.danodlaska,k.godisnji_ostalo from kompetencije k left join ( select * from fxsap.dbo.plansatirada where danodlaska != '' and danodlaska<(dateadd(day,30,getdate())) and danodlaska> (dateadd(day, -30, getdate()))  and danodlaska is not NULL" +
                          ") j on j.radnikid = k.id where projekt not in ('xNERADI', 'xNa odlasku') and j.danodlaska is not null order by PrezimeIme desc ";

            sql1 = "select x1.acregno ID, x1.acworker PrezimeIme, x1.addateenter DatumDolaska, x1.addateend DatumOdlaska, x1.acjob RadnoMjesto, x1.acdept Lokacija, x1.accostdrv MjestoTroška,x1.BrojDanaGodisnjeg, x1.poduzece Poduzeće from( select p.acregno, p.acworker, p.addateenter, j.addateend, j.acjob, j.acdept, j.accostdrv,v.anvacationf1 BrojDanaGodisnjeg, 'FX' poduzece from PantheonFxAt.dbo.thr_prsnjob j left join PantheonFxAt.dbo.thr_prsn p on p.acworker = j.acworker left join PantheonFxAt.dbo.thr_prsnvacation v on p.acworker = v.acworker where j.addateend between '" + dat10 + "' and '" + dat20 + "'  " +
                  "union all select p.acregno, p.acworker, p.addateenter, j.addateend, j.acjob, j.acdept, j.accostdrv,v.anvacationf1 BrojDanaGodisnjeg, 'TKB' poduzece from PantheonTKB.dbo.thr_prsnjob j  left join PantheonTKB.dbo.thr_prsn p on p.acworker = j.acworker left join PantheonTKB.dbo.thr_prsnvacation v on p.acworker = v.acworker   where j.addateend between '" + dat10 + "' and '" + dat20 + "' ) x1 order by x1.acworker ";
            sql1 = "rfind.dbo.satniceoznake1 '" + dat20 + "'" ;

            string connecStrin = connectionString;

            //if (firmaid == 1)
            {
                //    connecStrin = connectionStringpb;
            }

            SqlConnection connection = new SqlConnection(connecStrin);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            dgv_naodlasku.DataSource = ds;
            dgv_naodlasku.DataMember = "event";
            dgv_naodlasku.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
        }

        private void pregledIskorištenihGodišnjihToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OcistiPanele();
            //int firmaid = comb_poduzece.SelectedIndex;
            //string mtid = combo_odred.SelectedValue.ToString();
            pnl_naodlasku.Visible = true;
            DateTime danas = DateTime.Now;
            int m1 = danas.Month - 1;
            int y1 = danas.Year;
            if (m1 == 0)
            {
                m1 = 12;
                y1 = danas.Year - 1;
            }

            DateTime dat1 = new DateTime(y1, m1, 1);
            label130.Text = "Pregled iskorištenih godišnjih odmora";
            using (SqlConnection cnn1 = new SqlConnection(connectionString))
            {

                var dataSource = new List<mjesto_troska>();

                dataSource.Add(new mjesto_troska() { naziv = " - ", id = "0" });

                foreach (var mt1 in lista_mt)
                {
                    dataSource.Add(new mjesto_troska() { naziv = mt1.naziv, id = mt1.id });
                }

                combo_mt.MaxDropDownItems = 60;
                this.combo_mt.DataSource = dataSource;
                this.combo_mt.DisplayMember = "Naziv";
                this.combo_mt.ValueMember = "id";
            }



        }

        private void btn_godisnji_Click(object sender, EventArgs e)
        {
            OcistiPanele();
            //int firmaid = comb_poduzece.SelectedIndex;
            //string mtid = combo_odred.SelectedValue.ToString();
            pnl_naodlasku.Visible = true;
            DateTime danas = DateTime.Now;
            int m1 = danas.Month - 1;
            int y1 = danas.Year;
            if (m1 == 0)
            {
                m1 = 12;
                y1 = danas.Year - 1;
            }
            string mt1 = combo_mt.SelectedValue.ToString();
            string godina1 = ((int.Parse)(combo_godina.SelectedIndex.ToString())+2017).ToString();
            string mjesec = txbox_mjesec.Text.ToString();
            DateTime dat1 = new DateTime(y1, m1, 1);
            label130.Text = "Pregled iskorištenih godišnjih odmora";
            string sql1 = "select ime,sum(brojdanagodisnjeg) BrojDanaGodišnjeg from(select ime, mjesec, (LEN(dani) - LEN(REPLACE(dani, 'g', ''))) brojdanagodisnjeg from( select ime, mjesec, radnikid, (isnull(dan01, '') + isnull(dan02, '') + isnull(dan03, '') + isnull(dan04, '') + isnull(dan05, '') + isnull(dan06, '') + isnull(dan07, '') + isnull(dan08, '') + isnull(dan09, '') + isnull(dan10, '') + isnull(dan11, '') + isnull(dan12, '') + isnull(dan13, '') + isnull(dan14, '') + isnull(dan15, '') + isnull(dan16, '') + isnull(dan17, '') + isnull(dan18, '') + isnull(dan19, '') + isnull(dan20, '') + isnull(dan21, '') + isnull(dan22, '') + isnull(dan23, '') + isnull(dan24, '') + isnull(dan25, '') + isnull(dan26, '') + isnull(dan27, '') + isnull(dan28, '') + isnull(dan29, '') + isnull(dan30, '') + isnull(dan31, '')) dani from fxsap.dbo.plansatirada where mt ="+mt1+" and godina = "+godina1+" and mjesec >= "+mjesec+" and (dan01 like '%g'  or dan02 like '%g'  or dan03 like '%g'  or dan04 like '%g'  or dan05 like '%g'  or dan06 like '%g' or dan07 like '%g'  or dan08 like '%g' or dan09 like '%g'  or dan10 like '%g'  or dan11 like '%g'  or dan12 like '%g'  or dan13 like '%g' or dan14 like '%g'  or dan15 like '%g' or dan16 like '%g' or dan17 like '%g' or dan18 like '%g' or dan19 like '%g' or dan20 like '%g' or dan21 like '%g' or dan22 like '%g' or dan23 like '%g' or dan24 like '%g' or dan25 like '%g' or dan26 like '%g' or dan27 like '%g' or dan28 like '%g'  or dan29 like '%g'  or dan30 like '%g' or dan31 like '%g'  ) " +
                           " ) x1 ) x2 group by ime order by ime ";

            string connecStrin = connectionStringFeroapp;

            //if (firmaid == 1)
            {
                //    connecStrin = connectionStringpb;
            }

            SqlConnection connection = new SqlConnection(connecStrin);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            dgv_naodlasku.DataSource = ds;
            dgv_naodlasku.DataMember = "event";
            dgv_naodlasku.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);


        }

        private void pnl_aktivnost_Paint(object sender, PaintEventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            pnl_psr_izmjena.Visible = false;
        }

        private void dgv_psr_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }


        // click na tablicu sa normama
        private void Dgv_rfid_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // klik sa npoamenama radi sadmo u 1. koloni

            int row1 = e.RowIndex;
            int col1 = e.ColumnIndex;

            if (col1 == 0)
            {
                pnl_napomene.Visible = true;

                DataGridViewRow selectedRow = dgv_psr.Rows[row1];
                //            ssindex1 
                g1 = (int.Parse)(Convert.ToString(selectedRow.Cells["Godina"].Value));
                m1 = (int.Parse)(Convert.ToString(selectedRow.Cells["Mjesec"].Value));
                //int d1 = col1 - 7;
                //DateTime dat1 = new DateTime(g1, m1, d1);

                string sql11 = "select n.DanNapomene,n.napomena from fxsap.dbo.plansatirada p left join fxsap.dbo.plansatiradanapomene n on p.psrid=n.psrid where year(n.dannapomene)=" + g1.ToString() + " and month(n.dannapomene)=" + m1.ToString() + " and radnikid=" + ssindex1.ToString();


                SqlConnection connection = new SqlConnection(connectionString);
                SqlDataAdapter dataadapter = new SqlDataAdapter(sql11, connection);
                DataSet ds = new DataSet();
                connection.Open();
                dataadapter.Fill(ds, "event3");
                connection.Close();

                dgv_napomene.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dgv_napomene.DataSource = ds;
                dgv_napomene.DataMember = "event3";

                dgv_napomene.AutoResizeColumns();
                dgv_napomene.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

            }
            if ((col1 > 7) && (idloged == "8"))
            {
                pnl_psr_izmjena.Visible = true;

                DataGridViewRow selectedRow = dgv_psr.Rows[row1];
                //            ssindex1 
                g1 = (int.Parse)(Convert.ToString(selectedRow.Cells["Godina"].Value));
                m1 = (int.Parse)(Convert.ToString(selectedRow.Cells["Mjesec"].Value));
                //string danx;
                if ((col1 - 7) < 10)
                {
                    danx = "dan0" + (col1 - 7).ToString();
                }
                else
                {
                    danx = "dan" + (col1 - 7).ToString();
                }



                //int d1 = col1 - 7;
                //DateTime dat1 = new DateTime(g1, m1, d1);

                string sql11 = "select " + danx + " from fxsap.dbo.plansatirada p where godina=" + g1.ToString() + " and mjesec=" + m1.ToString() + " and radnikid=" + ssindex1.ToString();
                using (SqlConnection cn = new SqlConnection(connectionString))
                {
                    cn.Open();
                    string sql1 = "select " + danx + " from fxsap.dbo.plansatirada p where godina=" + g1.ToString() + " and mjesec=" + m1.ToString() + " and radnikid=" + ssindex1.ToString();

                    SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                    SqlDataReader reader = sqlCommand.ExecuteReader();
                    string oldvalue = "";
                    while (reader.Read())
                    {
                        oldvalue = reader[danx].ToString();
                    }
                    cn.Close();
                    txtbox_sv.Text = oldvalue;
                    txtbox_nv.Text = oldvalue;
                }




            }

        }
// tablica sa normama, click otvara aktivnosti
        private void Dgv_norme_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // klik sa napomenama radi sadmo u 1. koloni

            int row1 = e.RowIndex;
            int col1 = e.ColumnIndex;

            if (col1 > 3)
            {

                pnl_napomene.Visible = true;

                DataGridViewRow selectedRow = dgv_norme.Rows[row1];
                //            ssindex1 
                string sdat1 = (Convert.ToString(selectedRow.Cells["Datum"].Value));
                string slinija = (Convert.ToString(selectedRow.Cells["Linija"].Value));
                string shala = (Convert.ToString(selectedRow.Cells["Hala"].Value));
                string vrijemeod = (Convert.ToString(selectedRow.Cells["VrijemeOd"].Value));
                string vrijemedo = (Convert.ToString(selectedRow.Cells["VrijemeDo"].Value));
                DateTime dat1 = DateTime.Parse(sdat1);
                DateTime datp = DateTime.Parse(vrijemeod);
                DateTime datk = DateTime.Parse(vrijemedo);
                DateTime dat2 = DateTime.Parse(sdat1);
                if (datk < datp)
                {
                    dat2 = DateTime.Parse(sdat1).AddDays(1);
                }

                string sdatp = dat1.Year.ToString() + '-' + dat1.Month.ToString() + '-' + dat1.Day.ToString() + " " + datp.Hour.ToString() + ":" + datp.Minute.ToString() + ":00";
                string sdatk = dat2.Year.ToString() + '-' + dat2.Month.ToString() + '-' + dat2.Day.ToString() + " " + datk.Hour.ToString() + ":" + datk.Minute.ToString() + ":00";

                int z = 0;
                //if ( 1 ==2 )
                //{ 

                //    m1 = (int.Parse)(Convert.ToString(selectedRow.Cells["Mjesec"].Value));
                //    //int d1 = col1 - 7;
                //    //DateTime dat1 = new DateTime(g1, m1, d1);

                string sqla = "select * from rfind.dbo.pregled_po_liniji( '" + sdatp + "','" + sdatk + "','" + shala + "','" + slinija + "')";

                pnl_aktivnosti.Visible = true;

                SqlConnection connection = new SqlConnection(connectionString);
                SqlDataAdapter dataadapter = new SqlDataAdapter(sqla, connection);
                DataSet ds = new DataSet();
                connection.Open();
                dataadapter.Fill(ds, "event_akt");
                connection.Close();

                dgv_aktivnosti.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
                dgv_aktivnosti.DataSource = ds;
                dgv_aktivnosti.DataMember = "event_akt";
                dgv_aktivnosti.AutoResizeColumns();
                dgv_aktivnosti.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
            }
            }

        private void Button14_Click_1(object sender, EventArgs e)
        {
            pnl_aktivnosti.Visible = false;
        }

        //  popis školovanja
        private void ComboBox1_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            int ind1 = combox_skolovanja.SelectedIndex;
            string idradnika1 = cbx_vj_djelatnici.SelectedValue.ToString();
            

            if (ind1 <= 0)
                return;

            var o10 = combox_skolovanja.SelectedItem;
            skolovanje1 sk10 = (skolovanje1)(o10);

            string ids = sk10.idskolovanje.ToString();

            int idskolovanja = 0;

            string sql11 = "select * from skolovanje where idskolovanja=" + ids + " and idradnika= " + idradnika1 + " order by izmjenaod";
            //string sql11 = "select * from skolovanje1 where idradnika= " + idradnika1 + " order by izmjenaod";
            SqlConnection cn2 = new SqlConnection(connectionString);
            cn2.Open();
            SqlCommand sqlcommand1 = new SqlCommand(sql11, cn2);
            SqlDataReader reader1;
            reader1 = sqlcommand1.ExecuteReader();
            lista_skolovanja.Clear();

            while (reader1.Read())
            {
                pnl_škola.Visible = true;
                checkBox14.Checked = true;
                textBox4.Text = reader1["mentor"].ToString();
                idskolovanja = (int.Parse)(reader1["idskolovanja"].ToString());
                tb_s_napomena.Text = reader1["napomena"].ToString();

                skolovanje1 sk1 = new skolovanje1();
                sk1.idskolovanje = idskolovanja.ToString();
                sk1.opis = tb_s_napomena.Text;

                lista_skolovanja.Add(sk1);

                string dat0 = reader1["oddatuma"].ToString();
                DateTime dt = Convert.ToDateTime(dat0);
                //DateTime dat1 = new DateTime(2012, 05, 28);
                dateTimePicker2.Value = dt;
                dat0 = reader1["dodatuma"].ToString();
                dt = Convert.ToDateTime(dat0);
                dateTimePicker3.Value = dt;
            }

            cbxlista_skolovanje.Visible = false;
            cbxlista_skolovanje.Items.Clear();

            cbx_skolo_marked.Visible = false;
            cboxlistaskol.Clear();


            if ((pnl_škola.Visible) && (idloged == "5"))
            {
                cbxlista_skolovanje.Visible = true;
                //cbx_skolo_marked.Visible = true;
                string vje1 = "";
                
                    using (SqlConnection cn = new SqlConnection(connectionString))
                    {
                        cn.Open();
                        SqlCommand sqlCommand = new SqlCommand("SELECT s.*,v.naziv from radnicivjestines s left join vjestine v on v.id=s.idvjestine where idskolovanja=" + ids+ " and idradnika="+idradnika1, cn);
                        SqlDataReader reader = sqlCommand.ExecuteReader();


                        while (reader.Read())
                        {
                            // reader["Datum"].ToString();
                            CheckboxItem cbi1 = new CheckboxItem();

                            cbi1.Value = reader["idvjestine"].ToString();
                            cbi1.Text = reader["naziv"].ToString();

                            if (reader["Vrijednost"]==DBNull.Value)
                            {
                                    
                            }
                            else
                            {
                                if  (reader["Vrijednost"].ToString()=="1")
                                {
                                    cbi1.checked1 = true;
                                }
                                
                            }
                        var cb11 = cbi1;
                        cboxlistaskol.Add(cbi1);
                        

                    }
                    }                   

                cbxlista_skolovanje.Items.Clear();
                cbxlista_skolovanje.Items.AddRange(cboxlistaskol.ToArray());
                
                for( int z1=0;z1<cbxlista_skolovanje.Items.Count;z1++)
                {
                    var o1 = cbxlista_skolovanje.Items[z1];
                    CheckboxItem cb1 = (CheckboxItem)(o1);
                    if (cb1.checked1)
                          cbxlista_skolovanje.SetItemChecked(z1, true);
                }

            }
            



        }

        private void Cbxlista_skolovanje_SelectedIndexChanged(object sender, EventArgs e)
        {
           

        }

        private void Cbxlista_skolovanje_ItemCheck(object sender, ItemCheckEventArgs e)
        {

            //int i1 = cbxlista_skolovanje.SelectedIndex;
            //if (cbxlista_skolovanje.GetItemCheckState(i1) == CheckState.Checked)
            //{
            //    cbxlista_skolovanje.SetItemCheckState(i1, CheckState.Unchecked);
            //}
            //else
            //{
            //    cbxlista_skolovanje.SetItemCheckState(i1, CheckState.Checked);
            //}
        }

        private void button14_Click(object sender, EventArgs e)
        {
            if (prikaz1 == 1)
            {
                pnl_norme.Height = 1000;
                pnl_norme.Width = 1900;
                dgv_norme.Width = 1800;
                dgv_norme.Height = 900;
                pnl_rfid.Visible = false;
                panel2.Visible = false;
                prikaz1 = 2;
                btn_povecajprikaz.Text = "Smanji prikaz";
            }
            else
            {
                prikaz1 = 1;
                pnl_norme.Width = 2337;
                pnl_norme.Height = 247;
                dgv_norme.Width = 1300;
                dgv_norme.Height = 205;
                pnl_rfid.Visible = false;
                panel2.Visible = true;
                dgv_norme.Refresh();
                combo_listadjelatnika.Refresh();
                btn_povecajprikaz.Text = "Povećaj prikaz";

            }


        }

        private void button12_Click(object sender, EventArgs e)
        {
            
            using (SqlConnection cn = new SqlConnection(connectionString))
            {
                cn.Open();
                string sql1 = "update fxsap.dbo.plansatirada set  " + danx + " = '" + txtbox_nv.Text.TrimEnd() + "'  where godina=" + g1.ToString() + " and mjesec=" + m1.ToString() + " and radnikid=" + ssindex1.ToString();

                SqlCommand sqlCommand = new SqlCommand(sql1, cn);
                SqlDataReader reader = sqlCommand.ExecuteReader();
            }
            pnl_psr_izmjena.Visible = false;
            this.dgv_psr.EndEdit();
            dgv_psr.Update();
            dgv_psr.Refresh();
            this.dgv_psr.Parent.Refresh();
    

            //var host = Dns.GetHostEntry(Dns.GetHostName());
            //foreach (var ip in host.AddressList)
            //{
            //    if (ip.AddressFamily == AddressFamily.InterNetwork)
            //    {
            //        ip= ip.ToString()
            //    }
            //}
            string host1=SystemInformation.ComputerName;


            using (SqlConnection cnn1 = new SqlConnection(connectionString))
            {
                cnn1.Open();
                SqlCommand cmd1 = new SqlCommand("insert into kks_log (datum,korisnik,idprijave,opis) values  ( getdate(),'" + korisnik + "','" + idprijave + "','"+host1+" - izmjena satnice za ssindex1=" + ssindex1 + " stara vrijednost za "+danx+" ="+ txtbox_sv.Text.TrimEnd() + " , nova vrijednost "+ txtbox_nv.Text.TrimEnd() + " '    )", cnn1);
                SqlDataReader reader1 = cmd1.ExecuteReader();
                cnn1.Close();
            }

        }

        private void pregledDjelatnikaOtišlihUZadnjihMjesecDanaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OcistiPanele();
            //int firmaid = comb_poduzece.SelectedIndex;
            //string mtid = combo_odred.SelectedValue.ToString();
            pnl_naodlasku.Visible = true;
            DateTime danas = DateTime.Now;
            int m1 = danas.Month-1;
            int y1 = danas.Year;
            if (m1 == 0)
            {
                m1 = 12;
                y1 = danas.Year - 1;
            }

            DateTime dat1 = new DateTime(y1, m1, 1);
            label130.Text = "Pregled djelatnika otišlih prethodni i ovaj mjesec";

            string dat10 = dat1.Year.ToString() + '-' + dat1.Month.ToString() + '-' + dat1.Day.ToString();
            string dat20 = danas.Year.ToString() + '-' + danas.Month.ToString() + '-' + danas.Day.ToString();

            string sql1 = "select k.id,k.prezimeime,k.funkcija,k.projekt,k.hala,k.linija,k.mjesto_troska,k.vještine,k.školovanje_posto,k.radnomjesto,k.datumzaposlenja,k.istek_ugovora,k.staz,j.danodlaska,k.godisnji_ostalo from kompetencije k left join ( select * from fxsap.dbo.plansatirada where danodlaska != '' and danodlaska<(dateadd(day,30,getdate())) and danodlaska> (dateadd(day, -30, getdate()))  and danodlaska is not NULL" +
                          ") j on j.radnikid = k.id where projekt not in ('xNERADI', 'xNa odlasku') and j.danodlaska is not null order by PrezimeIme desc ";

            sql1 = "select x1.acregno ID, x1.acworker PrezimeIme, x1.addateenter DatumDolaska, x1.addateend DatumOdlaska, x1.acjob RadnoMjesto, x1.acdept Lokacija, x1.accostdrv MjestoTroška,x1.BrojDanaGodisnjeg, x1.poduzece Poduzeće from( select p.acregno, p.acworker, p.addateenter, j.addateend, j.acjob, j.acdept, j.accostdrv,v.anvacationf1 BrojDanaGodisnjeg, 'FX' poduzece from PantheonFxAt.dbo.thr_prsnjob j left join PantheonFxAt.dbo.thr_prsn p on p.acworker = j.acworker left join PantheonFxAt.dbo.thr_prsnvacation v on p.acworker = v.acworker where j.addateend between '" + dat10+"' and '"+dat20+ "'  " +
                "and p.acworker not in ( select acworker from PantheonFxAt.dbo.thr_prsnjob where addateend is null  )  " +
                  "union all select p.acregno, p.acworker, p.addateenter, j.addateend, j.acjob, j.acdept, j.accostdrv,v.anvacationf1 BrojDanaGodisnjeg, 'TKB' poduzece from PantheonTKB.dbo.thr_prsnjob j  left join PantheonTKB.dbo.thr_prsn p on p.acworker = j.acworker left join PantheonTKB.dbo.thr_prsnvacation v on p.acworker = v.acworker   where j.addateend between '" + dat10+"' and '"+dat20+ "' " +
                  "and p.acworker not in (select acworker from PantheonTKB.dbo.thr_prsnjob where addateend is null  )  ) x1 order by x1.acworker ";

            
          string connecStrin = connectionStringpa;

            //if (firmaid == 1)
            {
                //    connecStrin = connectionStringpb;
            }

            SqlConnection connection = new SqlConnection(connecStrin);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            dgv_naodlasku.DataSource = ds;
            dgv_naodlasku.DataMember = "event";
            dgv_naodlasku.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);
        }

        private void checkBox15_CheckedChanged(object sender, EventArgs e)
        {
            OcistiPanele();

            pnl_novi_djelatnici.Visible = true;
            string sql1 = "";
            if (idloged == "8")
            {
                sql1 = "select cast(acregno as int ) as Id,acname Ime,acsurname Prezime,j.acCostDrv Mjesto_troška,j.acDept Lokacija,d.acnumber RFID,j.acjob Radno_mjesto,'' Radni_staz,p.adDateEnter DatumZaposlenja,'' sifra_rm,p.acstreet Ulica,acpost Pošta,'' vrijeme,accity Grad,j.acFieldSA Vrsta_isplate, adDateExit Datum_odlaska from thr_prsn p left join thr_prsnjob j on p.acworker = j.acworker left join thr_prsnadddoc d on d.acWorker = p.acworker and d.actype = 8 where d.actype=8 and j.addateend is null " +
                "order by cast(acregno as int) desc ";
            }
            else
            {
                sql1 = "select cast(acregno as int ) as Id,acname Ime,acsurname Prezime,j.acCostDrv Mjesto_troška,j.acDept Lokacija,d.acnumber RFID,j.acjob Radno_mjesto,'' Radni_staz,p.adDateEnter DatumZaposlenja,'' sifra_rm,p.acstreet Ulica,acpost Pošta,'' vrijeme,accity Grad,j.acFieldSA Vrsta_isplate, adDateExit Datum_odlaska from thr_prsn p left join thr_prsnjob j on p.acworker = j.acworker left join thr_prsnadddoc d on d.acWorker = p.acworker and d.actype = 8 " +
                "where d.actype=8 and j.dateend is null and p.addateenter>=dateadd(d,-3,getdate()) order by cast(acregno as int) desc ";
            }
            
            string connecStrin = connectionStringpa;
            if (chbx_Feroimpex.Checked)
            {
                connecStrin = connectionStringpa;
                chbx_TKB.Checked = false;
            }
            if (chbx_TKB.Checked)
            {
                connecStrin = connectionStringpb;
                chbx_Feroimpex.Checked = false;
            }
            //if (firmaid == 1)
            {
                //    connecStrin = connectionStringpb;
            }

            SqlConnection connection = new SqlConnection(connecStrin);
            SqlDataAdapter dataadapter = new SqlDataAdapter(sql1, connection);
            DataSet ds = new DataSet();
            connection.Open();
            dataadapter.Fill(ds, "event");
            connection.Close();

            // dgv_odred.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.Fill);
            dgv_novi_djelatnici.DataSource = ds;
            dgv_novi_djelatnici.DataMember = "event";
            dgv_novi_djelatnici.AutoResizeColumns(DataGridViewAutoSizeColumnsMode.AllCells);

        }

        private void btn_v_b_cancel_Click(object sender, EventArgs e)
        {
            
            pnl_v_dialogbrisi.Visible = false;
        }
        
        // mjesečni izvještaj bol./godisn/izostanci  export to excell
        private void btn_a_export_Click(object sender, EventArgs e)
        {
            
            //Creating DataTable
            DataTable dt = new DataTable();

            //Adding the Columns
            foreach (DataGridViewColumn column in dgv_aktivnost.Columns)
            {
                dt.Columns.Add(column.HeaderText, column.ValueType);
            }

            //Adding the Rows
            foreach (DataGridViewRow row in dgv_aktivnost.Rows)
            {
                dt.Rows.Add();
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.Value != null && cell.Value != DBNull.Value)
                    {
                        dt.Rows[dt.Rows.Count - 1][cell.ColumnIndex] = cell.Value.ToString();
                    }
                }
            }

            //Exporting to Excel
            string folderPath = "C:\\KKS\\";
            if (!Directory.Exists(folderPath))
            {
                Directory.CreateDirectory(folderPath);
            }
            using (XLWorkbook wb = new XLWorkbook())
            {
                wb.Worksheets.Add(dt, "Djelatnici");
                wb.SaveAs(folderPath + "MjesečnaAktivnost.xlsx");
            }
            button2.Text = "Export done";


            FileInfo fi = new FileInfo("C:\\kks\\MjesečnaAktivnost.xlsx");
            if (fi.Exists)
            {
                System.Diagnostics.Process.Start(@"C:\\kks\\MjesečnaAktivnost.xlsx");
            }
            else
            {
                //file doesn't exist
            }


        }

        private void btn_a_prebacaj_Click(object sender, EventArgs e)
        {
            prikaziizostanak = 0;
            prikazinormu = 1;
            prikazigodisnji = 0;
            prikazibolovanje = 0;
            lbl_izabrano.Text = "Prikaz izvršenja norme";
        }
#endregion

#region export to word, print
        private void bt_print1_Click(object sender, EventArgs e)
        {

            Object oMissing = System.Reflection.Missing.Value;
            oMissing = System.Reflection.Missing.Value;
            Object oTrue = true;
            Object oFalse = false;
            Microsoft.Office.Interop.Word.Application oWord = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Document oWordDoc = new Microsoft.Office.Interop.Word.Document();
            oWord.Visible = true;
         
            Object oTemplatePath = "c:\\kks\\kompetencijetmpl7.dotx";
            oWordDoc = oWord.Documents.Add(ref oTemplatePath, ref oMissing, ref oMissing, ref oMissing);

            foreach (Microsoft.Office.Interop.Word.Field myMergeField in oWordDoc.Fields)
            {
                Microsoft.Office.Interop.Word.Range rngFieldCode = myMergeField.Code;
                String fieldText = rngFieldCode.Text;
                if (fieldText.StartsWith(" MERGEFIELD"))
                {
                    //Int32 endMerge = fieldText.IndexOf("\\");
                    //Int32 fieldNameLength = fieldText.Length - endMerge;
                    //String fieldName = fieldText.Substring(11, endMerge - 11);
                    string[] arr1 = fieldText.Split(' ');
                    //fieldName = fieldName.Trim();
                    string fieldName = arr1[2];
                    // foreach (var item in pDictionaryMerge)
                    // {


                    myMergeField.Select();

                    switch (fieldName)
                    {
                        case "PrezimeIme":
                            myMergeField.Select();
                            oWord.Selection.TypeText(label24.Text);
                            break;
                        case "DatumZaposlenja":
                            oWord.Selection.TypeText(label25.Text);
                            break;
                        case "Funkcija":
                            oWord.Selection.TypeText(label26.Text);
                            break;
                        case "Mjesto_troska":
                            oWord.Selection.TypeText(label28.Text);
                            break;
                        case "RadnoMjesto":
                            oWord.Selection.TypeText(label29.Text);
                            break;
                        case "Vjestina":
                            if (tb_vjestina.Text.ToString().Length==0 )
                            {
                                oWord.Selection.TypeText("  ");
                            }
                            else
                                oWord.Selection.TypeText(tb_vjestina.Text);
                            
                            break;
                        case "Istek_Ugovora":

                            if (label30.Text.ToString().Length == 0)
                            {
                                oWord.Selection.TypeText("    ");                                
                            }
                            else
                                oWord.Selection.TypeText(label30.Text);


                            break;
                        case "Staz":
                            oWord.Selection.TypeText(label31.Text);
                            break;
                        case "SatnicaStara":
                            oWord.Selection.TypeText(label32.Text);
                            break;
                        case "SatnicaBruto":
                            oWord.Selection.TypeText(label33.Text);
                            break;
                        case "SatnicaNovaOd":
                            oWord.Selection.TypeText(label34.Text);
                            break;
                        case "DolasciNedjeljom1":
                            oWord.Selection.TypeText(label46.Text);
                            break;
                        case "DolasciNedjeljom3":
                            oWord.Selection.TypeText(label45.Text);
                            break;
                        case "DolasciNedjeljom6":
                            oWord.Selection.TypeText(label23.Text);
                            break;
                        case "DolasciNedjeljom12":
                            oWord.Selection.TypeText(label22.Text);
                            break;

                        case "NedolasciNedjeljom1":
                            oWord.Selection.TypeText(label35.Text);
                            break;
                        case "NedolasciNedjeljom3":
                            oWord.Selection.TypeText(label57.Text);
                            break;
                        case "NedolasciNedjeljom6":
                            oWord.Selection.TypeText(label68.Text);
                            break;
                        case "NedolasciNedjeljom12":
                            oWord.Selection.TypeText(label53.Text);
                            break;

                        case "NedolasciSubotom1":
                            oWord.Selection.TypeText(label36.Text);
                            break;
                        case "NedolasciSubotom3":
                            oWord.Selection.TypeText(label56.Text);
                            break;
                        case "NedolasciSubotom6":
                            oWord.Selection.TypeText(label67.Text);
                            break;
                        case "NedolasciSubotom12":
                            oWord.Selection.TypeText(label82.Text);
                            break;

                        case "Bolovanje_broj1":
                            oWord.Selection.TypeText(label37.Text);
                            break;
                        case "Bolovanje_broj3":
                            oWord.Selection.TypeText(label55.Text);
                            break;
                        case "Bolovanje_broj6":
                            oWord.Selection.TypeText(label66.Text);
                            break;
                        case "Bolovanje_broj12":
                            oWord.Selection.TypeText(label81.Text);
                            break;

                        case "Bolovanja_dana1":
                            oWord.Selection.TypeText(label38.Text);
                            break;
                        case "Bolovanja_dana3":
                            oWord.Selection.TypeText(label54.Text);
                            break;
                        case "Bolovanja_dana6":
                            oWord.Selection.TypeText(label65.Text);
                            break;
                        case "Bolovanja_dana12":
                            oWord.Selection.TypeText(label80.Text);
                            break;

                        case "Stimulacija1":
oWord.Selection.TypeText(label39.Text);
                            break;
                        case "Stimulacija3":
                            oWord.Selection.TypeText(label53.Text);
                            break;
                        case "Stimulacija6":
                            oWord.Selection.TypeText(label64.Text);
                            break;
                        case "Stimulacija12":
                            oWord.Selection.TypeText(label79.Text);
                            break;

                        case "NedostajeDo8sati1":
                            oWord.Selection.TypeText(label40.Text);
                            break;
                        case "NedostajeDo8sati3":
                            oWord.Selection.TypeText(label52.Text);
                            break;
                        case "NedostajeDo8sati6":
                            oWord.Selection.TypeText(label63.Text);
                            break;
                        case "NedostajeDo8sati12":
                            oWord.Selection.TypeText(label78.Text);
                            break;
                        case "Kasni1":
                            oWord.Selection.TypeText(label41.Text);
                            break;
                        case "Kasni3":
                            oWord.Selection.TypeText(label51.Text);
                            break;
                        case "Kasni6":
                            oWord.Selection.TypeText(label62.Text);
                            break;
                        case "Kasni12":
                            oWord.Selection.TypeText(label77.Text);
                            break;
                        case "PreranoOtisao1":
                            oWord.Selection.TypeText(label42.Text);
                            break;
                        case "PreranoOtisao3":
                            oWord.Selection.TypeText(label50.Text);
                            break;
                        case "PreranoOtisao6":
                            oWord.Selection.TypeText(label61.Text);
                            break;
                        case "PreranoOtisao12":
                            oWord.Selection.TypeText(label76.Text);
                            break;
                        case "NeopravdaniDani1":
                            oWord.Selection.TypeText(label43.Text);
                            break;
                        case "NeopravdaniDani3":
                            oWord.Selection.TypeText(label49.Text);
                            break;
                        case "NeopravdaniDani6":
                            oWord.Selection.TypeText(label60.Text);
                            break;
                        case "NeopravdaniDani12":
                            oWord.Selection.TypeText(label75.Text);
                            break;
                        case "NormaPosto1":
                            oWord.Selection.TypeText(label44.Text);
                            break;
                        case "NormaPosto3":
                            oWord.Selection.TypeText(label48.Text);
                            break;
                        case "NormaPosto6":
                            oWord.Selection.TypeText(label59.Text);
                            break;
                        case "NormaPosto12":
                            oWord.Selection.TypeText(label74.Text);
                            break;
                        case "Godisnji_dana1":
                            oWord.Selection.TypeText(label85.Text);
                            break;
                        case "Godisnji_dana3":
                            oWord.Selection.TypeText(label84.Text);
                            break;
                        case "Godisnji_dana6":
                            oWord.Selection.TypeText(label73.Text);
                            break;
                        case "Godisnji_dana12":
                            oWord.Selection.TypeText(label58.Text);
                            break;
                        default:
                            break;

                    }
                    // }

                }
            }
            string ssindex = combo_listadjelatnika.SelectedValue.ToString();
            oTemplatePath = "c:\\kks\\kompetencijetmp_"+ssindex+".doc";
            //oWordDoc.ExportAsFixedFormat(    oTemplatePath.Replace(".doc", ".pdf"),     oWordDoc.wdExportFormatPDF,     BitmapMissingFonts: true, DocStructureTags: false);

            oWordDoc.SaveAs(oTemplatePath);
            oWord.Documents.Open(oTemplatePath);
           // oWord.Application.Quit();

        }
#endregion


    }


    public static class ExtensionMethods
    {
        public static void DoubleBuffered(this DataGridView dgv, bool setting)
        {
            Type dgvType = dgv.GetType();
            PropertyInfo pi = dgvType.GetProperty("DoubleBuffered",
                BindingFlags.Instance | BindingFlags.NonPublic);
            pi.SetValue(dgv, setting, null);
        }
    }
}
