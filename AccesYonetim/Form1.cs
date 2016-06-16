using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace AccesYonetim
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        static OleDbConnection Baglan = new OleDbConnection();

        static string Veritabani;
        int SayfaSayisi = 0;
        
        static int DiziSayisi = 1;
        string[] Sayfalar = new string[DiziSayisi];
        private void buttonGozat_Click(object sender, EventArgs e)
        {
            OpenFileDialog Ac = new OpenFileDialog();
            Ac.Filter = " Acces(.mdb)|*.mdb| Acces(.accdb)| *.accdb ";
            Ac.ShowDialog();
            Ac.Title = "Bir Veritabanı Dosyası Seçiniz..";

            textBox1.Text = Ac.FileName;
            Veritabani = "provider=Microsoft.ACE.OLEDB.12.0;Data source=" + Ac.FileName;
            buttonbaglan.Enabled = true;


        }

        BindingSource Bs = new BindingSource();
        private void buttonbaglan_Click(object sender, EventArgs e)
        {
            if (Baglan.State==ConnectionState.Open)
            {
                Baglan.Close();
            }

            hangisi = "Acces";
            try
            {

                Baglan.ConnectionString = Veritabani;
                listBoxTbl.Items.Clear();
                DataTable Tablo = new DataTable();
                if (Baglan.State == ConnectionState.Closed)
                {
                    Baglan.Open();
                }

                DataTable TabloTable = Baglan.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });
                DataTable TabloSorgu = Baglan.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "VIEW" });


                foreach (DataRow item in TabloTable.Rows)
                {
                    listBoxTbl.Items.Add("T> " + item[2]);
                }
                foreach (DataRow item in TabloSorgu.Rows)
                {
                    listBoxTbl.Items.Add("S> " + item[2]);
                }

                

                buttonbaglanKes.Enabled = true;
                buttonYeniAcces.Enabled = true;
                buttonGorunumCalistir.Enabled = true;

            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message.ToString());
            }

        }


        private void buttonGorunumCalistir_Click(object sender, EventArgs e)
        {

            

            try
            {
                if (Baglan.State == ConnectionState.Closed)
                {
                    Baglan.Open();
                }
                

                if (richTextBox1.SelectedText != null)
                {
                    
                        DataTable STablo = new DataTable();
                        OleDbCommand SKomut = new OleDbCommand(richTextBox1.SelectedText, Baglan);
                        OleDbDataReader Oku = SKomut.ExecuteReader();
                        STablo.Load(Oku);
                        Bs.DataSource = STablo;
                        dataGridView1.DataSource = Bs;

                        toolStripLabelSayi.Text = dataGridView1.Rows.Count.ToString();

                        ToolStripTrue(true);
  
                }
              
            }
            catch (Exception)
            {


                try
                {
                    DataTable Tablo = new DataTable();
                    OleDbCommand Komut = new OleDbCommand(richTextBox1.Text, Baglan);

                    OleDbDataReader Oku = Komut.ExecuteReader();
                    Tablo.Load(Oku);
                    Bs.DataSource = Tablo;
                    dataGridView1.DataSource = Bs;

                    toolStripLabelSayi.Text = dataGridView1.Rows.Count.ToString();

                    ToolStripTrue(true);
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message.ToString());
                }
   
            }

        }

        private void ToolStripTrue(bool Dogrumu)
        {
            toolStripLabelSayi.Enabled = Dogrumu;
            toolStripButton3.Enabled = Dogrumu;
            toolStripButton4.Enabled = true;
            toolStripButton5.Enabled = true;
            toolStripButton6.Enabled = true;
            toolStripButton7.Enabled = true;
        }

        



        private void listBoxTablo_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            // richTextBox1.Text ="\n"+ listBoxTablo.SelectedItem.ToString();

            richTextBox1.SelectedText = listBoxTbl.SelectedItem.ToString();
        }


         static string Rich;
         
        private void richTextBox1_TextChanged(object sender, EventArgs e)
        {
            

            Sayfalar[tabControl2.SelectedIndex] = richTextBox1.Text;
            Rich = richTextBox1.Text; 
            MatchCollection MaviAra;
            MatchCollection KahveAra;
            MatchCollection YesilAra;
            MatchCollection KirmiziAra;
            MatchCollection PembeAra;
            Keywords(out MaviAra, out KahveAra, out YesilAra, out KirmiziAra, out PembeAra);

            

            
            int originalIndex = richTextBox1.SelectionStart;
            int originalLength = richTextBox1.SelectionLength;
            Color originalColor = Color.Black;


            richTextBox1.SelectionStart = 0;
            richTextBox1.SelectionLength = richTextBox1.Text.Length;
            richTextBox1.SelectionColor = originalColor;

            TextInfo Kultur = new CultureInfo("en-US", false).TextInfo;

            foreach (Match Mkahve in KahveAra)
            {
                richTextBox1.SelectionStart = Mkahve.Index;
                richTextBox1.SelectionLength = Mkahve.Length;
                richTextBox1.SelectionColor = Color.Brown;

                richTextBox1.SelectedText = Kultur.ToTitleCase(Mkahve.ToString());
            }

            foreach (Match Mpembe in PembeAra)
            {
                richTextBox1.SelectionStart = Mpembe.Index;
                richTextBox1.SelectionLength = Mpembe.Length;
                richTextBox1.SelectionColor = Color.Fuchsia;

                richTextBox1.SelectedText = Kultur.ToTitleCase(Mpembe.ToString());
    
            }
            

            foreach (Match Myesil in YesilAra)
            {
                richTextBox1.SelectionStart = Myesil.Index;
                richTextBox1.SelectionLength = Myesil.Length;
                richTextBox1.SelectionColor = Color.Green;

                richTextBox1.SelectedText = Kultur.ToTitleCase(Myesil.ToString());
            }

            foreach (Match Mkirmizi in KirmiziAra)
            {
                richTextBox1.SelectionStart = Mkirmizi.Index;
                richTextBox1.SelectionLength = Mkirmizi.Length;
                richTextBox1.SelectionColor = Color.Red;

                richTextBox1.SelectedText = Kultur.ToUpper(Mkirmizi.ToString());
            }

            

            foreach (Match Mara in MaviAra)
            {
                richTextBox1.SelectionStart = Mara.Index;
                richTextBox1.SelectionLength = Mara.Length;
                richTextBox1.SelectionColor = Color.Blue;

                richTextBox1.SelectedText = Kultur.ToTitleCase(Mara.ToString());
            }


            richTextBox1.SelectionStart = originalIndex;
            richTextBox1.SelectionLength = originalLength;
            richTextBox1.SelectionColor = originalColor;

            // giving back the focus
            richTextBox1.Focus();

        }

        private static void Keywords(out MatchCollection MaviAra, out MatchCollection KahveAra, out MatchCollection YesilAra, out MatchCollection KirmiziAra, out MatchCollection PembeAra)
        {
            string MaviSozluk = @"\b(iif|Iif|left|Left|right|Right|mid|Mid|len|Len|ucase|Ucase|lcase|Lcase|ltrim|Ltrim|rtrim|Rtrim|Trim|trim|round|Round|)\b";
            MaviAra = Regex.Matches(Rich.ToLower(new CultureInfo("en-US", false)), MaviSozluk);

            string KahveSozluk = @"\b(distinct|Distinct|as|As|set|Set|)\b";
            KahveAra = Regex.Matches(Rich.ToLower(new CultureInfo("en-US", false)), KahveSozluk);

            string YesilSozluk = @"\b(char|Char|varchar|Varchar|nvarchar|Nvarchar|int|Int|money|Money|smallint|Smallint|counter|Counter|primary|Primary|key|Key|tinyint|Tinyint|byte|Byte|integer|Integer|decimal|Decimal|single|Single|double|Double|currency|Currency|bit|Bit|logical|and|And|like|Like|or|Or|Logical|date|Date|memo|Memo|longbinary|Longbinary|identity|Identity|not|Not|null|Null|unique|Unique|default|DEFAULT|)\b";
            YesilAra = Regex.Matches(Rich.ToLower(new CultureInfo("en-US", false)), YesilSozluk);

            string KirmiziSozluk = @"\b(SELECT|values|VALUES|HAVING|TOP|PERCENT|DESC|ASC|DELETE|UPDATE|INSERT|INTO|ALTER|DROP|TABLE|ADD|CREATE|ORDER|WHERE|GROUP|BY|INNER|JOIN|ON|OUTER|UNION|FROM|)\b";
            KirmiziAra = Regex.Matches(Rich.ToUpper(new CultureInfo("en-US", false)), KirmiziSozluk);

            string PembeSozluk = @"\b(min|Min|max|Max|avg|var|Var|varp|varp|stdev|Stdev|stdevp|Stdevp|last|Last|Avg|sum|Sum|count|Count)\b";
            PembeAra = Regex.Matches(Rich.ToLower(new CultureInfo("en-US", false)), PembeSozluk);
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            Bs.MoveLast();


            Sıra = dataGridView1.CurrentRow.Index + 1;
            toolStripTextBox1.Text = Sıra.ToString();

        }

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            Bs.MoveNext();

            if (Sıra == 1)
            {
                toolStripButton2.Enabled = false;
                toolStripButton1.Enabled = false;
            }
            else
            {
                toolStripButton2.Enabled = true;
                toolStripButton1.Enabled = true;

            }

            Sıra = dataGridView1.CurrentRow.Index + 1;
            toolStripTextBox1.Text = Sıra.ToString();

        }


        int Sıra;
        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            Sıra = dataGridView1.CurrentRow.Index + 1;
            toolStripTextBox1.Text = Sıra.ToString();

            if (Sıra == 1)
            {
                toolStripButton2.Enabled = false;
            }
            else
            {
                toolStripButton2.Enabled = true;

            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            Bs.MovePrevious();

            if (Sıra == 1)
            {
                toolStripButton1.Enabled = false;
                toolStripButton2.Enabled = false;
            }
            else
            {
                toolStripButton1.Enabled = false;
                toolStripButton2.Enabled = true;

            }
            Sıra = dataGridView1.CurrentRow.Index + 1;
            toolStripTextBox1.Text = Sıra.ToString();

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            Bs.MoveFirst();
            
        }

        private void buttonbaglanKes_Click(object sender, EventArgs e)
        {
            ToolStripTrue(false);

            Baglan.Close();
            listBoxTbl.Items.Clear();
            listBoxAlan.Items.Clear();
            richTextBox1.Clear();
            textBox1.Clear();

            buttonbaglan.Enabled = false;
            buttonbaglanKes.Enabled = false;
            buttonYeniAcces.Enabled = false;
            buttonGorunumCalistir.Enabled = false;



        }

        private void buttonYeniAcces_Click(object sender, EventArgs e)
        {

        }
        SqlConnection SqlBaglan = new SqlConnection();

        static string BaglantiTuru, KullaniciAdi, Parola, Server;
        static string hangisi;
        private void buttonSqlBaglan_Click(object sender, EventArgs e)
        {
            hangisi = "SQL";
            try
            {
                if (comboBoxKimlik.Text == "SQL Server Authentication")
                {
                    SqlBaglan.ConnectionString = "Server=" + comboBoxSunucu.Text + ";User Id=" + textBoxKullanici.Text + ";Password=" + textBoxParola.Text + "; ";
                    BaglantiTuru = "SQL Server Authentication";
                    KullaniciAdi = textBoxKullanici.Text;
                    Parola = textBoxParola.Text;
                    Server = comboBoxSunucu.Text;
                    Baglanti();



                }
                else
                {
                    SqlBaglan.ConnectionString = "Server=" + comboBoxSunucu.Text + ";Integrated Security=True";
                    BaglantiTuru = "Windows Authentication";
                    Server = comboBoxSunucu.Text;
                    Baglanti();


                }



            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message.ToString());
            }
        }

        private void Baglanti()
        {
            if (SqlBaglan.State == ConnectionState.Closed)
            {
                SqlBaglan.Open();

                // MessageBox.Show("Bağlantı Acık");

            }
            DataTable Tablo = new DataTable();
            SqlCommand Komut = new SqlCommand("Select * From sys.Databases", SqlBaglan);
            SqlDataReader Oku = Komut.ExecuteReader();
            Tablo.Load(Oku);
            comboBoxVeritabani.DataSource = Tablo;
            comboBoxVeritabani.DisplayMember = "name";

            comboBoxSunucu.Enabled = false;
            comboBoxKimlik.Enabled = false;
            buttonSqlBaglan.Text = "Bağlandı";
            buttonSqlBaglan.BackColor = Color.Green;
            buttonSqlBaglan.Enabled = false;
            buttonTabloBaglan.Enabled = true;
            buttonBaglantiKes.Enabled = true;

        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBoxKimlik.Items.Add("Windows Authentication");
            comboBoxKimlik.Items.Add("SQL Server Authentication");


            DiziSayisi--;
            Sayfalar[DiziSayisi] = richTextBox1.Text;
            Array.Resize(ref Sayfalar, Sayfalar.Length + 1);
            SayfaSayisi++;
            tabControl2.TabPages.Add("Sayfa " + SayfaSayisi);
            richTextBox1.Text = "SELECT * FROM ";

        }

        

        private void comboBoxKimlik_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxKimlik.Text == "SQL Server Authentication")
            {

                textBoxKullanici.Enabled = true;
                textBoxParola.Enabled = true;
                buttonSqlBaglan.Enabled = true;
            }
            buttonSqlBaglan.Enabled = true;
        }

        private void buttonBaglantiKes_Click(object sender, EventArgs e)
        {
            SqlBaglan.Close();
            buttonBaglantiKes.Enabled = false;

            buttonSqlBaglan.Enabled = true;
            buttonTabloBaglan.Enabled = false;
            buttonSqlBaglan.Text = "Bağlan";
            buttonSqlBaglan.BackColor = Color.Transparent;
            buttonSqlGorunumCalistir.Enabled = false;
            buttonİslemSorguCalistir.Enabled = false;

            listBoxTbl.DataSource = null;
            listBoxAlan.DataSource = null;
        }

        private void buttonTabloBaglan_Click(object sender, EventArgs e)
        {
            buttonSqlGorunumCalistir.Enabled = true;
            buttonİslemSorguCalistir.Enabled = true;
            if (BaglantiTuru == "SQL Server Authentication")
            {
                SqlConnection BaglanTablo = new SqlConnection("Server=" + Server + ";Database=" + comboBoxVeritabani.SelectedText + ";User Id=" + KullaniciAdi + ";Password=" + Parola + ";");
                SqlCommand KomutTablo = new SqlCommand("SELECT TABLE_CATALOG ,TABLE_SCHEMA ,TABLE_NAME FROM INFORMATION_SCHEMA.TABLES", BaglanTablo);

                if (BaglanTablo.State == ConnectionState.Closed)
                {
                    BaglanTablo.Open();

                }
                DataTable SqlTablo = new DataTable();
                SqlDataReader Oku = KomutTablo.ExecuteReader();
                SqlTablo.Load(Oku);

                listBoxTbl.DataSource = SqlTablo;
                listBoxTbl.DisplayMember = "TABLE_NAME";
                

            }
            else
            {
                SqlConnection BaglanTablo = new SqlConnection("Server=" + Server + ";Database=" + comboBoxVeritabani.Text + ";Integrated Security=true");
                SqlCommand KomutTablo = new SqlCommand("SELECT TABLE_CATALOG ,TABLE_SCHEMA ,TABLE_NAME FROM INFORMATION_SCHEMA.TABLES", BaglanTablo);

                if (BaglanTablo.State == ConnectionState.Closed)
                {
                    BaglanTablo.Open();

                }
                DataTable SqlTablo = new DataTable();
                SqlDataReader Oku = KomutTablo.ExecuteReader();
                SqlTablo.Load(Oku);

                listBoxTbl.DataSource = SqlTablo;
                listBoxTbl.DisplayMember = "TABLE_NAME";

            }



        }


        private void listBoxTablo_DoubleClick(object sender, EventArgs e)
        {
            richTextBox1.SelectedText = listBoxTbl.Text;
        }

        private void buttonSqlGorunumCalistir_Click(object sender, EventArgs e)
        {
            if (BaglantiTuru == "SQL Server Authentication")
            {
                SqlConnection BaglanTablo = new SqlConnection("Server=" + Server + ";Database=" + comboBoxVeritabani.SelectedText + ";User Id=" + KullaniciAdi + ";Password=" + Parola + ";");
                SqlCommand KomutTablo = new SqlCommand(richTextBox1.Text, BaglanTablo);

                if (BaglanTablo.State == ConnectionState.Closed)
                {
                    BaglanTablo.Open();

                }
                DataTable SqlTablo = new DataTable();
                SqlDataReader Oku = KomutTablo.ExecuteReader();
                SqlTablo.Load(Oku);

                dataGridView1.DataSource = SqlTablo;

            }
            else
            {
                SqlConnection BaglanTablo = new SqlConnection("Server=" + Server + ";Database=" + comboBoxVeritabani.Text + ";Integrated Security=true");
                SqlCommand KomutTablo = new SqlCommand(richTextBox1.Text, BaglanTablo);

                if (BaglanTablo.State == ConnectionState.Closed)
                {
                    BaglanTablo.Open();

                }
                DataTable SqlTablo = new DataTable();
                SqlDataReader Oku = KomutTablo.ExecuteReader();
                SqlTablo.Load(Oku);

                dataGridView1.DataSource = SqlTablo;

            }
        }

        private void richTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            

            if (e.KeyCode == Keys.F5)
            {


                buttonGorunumCalistir_Click(null, null);
            }
            if (e.KeyCode == Keys.F6)
            {
                buttonIslemSorgu_Click(null,null);
            }
        }

        private void buttonSayfaEkle_Click(object sender, EventArgs e)
        {
            SayfaSayisi++;
            tabControl2.TabPages.Add("Sayfa " + SayfaSayisi);
            tabControl2.SelectedIndex = SayfaSayisi - 1;

            Sayfalar[tabControl2.SelectedIndex] = richTextBox1.Text;
            Array.Resize(ref Sayfalar, Sayfalar.Length + 1);
            richTextBox1.Text = "SELECT * FROM ";
        }

        private void silToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show(tabControl2.SelectedIndex.ToString());


            tabControl2.TabPages.Remove(tabControl2.SelectedTab);

            //Sayfalar[tabControl2.SelectedIndex] = "";
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {

            DialogResult Mesaj;
            Mesaj= MessageBox.Show("Yaptığınız Değişikleri Kaydetmek İstiyormusunuz...","Kaydet",MessageBoxButtons.YesNoCancel,MessageBoxIcon.Information);

            if (Mesaj == DialogResult.Yes)
            {
                SaveFileDialog Kaydet = new SaveFileDialog();

               
                if (Kaydet.ShowDialog() == DialogResult.Cancel)
                {
                    e.Cancel = true;
                }


            }
            else if (Mesaj == DialogResult.No)
            {
                e.Cancel = false;
                
            }
            else if (Mesaj == DialogResult.Cancel)
            {
                e.Cancel = true;
            }
            
        }
        void Primary(string AlanAd)
        {
            try
            {
                OleDbCommand Pkomut = new OleDbCommand("Select " + AlanAd + " From " + TabloAd + "", Baglan);

                OleDbDataReader Oku = Pkomut.ExecuteReader(CommandBehavior.KeyInfo);
                DataTable AlanTablo = Oku.GetSchemaTable();

                foreach (DataRow Alanlar in AlanTablo.Rows)
                {
                    foreach (DataColumn Aozelik in AlanTablo.Columns)
                    {


                        if (Aozelik.ColumnName == "IsKey" && Alanlar[Aozelik].ToString() == "True")
                        {

                            Birincil = "*";
                        }
                        else if (Aozelik.ColumnName == "IsKey" && Alanlar[Aozelik].ToString() == "False")
                        {
                            Birincil = "";
                        }

                    }
                }
            }
            catch (Exception)
            {


            }

        }
        static string Birincil;
        string TabloAd;
        ArrayList AlanAdi = new ArrayList();
        static int AlanSira = 0;
        private void listBoxTbl_SelectedIndexChanged(object sender, EventArgs e)
        {
            AlanAdi.Clear();
            AlanSira = 0;
            if (hangisi == "Acces")
            {

                TabloAd = listBoxTbl.SelectedItem.ToString().Substring(3, listBoxTbl.SelectedItem.ToString().Length - 3);
                listBoxAlan.Items.Clear();

                OleDbCommand Komut = new OleDbCommand("Select * from " + TabloAd + "", Baglan);



                OleDbDataAdapter Da = new OleDbDataAdapter(Komut);
                DataSet Ds = new DataSet();
                Da.Fill(Ds, TabloAd);
                DataTable Tablo = Ds.Tables[0];





                foreach (DataColumn item in Tablo.Columns)
                {
                    Primary(item.ColumnName);
                    listBoxAlan.Items.Add(item.ColumnName + Birincil + "-(" + item.DataType + ")");
                    AlanAdi.Add(item.ColumnName);
                    

                }
                

            }
            else
            {

                    SqlCommand KomutAlan = new SqlCommand("SELECT COLUMN_NAME +' - ('+DATA_TYPE+')' as Alan FROM INFORMATION_SCHEMA.COLUMNS where TABLE_NAME = '" + listBoxTbl.Text + "' ORDER BY TABLE_SCHEMA, TABLE_NAME, ORDINAL_POSITION use [" + comboBoxVeritabani.Text + "] ", SqlBaglan);

                    if (SqlBaglan.State == ConnectionState.Closed)
                    {
                        SqlBaglan.Open();
                    }
                    DataTable TabloAlan = new DataTable();
                    SqlDataReader OkuAlan = KomutAlan.ExecuteReader();
                    TabloAlan.Load(OkuAlan);

                    listBoxAlan.DataSource = TabloAlan;
                    listBoxAlan.DisplayMember = "Alan";

            }
        }

        private void splitContainer6_Panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void Form1_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Control && e.KeyCode ==Keys.G)
            {
                if (panel4.Visible == false)
                {
                    panel4.Visible = true;
                }
                else
                {
                    panel4.Visible = false;
                }
                
                
            }
        }

        private void buttonCokVeritabani_Click(object sender, EventArgs e)
        {
            
            string[] Yollar;

            OpenFileDialog Ac = new OpenFileDialog();
            Ac.Multiselect = true;
            

            Ac.ShowDialog();
            Yollar = Ac.FileNames;
            FileInfo DosyaBilgi = new FileInfo(Ac.FileName);

            textBox2.Text = DosyaBilgi.DirectoryName;

            foreach (string Ekle in Yollar)
            {
                listBoxVeritabaniListe.Items.Add(Ekle);
            }
        }

        private void tabControl2_Click(object sender, EventArgs e)
        {
            if (Sayfalar[tabControl2.SelectedIndex] != null)
            {
                richTextBox1.Text = Sayfalar[tabControl2.SelectedIndex].ToString();

            }
        }

        private void listBoxTbl_DoubleClick(object sender, EventArgs e)
        {
            richTextBox1.SelectedText = " "+TabloAd+" ";
        }

        private void listBoxAlan_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void listBoxAlan_DoubleClick(object sender, EventArgs e)
        {
            richTextBox1.SelectedText =" "+ AlanAdi[listBoxAlan.SelectedIndex].ToString()+" ";
        }

        private void buttonIslemSorgu_Click(object sender, EventArgs e)
        {
            int EtkilenenSatir =0;

            try
            {
                if (Baglan.State == ConnectionState.Closed)
                {
                    Baglan.Open();
                }


                if (richTextBox1.SelectedText != null)
                {

                    //DataTable STablo = new DataTable();
                    OleDbCommand SKomut = new OleDbCommand(richTextBox1.SelectedText, Baglan);

                    EtkilenenSatir = SKomut.ExecuteNonQuery();
                    
                        MessageBox.Show(EtkilenenSatir.ToString()+" Satır kayıt Etkilendi...");
                    
                   

                }

            }
            catch (Exception)
            {


                try
                {
                    //DataTable Tablo = new DataTable();
                    OleDbCommand Komut = new OleDbCommand(richTextBox1.Text, Baglan);
                    EtkilenenSatir = Komut.ExecuteNonQuery();
                    
                        MessageBox.Show(EtkilenenSatir.ToString() + " Satır kayıt Etkilendi...");
                    
                   
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.Message.ToString());
                }

            }
        }

        private void gösterToolStripMenuItem_Click(object sender, EventArgs e)
        {

            if (Baglan.State==ConnectionState.Closed)
            {
                Baglan.Open();
            }
            DataTable Tablo = new DataTable();
            OleDbCommand Komut = new OleDbCommand("Select * from "+TabloAd+"", Baglan);

            OleDbDataReader Oku = Komut.ExecuteReader();
            Tablo.Load(Oku);
            Bs.DataSource = Tablo;
            dataGridView1.DataSource = Bs;


            toolStripLabelSayi.Text = dataGridView1.Rows.Count.ToString();

            ToolStripTrue(true);
        }

        private void buttonİslemSorguCalistir_Click(object sender, EventArgs e)
        {

        }

        private void buttonKaydet_Click(object sender, EventArgs e)
        {

        }

        


    }
}
