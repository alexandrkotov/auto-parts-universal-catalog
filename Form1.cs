//версия 0.94.01.RC18 
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SQLite;
using System.IO;
using System.Drawing.Printing;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Mail;
using System.Xml.Linq;
using mshtml;
using System.Net.Sockets;
using Newtonsoft.Json;
using System.Net;
using System.Net.Cache;
using System.Text.RegularExpressions;
using EASendMail;


namespace GB_SQLite_SC
{
    public partial class Form1 : Form
    {
        public DateTime tf = new DateTime(2026, 5, 13);//время ограничения работы программы
        public bool copyCellData = false;//true;//false;//разрешить/запретить автоматическое копирование данных в буфер обмена из ячеки dataGridView_linktable
        public bool button_csv_visible = false;//true;//показ кнопки импорта в excel - для клиентов отключить
        public bool softutc = false;
        public bool pcd_omega = false;
        public Int32 utcY = 0;//год с сервера time.nist.gov
        public Int32 utcM = 0;//месяц с сервера time.nist.gov
        public Int32 utcD = 0;//день с сервера time.nist.gov

        public string current_brandname;
        public string current_brandname_reverse;
        public string current_gearboxgroup;
        public string current_gearboxgroup_reverse;
        public int current_num_gb;
        public int current_num_gb_reverse;
        public string current_omegacode;
        public string current_description;
        public string current_specification;
        public string current_notes;
        public int current_fig;
        public int current_fig_reverse;
        public string current_fig_linktable;
        public bool filter_empty;
        public string adtsc;
        public int curprintpage;
        public int last_row_to_print;
        public bool no_notes_on_sorting;

        private Excel.Application excelapp;

        public const string GMAIL_SERVER = "smtp.gmail.com";
        public const int PORT = 587;
        public string ip = "";
        public string GMailUser = "";
        public string GMailPass = "";
        public string GMailUserOMEGA = "mail@gmail.com";
        public string GMailPassOMEGA = "password";
        public bool email_error = false;

        DataTable dt_gearbox_combobox = new DataTable(null);
        DataTable dt_gearbox = new DataTable(null);
        DataTable dt_infobase = new DataTable(null);
        DataTable dt_linktable = new DataTable(null);
        DataTable dt_reverse = new DataTable(null);
        DataTable dt_expdate = new DataTable(null);

        

        public bool bigsize = false;
        public string normal_search_string = "";

        public bool clear_background;// = true;
        public string header_ourno = "№ ourno";
        public string header_partno = "№ partno";
        public string header_itemcod = "код товара";
        public string header_type = "тип";

        public string current_version = "";
        public bool one_image;
        public byte[] htmlBytes_one;
        public byte[] htmlBytes_one_print;

        public class JsonConfig
        {
            public string SplashText { get; set; }
            public string ToolTipButtonHelp { get; set; }
            public string ToolTipButtonReverseSearch { get; set; }
            public string ToolTipButtonWs { get; set; }
            public string ToolTipListViewOrder { get; set; }
            public string FormText { get; set; }
            public string LabelReverseSearchInfo { get; set; }
            public string GroupBoxWsLogin { get; set; }
            public string HeaderOurNo { get; set; }
            public string HeaderPartNo { get; set; }
            public string HeaderItemCod { get; set; }
            public string HeaderType { get; set; }
            public string Olga { get; set; }
            public string softutc { get; set; }
        }

        public Form1()
        {
            
            InitializeComponent();
            dataGridView_gearbox_select.MouseWheel += new MouseEventHandler(dataGridView_gearbox_select_MouseWheel);
            dataGridView_gearbox.MouseWheel += new MouseEventHandler(dataGridView_gearbox_MouseWheel);
            dataGridView_reverse_search.MouseWheel += new MouseEventHandler(dataGridView_reverse_search_MouseWheel);
            listView_order.MouseDown += new MouseEventHandler(listView_order_right_click); 
            try
            {
                ip = new System.Net.WebClient().DownloadString("https://ipinfo.io/ip").Replace("\n", "");
            }
            catch { }
            check_login_file();
            check_gmail_login_file();
            check_pcd_file();
            check_config_file();
            listView_order.Columns[1].Text = header_ourno;
            listView_order.Columns[2].Text = header_partno;
            radioButton_type_ourno.Text = header_ourno;
            radioButton_type_partno.Text = header_partno;

        }

        //http://stackoverflow.com/questions/6435099/how-to-get-datetime-from-the-internet
        public static DateTime GetNistTime()
        {
            DateTime dateTime = DateTime.MinValue;

            HttpWebRequest request = (HttpWebRequest)WebRequest.Create("http://nist.time.gov/actualtime.cgi?lzbc=siqm9b");
            request.Method = "GET";
            request.Accept = "text/html, application/xhtml+xml, */*";
            request.UserAgent = "Mozilla/5.0 (compatible; MSIE 10.0; Windows NT 6.1; Trident/6.0)";
            request.ContentType = "application/x-www-form-urlencoded";
            request.CachePolicy = new RequestCachePolicy(RequestCacheLevel.NoCacheNoStore); //No caching
            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            if (response.StatusCode == HttpStatusCode.OK)
            {
                StreamReader stream = new StreamReader(response.GetResponseStream());
                string html = stream.ReadToEnd();//<timestamp time=\"1395772696469995\" delay=\"1395772696469995\"/>
                string time = Regex.Match(html, @"(?<=\btime="")[^""]*").Value;
                double milliseconds = Convert.ToInt64(time) / 1000.0;
                dateTime = new DateTime(1970, 1, 1).AddMilliseconds(milliseconds).ToLocalTime();
            }

            return dateTime;
        }

        //запрос даты с сервера точного времени
        public void get_utc()
        {
            try
            {
                DateTime t = GetNistTime();
                //DateTime t = DateTime.Now;
                utcY = t.Year;
                utcM = t.Month;
                utcD = t.Day;
            }
            
            /*
            try
            {
                //MessageBox.Show("дата из expdate: " + expdate);
                //http://stackoverflow.com/questions/6435099/how-to-get-datetime-from-the-internet/9830462#9830462
                var client = new TcpClient("time.nist.gov", 13);
                //var client = new TcpClient("ntp.nasa.gov", 13);
                //var client = new TcpClient("ntp1.vniiftri.ru", 13);
                var streamReader = new StreamReader(client.GetStream());
                var response = streamReader.ReadToEnd();
                utcY = 2000 + Int32.Parse(response.Substring(7, 2).ToString());
                utcM = Int32.Parse(response.Substring(10, 2).ToString());
                utcD = Int32.Parse(response.Substring(13, 2).ToString());
            }
            */
            catch
            {
                if (softutc == true)
                {
                    //MessageBox.Show("Проверьте соединение с интернетом и перезапустите программу");
                    //MessageBox.Show("Нестабильное соединение с интернетом");
                    DateTime t = DateTime.Now;
                    utcY = t.Year;
                    utcM = t.Month;
                    utcD = t.Day;

                    return;
                }
                else
                {
                    MessageBox.Show("Проверьте, есть ли доступ к сайту http://nist.time.gov/ и перезапустите программу");
                    //return; //так нельзя, иначе если продолжить после ошибки, которую выдает framework, программа запустится
                    Application.Exit();
                }

            }
        }

        //проверка триального периода 
        public void check_date_finish()
        {
            //по простому, пользователь может откатить системную дату назад и работать дальше
            DateTime t = DateTime.Now;//проверка по локальному времени компьютера, обходится откатом времени на ПК
            DateTime t2 = new DateTime(utcY, utcM, utcD);//проверка по серверу точного времени в интернете

            if (t.CompareTo(tf) != -1)
            {
                MessageBox.Show("Программа устарела. Запросите новую версию.");
                Application.Exit();
            }
            else
            {
                return;
            }            
        }

        
        //проверка даты подписки на схемы
        public bool check_sch_expdate()
        {

            string expdate = dt_expdate.Rows[0].ItemArray[0].ToString();
            //int utcY = 0;
            //int utcM = 0;
            //int utcD = 0;

            int expY = 2000 + Int32.Parse(expdate.Substring(8, 2).ToString());
            int expM = Int32.Parse(expdate.Substring(3, 2).ToString());
            int expD = Int32.Parse(expdate.Substring(0, 2).ToString());



            DateTime exp = new DateTime(expY, expM, expD);
            DateTime utc = new DateTime(utcY, utcM, utcD);

            if (utc.CompareTo(exp) != -1)
            {
                //MessageBox.Show("Срок подписки на схемы истёк "+ exp.ToString().Substring(0,10));
                dataGridView_linktable.Visible = false;
                label_expdate.Text = "Срок подписки на схемы истёк " + exp.ToString().Substring(0, 10);
                label_expdate.Visible = true;
                return false;
            }

            //MessageBox.Show("Подписка - Ок");
            dataGridView_linktable.Visible = true;
            label_expdate.Visible = false;
            this.Text = this.Text + "           Подписка на схемы действительна до " + exp.ToString().Substring(0, 10);
            return true;


        }
        
        //проверка разрешения экрана
        public bool check_screen()
        {
            int hScreen = Screen.PrimaryScreen.WorkingArea.Height;
            int wScreen = Screen.PrimaryScreen.WorkingArea.Width;
            if(hScreen < 900)
            {
                MessageBox.Show("Внимание! Несоблюдение технических требований\nРазрешение монитора по вертикали меньше 900 px");
                return false;
            }
            if (wScreen < 1200)
            {
                MessageBox.Show("Внимание! Несоблюдение технических требований\nРазрешение монитора по горизонтали меньше 1200 px");
                return false;
            }
            return true;
        }


        public void SplashScreen()
        {
            try
            {
                Application.Run(new Form2());
            }
            catch { }
        }

        void dataGridView_gearbox_select_MouseWheel(object sender, MouseEventArgs e)
        {
            int ccol = dataGridView_gearbox_select.CurrentCell.ColumnIndex;
            int crow = dataGridView_gearbox_select.CurrentCell.RowIndex;

            try
            {
                if (e.Delta < 0)
                {
                    dataGridView_gearbox_select.CurrentCell = dataGridView_gearbox_select[ccol, crow + 1];
                }
                else
                {
                    dataGridView_gearbox_select.CurrentCell = dataGridView_gearbox_select[ccol, crow - 1];
                }
            }
            catch { }
        }

        void dataGridView_gearbox_MouseWheel(object sender, MouseEventArgs e)
        {
            int ccol = dataGridView_gearbox.CurrentCell.ColumnIndex;
            int crow = dataGridView_gearbox.CurrentCell.RowIndex;

            try
            {
                if (e.Delta < 0)
                {
                    dataGridView_gearbox.CurrentCell = dataGridView_gearbox[ccol, crow + 1];
                }
                else
                {
                    dataGridView_gearbox.CurrentCell = dataGridView_gearbox[ccol, crow - 1];
                }
            }
            catch { }
        }
        //dataGridView_reverse_search_MouseWheel
        void dataGridView_reverse_search_MouseWheel(object sender, MouseEventArgs e)
        {
            int ccol = dataGridView_reverse_search.CurrentCell.ColumnIndex;
            int crow = dataGridView_reverse_search.CurrentCell.RowIndex;

            try
            {
                if (e.Delta < 0)
                {
                    dataGridView_reverse_search.CurrentCell = dataGridView_reverse_search[ccol, crow + 1];
                }
                else
                {
                    dataGridView_reverse_search.CurrentCell = dataGridView_reverse_search[ccol, crow - 1];
                }
            }
            catch { }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            
            //заставка
            Thread t = new Thread(new ThreadStart(SplashScreen));
            try
            {
                t.Start();
            }
            catch { }
            /*
            if(button_csv_visible == true)
            {
                button_csv.Visible = true;
            }
            */
            //инициация глобальных переменных
            current_brandname = "";
            current_brandname_reverse = "";
            current_gearboxgroup = "";
            current_gearboxgroup_reverse = "0";
            current_num_gb = 0;
            current_num_gb_reverse = 0;
            current_omegacode = "";
            current_fig = 0;
            current_fig_reverse = 0;
            current_fig_linktable = "";
            filter_empty = true;
            adtsc = "";
            curprintpage = 0;
            last_row_to_print = 0;
            current_description = "";
            current_specification = "";
            current_notes = "";
            clear_background = true;
            no_notes_on_sorting = false;
            one_image = false;
            //
            webBrowser1.DocumentStream = null;
            dataGridView_gearbox.ForeColor = Color.Gray;
            //привязываем тултипы
            toolTip_filter.SetToolTip(textBox_filter_series, "Фильтр для поиска, начните вводить текст и список ниже будет автоматически отфильтрован.");
            toolTip_filter_icon.SetToolTip(pictureBox2, "Фильтр для поиска, начните вводить текст в поле справа и список ниже будет автоматически отфильтрован.");
            //включение подсветки всей строки
            dataGridView_gearbox_select.SelectionMode = DataGridViewSelectionMode.FullRowSelect;

            //выборка из БД
            
            string conStr = "Data Source=current.sch;Password=Uhjlyj1993Ytv@y;";//gb58-62.sqlite//gb58-213.sqlite
            string mSQL_gearbox_combobox = "SELECT distinct brandname,gearboxgroup FROM gearbox";
            string mSQL_gearbox = "SELECT * FROM gearbox";
            string mSQL_infobase = "SELECT num, html, jpg FROM infobase";
            string mSQL_linktable = "SELECT * FROM linktable";
            string mSQL_expdate = "SELECT * FROM expdate";
            
            //вью. чтобы избежать колизий при наложении разных фильтров на одну и ту же таблицу
            DataView dv_gearbox1 = new DataView(null);
            
            try
            {
                SQLiteConnection concheck = new SQLiteConnection(conStr);
                concheck.Open();
                concheck.ChangePassword("password");
                concheck.Close();
            }
            catch
            {
                SQLiteConnection concheck = new SQLiteConnection("Data Source=current.sch;");
                concheck.Open();
                concheck.ChangePassword("password");
                concheck.Close();
            }

            try
            {
                SQLiteConnection con = new SQLiteConnection(conStr);
                con.Open();

                SQLiteCommand command_gearbox_combobox = new SQLiteCommand(mSQL_gearbox_combobox, con);
                SQLiteDataReader reader_gearbox_combobox = command_gearbox_combobox.ExecuteReader();
                dt_gearbox_combobox.Load(reader_gearbox_combobox);

                SQLiteCommand command_gearbox = new SQLiteCommand(mSQL_gearbox, con);
                SQLiteDataReader reader_gearbox = command_gearbox.ExecuteReader();
                dt_gearbox.Load(reader_gearbox);

                SQLiteCommand command_infobase = new SQLiteCommand(mSQL_infobase, con);
                SQLiteDataReader reader_infobase = command_infobase.ExecuteReader();
                dt_infobase.Load(reader_infobase);



                SQLiteCommand command_linktable = new SQLiteCommand(mSQL_linktable, con);
                SQLiteDataReader reader_linktable = command_linktable.ExecuteReader();
                dt_linktable.Load(reader_linktable);

                SQLiteCommand command_expdate = new SQLiteCommand(mSQL_expdate, con);
                SQLiteDataReader reader_expdate = command_expdate.ExecuteReader();
                dt_expdate.Load(reader_expdate);
                con.Close();

                //считаем количество схем в базе
                label_count_schema.Text  = dt_gearbox.Rows.Count.ToString() + " схем";   


                //грид для выбора бренда и группы КПП, цепляем напрямую к таблице 
                dataGridView_gearbox_select.DataSource = dt_gearbox_combobox;
                dataGridView_gearbox_select.Columns[0].Width = 80;// 100;
                dataGridView_gearbox_select.Columns[1].Width = 204;

                //грид для отображения КПП, цепляем таблица->вью->биндинг->грид
                dv_gearbox1.Table = dt_gearbox;
                bindingSource_gearbox.DataSource = dv_gearbox1;
                dataGridView_gearbox.DataSource = bindingSource_gearbox; //dt_gearbox;
                //настройка ширина колонок в гриде
                dataGridView_gearbox.Columns[0].Width = 0;
                dataGridView_gearbox.Columns[1].Width = 0;
                dataGridView_gearbox.Columns[2].Width = 0;
                dataGridView_gearbox.Columns[3].Width = 530;
                dataGridView_gearbox.Columns[4].Width = 215;
                dataGridView_gearbox.Columns[5].Visible = false;
                dataGridView_gearbox.Columns[6].Visible = false;


                //грид для отображения кросс-таблицы ссылок таблица->биндинг->грид
                bindingSource_linktable.DataSource = dt_linktable;
                dataGridView_linktable.DataSource = bindingSource_linktable;//  dt_linktable;   

                //просто ПОДСКАЗКА для возможной реализации, так же см. строку 617
                //красный шрифт в ячейках с товаром, начинающимся со знака # - делать надо при любом изменении грида
                //dataGridView_linktable.Columns[4].DefaultCellStyle.ForeColor = Color.Red;
                //dataGridView_linktable.Rows[2].Cells[4].Style = new DataGridViewCellStyle { ForeColor = Color.Red };


                //названия столбцов в гриде linktable
                dataGridView_linktable.Columns[0].HeaderText = "№ схемы";
                dataGridView_linktable.Columns[0].Width = 60;
                dataGridView_linktable.Columns[1].HeaderText = "позиция";
                dataGridView_linktable.Columns[1].Width = 65;
                dataGridView_linktable.Columns[2].HeaderText = header_ourno;// "№ Euroricambi";
                dataGridView_linktable.Columns[2].Width = 115;
                dataGridView_linktable.Columns[3].HeaderText = header_partno;//"№ оригинальный";
                dataGridView_linktable.Columns[3].Width = 120;
                dataGridView_linktable.Columns[4].HeaderText = header_itemcod;// "код товара в ADTS";
                dataGridView_linktable.Columns[4].Width = 130;
                dataGridView_linktable.Columns[5].HeaderText = "описание";
                dataGridView_linktable.Columns[5].Width = 380;// 280;
                dataGridView_linktable.Columns[6].HeaderText = header_type;// "Type(old/new)";
                dataGridView_linktable.Columns[6].Width = 90;// 85;
                dataGridView_linktable.Columns[7].HeaderText = "кол-во";
                dataGridView_linktable.Columns[7].Width = 55;// 50;
                dataGridView_linktable.Columns[8].Visible = false;


                t.Abort();

                dataGridView_gearbox_select.Focus();
                dataGridView_gearbox_select.Select();
            }
            catch
            {
                MessageBox.Show("Ошибка соединения с хранилищем схем.");
                t.Abort();
            }
            //строим dt_reverse
            dt_reverse.Columns.Add(new DataColumn("num", typeof(int)));
            dt_reverse.Columns.Add(new DataColumn("brandname", typeof(string)));
            dt_reverse.Columns.Add(new DataColumn("series", typeof(string)));
            dt_reverse.Columns.Add(new DataColumn("specification", typeof(string)));
            dt_reverse.Columns.Add(new DataColumn("gearboxgroup", typeof(string)));

            //запрос utcY, utcM, utcD
            get_utc();

            //проверка срока подписки на схемы
            check_sch_expdate();

            //проверка триального периода
            check_date_finish();

            //проверка PCD
            if(pcd_omega == true)
            {
                //textBox_email_to_FIO
                //textBox_email_to_job_title
                //textBox_email_to_company
                //textBox_email_to_phone
                //textBox_email_to_phone_ext
                //textBox_email_to_EMAIL
                //textBox_email_to_EMAIL_domain
                textBox_email_to_job_title.Text = "Менеджер по продажам";
                textBox_email_to_job_title.Enabled = false;
                textBox_email_to_company.Text = "КОМПАНИЯ";
                textBox_email_to_company.Enabled = false;
                textBox_email_to_phone.Text = "+7 (495) 123-45-67 ";
                textBox_email_to_phone.Enabled = false;
                textBox_email_to_EMAIL_domain.Text = "company.ru";
                textBox_email_to_EMAIL_domain.Enabled = false;
            }

            //проверка одной встроенной картинки для всех схем one_image=true если в dt_infobase в первом поле есть значение 0           
            DataRow[] rows = dt_infobase.Select("num = '0'");
            try
            {
                if (rows[0][0].ToString() == "0")
                {
                    //MessageBox.Show("База с одним рисунком");
                    one_image = true;
                    DataRow[] rows_one = dt_infobase.Select("num = '0'");
                    int irow_one = dt_infobase.Rows.IndexOf(rows_one[0]);
                    htmlBytes_one = (byte[])dt_infobase.Rows[irow_one].ItemArray[1]; // массив с общей картинкой
                    htmlBytes_one_print = (byte[])dt_infobase.Rows[irow_one].ItemArray[2]; // массив с общей картинкой для печати
                }
            }
            catch { }

            //проверка разрешения монитора
            check_screen();
        }

        public void getmap()
        {
            //http ://stackoverflow.com/questions/2658054/converting-to-byte-array-after-reading-a-blob-from-sql-in-c-sharp
            //http ://www.sql.ru/forum/754748/kak-poluchit-nomer-stroki-datatable

            DataRow[] rows = dt_infobase.Select("num = '" + current_num_gb + "'");
            DataRow[] ds = dt_gearbox.Select("num = '" + current_num_gb + "'");
            int irow = dt_infobase.Rows.IndexOf(rows[0]);

            byte[] htmlBytes = (byte[])dt_infobase.Rows[irow].ItemArray[1];//Rows[0].ItemArray[1];

            if (one_image == true)
            {
                //DataRow[] rows_one = dt_infobase.Select("num = '0'");
                //int irow_one = dt_infobase.Rows.IndexOf(rows_one[0]);
                //htmlBytes_one = (byte[])dt_infobase.Rows[irow_one].ItemArray[1]; // массив с общей картинкой

                byte[] htmlBytes_compiled = new byte[htmlBytes_one.Length + htmlBytes.Length];
                System.Buffer.BlockCopy(htmlBytes_one, 0, htmlBytes_compiled, 0, htmlBytes_one.Length);
                System.Buffer.BlockCopy(htmlBytes, 0, htmlBytes_compiled, htmlBytes_one.Length, htmlBytes.Length);

                MemoryStream stream = new MemoryStream(htmlBytes_compiled, 0, htmlBytes_compiled.Length);
                webBrowser1.DocumentStream = stream;
            }
            else {

                MemoryStream stream = new MemoryStream(htmlBytes, 0, htmlBytes.Length);
                webBrowser1.DocumentStream = stream;
            }

            current_description = ds[0][3].ToString();
            current_specification = ds[0][4].ToString();

            //select_fig_after_reverse();
            clear_background = true;
            panelNotes.Visible = false;
            //current_notes через try catch
            try
            {
                current_notes = ds[0][8].ToString();
                if(current_notes.Length > 3)
                {
                    buttonNotes.Enabled = true;
                }
                else
                {
                    buttonNotes.Enabled = false;
                }
            }
            catch
            {
                current_notes = "";
                buttonNotes.Enabled = false;
            }

        }

        private void button_order_clear_all_Click(object sender, EventArgs e)
                {
                    listView_order.Items.Clear();
                }

        private void button_order_clear_last_Click(object sender, EventArgs e)
                {
                    try
                    {
                        int lastitem = listView_order.Items.Count;
                        listView_order.Items[lastitem - 1].Remove();
                    }
                    catch { }
                }
        

        private void dataGridView_gerabox_select_SelectionChanged(object sender, EventArgs e)
        {

            int c = dataGridView_gearbox_select.CurrentCell.ColumnIndex;
            int r = dataGridView_gearbox_select.CurrentCell.RowIndex;
            current_brandname = dataGridView_gearbox_select[0, r].Value.ToString();  //.CurrentCell.Value.ToString(); //.ToString();
            current_gearboxgroup = dataGridView_gearbox_select[1, r].Value.ToString();

            if (c != 1)
            {
                dataGridView_gearbox_select.CurrentCell = dataGridView_gearbox_select[1, r];
            }

            label1.Text = current_brandname;
            label2.Text = current_gearboxgroup;
            if (checkBox_All_models.Checked == false)
            {
                bindingSource_gearbox.Filter = "brandname = '" + current_brandname + "' AND gearboxgroup = '" + current_gearboxgroup + "'";
            }
           // else
           // {
           //     bindingSource_gearbox.Filter = "";
           // }
            listView_order.Items.Clear();
            clear_background = true;
        }

        private void dataGridView_gearbox_SelectionChanged(object sender, EventArgs e)
        {
            int c = dataGridView_gearbox.CurrentCell.ColumnIndex;
            int r = dataGridView_gearbox.CurrentCell.RowIndex;


            current_num_gb = Int32.Parse((dataGridView_gearbox[0, r].Value.ToString()));
            current_omegacode = dataGridView_gearbox[1, r].Value.ToString();

            label3.Text = current_num_gb.ToString();
            label4.Text = current_omegacode;

            if (c != 3)
            {
                dataGridView_gearbox.CurrentCell = dataGridView_gearbox[3, r];
            }

            getmap();

            bindingSource_linktable.Filter = "num = '" + current_num_gb + "'";
            listView_order.Items.Clear();
            if (checkBox_All_models.Checked == true)
            {
             //вывести название производителя и название семейства (группы) на dt_gearbox_combobox
             DataRow[] dsall = dt_gearbox.Select("num = '" + current_num_gb + "'");
             label_no_groups_brandname.Text = dsall[0][2].ToString();
             label_no_groups_brandname.Refresh();
             label_no_groups_group.Text = dsall[0][5].ToString();
             label_no_groups_group.Refresh();
            }
        }



        private void dataGridView_gearbox_select_KeyDown(object sender, KeyEventArgs e)
        {

            int curcol = dataGridView_gearbox_select.CurrentCell.ColumnIndex;
            int currow = dataGridView_gearbox_select.CurrentCell.RowIndex;
            //переход по стрелке влево
            if (e.KeyValue == 39 & curcol == 1)
            {
                dataGridView_gearbox.Focus();
            }
            //переход по Tab в грид коробок
            if (e.KeyCode == Keys.Tab)
            {
                if (currow > 0)
                {
                    dataGridView_gearbox_select.CurrentCell = dataGridView_gearbox_select[1, currow - 1];
                }
                dataGridView_gearbox.Focus();
            }

        }

        private void dataGridView_gearbox_KeyDown(object sender, KeyEventArgs e)
        {

            int currow = 0;
            int curcol = 0;
            //защита от пустого грида, если фильтр кривой
            try
            {
                currow = dataGridView_gearbox.CurrentCell.RowIndex;
                curcol = dataGridView_gearbox.CurrentCell.ColumnIndex;
            }
            catch
            {
                textBox_filter_series.Text = "";
                bindingSource_gearbox.Filter = "brandname = '" + current_brandname + "' AND gearboxgroup = '" + current_gearboxgroup + "'";
                currow = dataGridView_gearbox.CurrentCell.RowIndex;
                curcol = dataGridView_gearbox.CurrentCell.ColumnIndex;
            }


            if (e.KeyValue == 37 & curcol == 3 & checkBox_All_models.Checked == false)
            {

                textBox_filter_series.Text = "";
                bindingSource_gearbox.Filter = "brandname = '" + current_brandname + "' AND gearboxgroup = '" + current_gearboxgroup + "'";
                dataGridView_gearbox.CurrentCell = dataGridView_gearbox[3, 0];
                dataGridView_gearbox_select.Focus();
            }
            if (e.KeyCode == Keys.Up & currow == 0)
            {
                textBox_filter_series.Focus();
            }
        }

        private void button_help_Click(object sender, EventArgs e)
        {
            if (pictureBox1.Visible == false)
            {
                pictureBox1.Visible = true;
                panel_reverse_search.Visible = false;
                panelNotes.Visible = false;
            }
            else
            {
                pictureBox1.Visible = false;
            }

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = false;
        }

        private void webBrowser1_Navigated(object sender, WebBrowserNavigatedEventArgs e)
        {
            try
            {
                //MessageBox.Show("111");
                string url = webBrowser1.Url.ToString();

                string[] qq = url.Split('#');
                current_fig = Int32.Parse(qq[1]);
                current_fig_linktable = current_fig.ToString();
                label5.Text = current_fig.ToString();

                bindingSource_linktable.Filter = "num = '" + current_num_gb + "' AND fig = '" + current_fig + "'";
                //MessageBox.Show(url + " длина=" + url.Length.ToString()+" фигура="+fig.ToString());
                dataGridView_linktable_cell_background();
                clear_background = false;
            }
            catch (Exception ex) { }
        }

        private void dataGridView_gearbox_select_Leave(object sender, EventArgs e)
        {
            dataGridView_gearbox_select.BorderStyle = 0;
            dataGridView_gearbox_select.ForeColor = Color.Gray;
            dataGridView_gearbox.BorderStyle = BorderStyle.FixedSingle;
            dataGridView_gearbox.ForeColor = Color.Black;
        }

        private void dataGridView_gearbox_select_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            dataGridView_gearbox_select.BorderStyle = BorderStyle.FixedSingle;
            dataGridView_gearbox_select.ForeColor = Color.Black;
            dataGridView_gearbox.BorderStyle = 0;
            dataGridView_gearbox.ForeColor = Color.Gray;
            
        }


        private void dataGridView_gearbox_DoubleClick(object sender, EventArgs e)
        {
            //вывод длинного текста спецификации в listBox_specification
            int r = dataGridView_gearbox.CurrentCell.RowIndex;
            string longspec = dataGridView_gearbox[4, r].Value.ToString();

            richTextBox_specification.Text = dataGridView_gearbox[4, r].Value.ToString();
            if (longspec.Length > 100)
            {
                richTextBox_specification.Visible = true;
            }
            else
            {
                richTextBox_specification.Visible = false;
            }
        }

        private void dataGridView_gearbox_Leave(object sender, EventArgs e)
        {
            richTextBox_specification.Visible = false;
            dataGridView_gearbox.BorderStyle = 0;
        }

        private void dataGridView_gearbox_CellLeave(object sender, DataGridViewCellEventArgs e)
        {
            richTextBox_specification.Visible = false;
        }

        private void textBox_filter_series_TextChanged(object sender, EventArgs e)
        {
            try
            {
                if (checkBox_All_models.Checked == false)
                {
                    bindingSource_gearbox.Filter = "brandname = '" + current_brandname + "' AND gearboxgroup = '" + current_gearboxgroup + "' AND series LIKE '%" + textBox_filter_series.Text + "%'";
                }
                else
                {
                    bindingSource_gearbox.Filter = "series LIKE '%" + textBox_filter_series.Text + "%'";
                    bindingSource_gearbox.Sort = "series ASC";
                }  
                filter_empty = false;
            }
            catch
            {
                //bindingSource_gearbox.Filter = "brandname = '" + current_brandname + "' AND gearboxgroup = '" + current_gearboxgroup + "'";
                filter_empty = true;
            }

        }

        private void textBox_filter_series_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Down)
            {
                dataGridView_gearbox.Focus();
            }

        }

        private void textBox_filter_series_Leave(object sender, EventArgs e)
        {

            textBox_filter_series.BackColor = Color.FromArgb(240, 240, 240); //#f0f0f0;
            textBox_filter_series.BorderStyle = BorderStyle.Fixed3D;
            if (filter_empty == true)
            {
                textBox_filter_series.Text = "";
                if (checkBox_All_models.Checked == false)
                {
                    bindingSource_gearbox.Filter = "brandname = '" + current_brandname + "' AND gearboxgroup = '" + current_gearboxgroup + "'";
                }
            }
        }

        private void textBox_filter_series_Enter(object sender, EventArgs e)
        {
            textBox_filter_series.BackColor = Color.White;
            textBox_filter_series.BorderStyle = BorderStyle.FixedSingle;
            pictureBox1.Visible = false;
        }

        private void dataGridView_gearbox_Enter(object sender, EventArgs e)
        {
            dataGridView_gearbox.BorderStyle = BorderStyle.FixedSingle;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            textBox_filter_series.Focus();
            pictureBox1.Visible = false;

        }

        private void dataGridView_gearbox_select_Enter(object sender, EventArgs e)
        {
            dataGridView_gearbox.CurrentCell = dataGridView_gearbox[3, 0];
            textBox_filter_series.Text = "";
        }

        private void dataGridView_linktable_cell_background()
        {
            int color = 0;
            bool adts_ok = false;
            int r = dataGridView_linktable.RowCount;
            int c = dataGridView_linktable.ColumnCount;
            //проверка наличия товара АДТС
            for (int ib = 0; ib < r; ib++)
            {
                if (dataGridView_linktable[4, ib].Value.ToString().Length > 3)
                {
                    adts_ok = true;
                    break;
                }
            }
            //проверка по количеству строк и наличию товара
            if (adts_ok == false)
            {
                color = 3;
            }
            else
            {
                if (r > 1)
                {
                    color = 2;
                }
                else
                {
                    color = 1;
                }
            }
            for (int ir = 0; ir < r; ir++)
            {
                for (int ic = 0; ic < c; ic++)
                {
                    if (color == 0)
                    {
                        dataGridView_linktable[ic, ir].Style.BackColor = Color.FromArgb(255, 255, 255);//белый
                    }
                    else if (color == 1)
                    {
                        dataGridView_linktable[ic, ir].Style.BackColor = Color.FromArgb(128, 255, 128);//зеленый
                    }
                    else if (color == 2)
                    {
                        dataGridView_linktable[ic, ir].Style.BackColor = Color.FromArgb(255, 255, 128);//желтый
                        //меняем цвет шрифта на красный в строках, где нет товара
                        if (dataGridView_linktable[4, ir].Value.ToString().Length < 3)
                        {
                            dataGridView_linktable[ic, ir].Style.ForeColor = Color.Red;
                        }

                    }
                    else if (color == 3)
                    {
                        dataGridView_linktable[ic, ir].Style.BackColor = Color.FromArgb(255, 128, 128);//красный
                    }
                }
            }
        }

        private void button_print_Click(object sender, EventArgs e)
        {
            //печать
            //сначала проверка, что listView_order не пустой
            curprintpage = 0;
            last_row_to_print = 0;
            if (listView_order.Items.Count > 0)
            {
                try
                {
                    PrintDialog print_dialog = new PrintDialog();
                    PrintDocument document = new PrintDocument();

                    document.PrintPage += new PrintPageEventHandler(document_PrintPage);
                    print_dialog.Document = document;

                    DialogResult result = print_dialog.ShowDialog();

                    if (result == DialogResult.OK)
                    {
                        document.Print();
                    }

                }
                catch
                {
                    throw;
                }
            }
        }
        private string listView_order_item_to_doc(int i)
        {
            string strtoprint = "";
            string strEuroricambi = "";
            //string strPattern = "                           ";//27 пробелов

            if(listView_order.Items[i].SubItems[1].Text.Length > 3)
            {
                strEuroricambi = listView_order.Items[i].SubItems[1].Text;
            }
            else
            {
                strEuroricambi = "- - - - - - - -        ";
            }
            if (listView_order.Items[i].SubItems[0].Text.Length == 1)
            {
                strtoprint = string.Format("{0,-17}{1,-24}{2,-25}{3,-30}{4,-30}",
                    listView_order.Items[i].SubItems[0].Text,
                    //listView_order.Items[i].SubItems[1].Text,
                    strEuroricambi,
                    listView_order.Items[i].SubItems[2].Text,
                    listView_order.Items[i].SubItems[3].Text,
                    listView_order.Items[i].SubItems[4].Text);
            }
            else if (listView_order.Items[i].SubItems[0].Text.Length == 2)
            {
                strtoprint = string.Format("{0,-16}{1,-24}{2,-25}{3,-30}{4,-30}",
                    listView_order.Items[i].SubItems[0].Text,
                    //listView_order.Items[i].SubItems[1].Text,
                    strEuroricambi,
                    listView_order.Items[i].SubItems[2].Text,
                    listView_order.Items[i].SubItems[3].Text,
                    listView_order.Items[i].SubItems[4].Text);
            }
            else if (listView_order.Items[i].SubItems[0].Text.Length == 3)
            {
                strtoprint = string.Format("{0,-15}{1,-24}{2,-25}{3,-30}{4,-30}",
                    listView_order.Items[i].SubItems[0].Text,
                    //listView_order.Items[i].SubItems[1].Text,
                    strEuroricambi,
                    listView_order.Items[i].SubItems[2].Text,
                    listView_order.Items[i].SubItems[3].Text,
                    listView_order.Items[i].SubItems[4].Text);
            }
            return strtoprint;

        }
        protected void document_PrintPage(object sender, PrintPageEventArgs ev)
        {
            //количество строк с товарами для печати
            int rowtoprint = listView_order.Items.Count;
            string strtoprint = "";
            int postop = 0;

            Font printFont = new Font("Arial", 10);

            if (curprintpage == 0)//первая страница отчёта
            {
                //формируем картинку 
                int irow = 0;
                if (one_image == false)
                {
                    DataRow[] rows = dt_infobase.Select("num = '" + current_num_gb + "'");
                    irow = dt_infobase.Rows.IndexOf(rows[0]);
                }
                else
                {
                    DataRow[] rows = dt_infobase.Select("num = '0'");
                    irow = dt_infobase.Rows.IndexOf(rows[0]);
                }
                //int irow = dt_infobase.Rows.IndexOf(rows[0]);
                byte[] jpgBytes;
                if (one_image == false)
                {
                    jpgBytes = (byte[])dt_infobase.Rows[irow].ItemArray[2];//картинка для репорта в третьем поле
                }
                else
                {
                    jpgBytes = htmlBytes_one_print;
                }
                MemoryStream stream = new MemoryStream(jpgBytes, 0, jpgBytes.Length);

                var image = System.Drawing.Image.FromStream(stream);
                ev.Graphics.DrawImage(image, 130, 50, 509, 700);
                curprintpage += 1;

                postop = 800;//отступ сверху для названия коробки
                //формируем название коробки
                int r_dgv_gb = dataGridView_gearbox.CurrentCell.RowIndex;
                strtoprint = current_brandname + " - " + current_gearboxgroup + " - " +
                             dataGridView_gearbox[3, r_dgv_gb].Value.ToString() + " " +
                             dataGridView_gearbox[4, r_dgv_gb].Value.ToString();
                ev.Graphics.DrawString(strtoprint, printFont, Brushes.Black, 50, postop, new StringFormat());
                postop = postop + 30;

                //формируем заголовок списка                       //21 
                strtoprint = string.Format("{0,-11}{1,-20}{2,-19}{3,-49}{4,-30}", "позиция", "№ Eurorecambi", "№ оригинальный", "описание", "наличие");
                ev.Graphics.DrawString(strtoprint, printFont, Brushes.Black, 50, postop, new StringFormat());
                postop = postop + 20;

                //формируем список
                for (int i = last_row_to_print; i < rowtoprint; i++)
                {
                    if (postop < ev.MarginBounds.Height)
                    {
                        ev.Graphics.DrawString(listView_order_item_to_doc(i), printFont, Brushes.Black, 50, postop, new StringFormat());
                        postop = postop + 20;
                        last_row_to_print += 1;
                    }
                    else
                    {
                        ev.HasMorePages = true;
                        return;
                    }
                }
            }
            else
            {
                postop = 50;
                ev.HasMorePages = false;
                //формируем список
                for (int i = last_row_to_print; i < rowtoprint; i++)
                {
                    if (postop < ev.MarginBounds.Height)
                    {
                        ev.Graphics.DrawString(listView_order_item_to_doc(i), printFont, Brushes.Black, 50, postop, new StringFormat());
                        postop = postop + 20;
                        last_row_to_print += 1;
                        ev.HasMorePages = false;
                    }
                    else
                    {
                        ev.HasMorePages = true;
                        return;
                    }
                }
            }
        }

        private void button_csv_Click(object sender, EventArgs e)
        {
            string adtscode_to_csv = "";
            string file_path = "";

            if (listView_order.Items.Count > 0)
            {
                if (saveFileDialog_csv.ShowDialog() == DialogResult.OK)
                {
                    try
                    {


                        StreamWriter streamwriter = new System.IO.StreamWriter(saveFileDialog_csv.FileName,
                            false, System.Text.Encoding.GetEncoding("utf-8"));//("windows-1251"));// ("utf-8"));
                        //заголовки столбцов № Производителя, Производитель, Количество
                        //streamwriter.Write("№ Производителя; Производитель; Количество;;\n");
                        //проблема, если в кодировке windows-1251, то Excel не разбивает текст по столбцам

                        for (int i = 0; i < listView_order.Items.Count; i++)
                        {
                            //adtscode_to_csv = listView_order.Items[i].SubItems[3].Text + ";\n"; 
                            /*
                            //номер Eurorecambi, Eurorecambi, количество 1, код товара АДТС
                            adtscode_to_csv = listView_order.Items[i].SubItems[1].Text + ";Euroricambi;1;" +
                                              listView_order.Items[i].SubItems[3].Text + ";\n";
                            */
                            //оригинальный номер, производитель (current_brand), количество 1, код товара АДТС
                            /*
                            adtscode_to_csv = listView_order.Items[i].SubItems[2].Text + ";"+
                                              current_brandname.ToString()+";1;" +
                                              listView_order.Items[i].SubItems[3].Text + ";\n";
                            */

                            adtscode_to_csv = listView_order.Items[i].SubItems[1].Text + ";Euroricambi;1;" +
                                              listView_order.Items[i].SubItems[2].Text + ";" + 
                                              current_brandname.ToString() + ";1;" +
                                              listView_order.Items[i].SubItems[3].Text + ";\n";


                            streamwriter.Write(adtscode_to_csv);
                        }
                        streamwriter.Close();
                        file_path = saveFileDialog_csv.FileName;

                        //проба в Excel
                        //http://wladm.narod.ru/C_Sharp/comexcel.html#1
                        try
                        {
                            excelapp = new Excel.Application();
                            excelapp.Visible = true;

                            excelapp.Workbooks.OpenText(
                             file_path,
                             Excel.XlPlatform.xlWindows,
                             1,
                             Excel.XlTextParsingType.xlDelimited,
                             Excel.XlTextQualifier.xlTextQualifierDoubleQuote, //
                             true,          //Разделители одинарные
                             false,          //Разделители :Tab
                             true,         //Semicolon
                             false,         //Comma
                             false,         //Space
                             false,         //Other
                             Type.Missing,  //OtherChar
                             new object[] { new object[] { 1, Excel.XlColumnDataType.xlTextFormat },
                                        new object[] { 2, Excel.XlColumnDataType.xlTextFormat },
                                        new object[] { 3, Excel.XlColumnDataType.xlTextFormat },
                                        new object[] { 4, Excel.XlColumnDataType.xlTextFormat },
                                        new object[] { 5, Excel.XlColumnDataType.xlTextFormat },
                                        new object[] { 6, Excel.XlColumnDataType.xlTextFormat },
                                        new object[] { 7, Excel.XlColumnDataType.xlTextFormat }},
                             Type.Missing,
                             ".",
                             ",");
                            //сохраняем в формате xls
                            excelapp.DisplayAlerts = false;

                            excelapp.ActiveWorkbook.SaveAs(file_path.Replace(".csv", ".xls"), Excel.XlFileFormat.xlExcel8,//.xlExcel5,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange,
                                Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                            MessageBox.Show("Дополнительно к файлу в формате CSV был успешно создан файл в формате XLS","Создан файл в формате Excel",MessageBoxButtons.OK,MessageBoxIcon.Information);



                        }
                        catch
                        {
                            MessageBox.Show("Файл формата CSV успешно создан. На вашем компьютере не обнаружен MS Excel.","Внимание",MessageBoxButtons.OK,MessageBoxIcon.Information);  
                        }


                    }
                    catch
                    {
                        MessageBox.Show("Файл занят другим приложением. \nЗакройте это приложение и повторите попытку.","Ошибка импорта",MessageBoxButtons.OK,MessageBoxIcon.Error);  
                    }
                }
            }
        }

        private void button_zoop_plus_Click(object sender, EventArgs e)
        {
            webBrowser1.Focus();
            SendKeys.Send("^{+}");    
        }

        private void button_zoom_minus_Click(object sender, EventArgs e)
        {
            webBrowser1.Focus();
            SendKeys.Send("^{-}");
        }

        private void button_zoom_normal_Click(object sender, EventArgs e)
        {
            webBrowser1.Focus();
            SendKeys.Send("^0");
        }

        private void button_email_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = false;
            panel_reverse_search.Visible = false;
            panel_ws_login.Visible = false;
            if (listView_order.Items.Count > 0)
            {
                panel_email_to.Visible = true;
            }
                
        }

        private void button_email_to_Cancel_Click(object sender, EventArgs e)
        {
            panel_email_to.Visible = false;
        }

        private void button_email_to_OK_Click(object sender, EventArgs e)
        {
            //сохраняем контактные данные отправителя в файл
            //textBox_email_to_FIO
            //textBox_email_to_job_title
            //textBox_email_to_company
            //textBox_email_to_phone
            //textBox_email_to_phone_ext
            //textBox_email_to_EMAIL
            //textBox_email_to_EMAIL_domain

            //проверка, что все контактные данные заполнены, по крайней мере не пустые
            if (textBox_email_to_FIO.Text != "" &&
                textBox_email_to_job_title.Text != "" &&
                textBox_email_to_company.Text != "" &&
                textBox_email_to_phone.Text != "" &&
                //textBox_email_to_phone_ext.Text != "" &&
                textBox_email_to_EMAIL.Text != "" &&
                textBox_email_to_EMAIL_domain.Text != "")
            {
                if(textBox_email_to_phone_ext.Text == "" && pcd_omega == true)
                {
                    MessageBox.Show("Поле добавочного телефонного номера не заполнено", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                string fname = Application.StartupPath.ToString() + "\\pcd.txt";
                System.IO.StreamWriter f = new System.IO.StreamWriter(fname, false);
                f.WriteLine(textBox_email_to_FIO.Text);
                f.WriteLine(textBox_email_to_job_title.Text);
                f.WriteLine(textBox_email_to_company.Text);
                f.WriteLine(textBox_email_to_phone.Text);
                f.WriteLine(textBox_email_to_phone_ext.Text);
                f.WriteLine(textBox_email_to_EMAIL.Text);
                f.WriteLine(textBox_email_to_EMAIL_domain.Text);
                f.Close();
            }
            else
            {
                MessageBox.Show("Заполнены не все контактные данные", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            //проверка, что e-mail отправителя и пароль не пустые
            if(textBox_email_sender.Text.Length < 3 || textBox_email_sender_password.Text.Length < 1)
            {
                MessageBox.Show("Не настроена почта Google", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }


            //panel_email_to.Visible = false;
            //string to = textBox_email_to.Text;
            try
            {
                GMailUser = textBox_email_sender.Text;
                GMailPass = textBox_email_sender_password.Text;

                var to = new System.Net.Mail.MailAddress(textBox_email_to.Text);
                send_podbor(to.ToString());
                panel_email_to.Visible = false;
            }
            catch
            {
                MessageBox.Show("Недопустимый e-mail получателя", "", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }            
        }


        private void send_podbor(string to)
        {
            //send_email_ssl(to, Subject, TextBody);
            string strmsgbody = "Добрый день.\nНиже приведён список деталей, запрошенных Вами.\n\n";
            try
            {
                string msgstr = "Подбор деталей для " + current_brandname + " - " + current_gearboxgroup + " - " + current_description;

                for (int j=0; j< listView_order.Items.Count; j++)
                {
                    //strmsgbody = strmsgbody + listView_order.Items[j].SubItems[3].Text + "\n";
                    //номер Eurorecambi, Eurorecambi, количество 1, код товара АДТС
                    strmsgbody = strmsgbody + listView_order.Items[j].SubItems[1].Text + ";Euroricambi/KS;1;" +
                                              listView_order.Items[j].SubItems[2].Text + ";" +
                                              current_brandname.ToString() + ";1;" +
                                              listView_order.Items[j].SubItems[3].Text + ";\n";
                }
                strmsgbody = strmsgbody + "\nиспользуйте эти коды для заказа на http://www.****.ru \n";//\nКомментарий: ";
                strmsgbody = strmsgbody + "\nинструкция как заказать http://www.****.ru/Documents/help/cart_import_mobile_help_v1.1.pdf \n\nКомментарий: ";
                strmsgbody = strmsgbody + richTextBox_email_comment.Text;
                //добавляем IP адрес отправителя
                strmsgbody = strmsgbody + "\n\nотправлено с " + ip + "\n" + current_version + "\n\n";
                strmsgbody = strmsgbody + "С уважением,\n" + textBox_email_to_FIO.Text + "\n";
                strmsgbody = strmsgbody + textBox_email_to_job_title.Text + "\n";
                strmsgbody = strmsgbody + textBox_email_to_company.Text + "\n";
                strmsgbody = strmsgbody + "тел: " + textBox_email_to_phone.Text + " доб. " + textBox_email_to_phone_ext.Text + "\n";
                strmsgbody = strmsgbody + "e-mail: " + textBox_email_to_EMAIL.Text + "@"+ textBox_email_to_EMAIL_domain.Text + "\n";

                
                send_email_ssl(GMailUserOMEGA, GMailPassOMEGA, GMailUserOMEGA, msgstr, strmsgbody, false);
                send_email_ssl(GMailUser, GMailPass, to, msgstr, strmsgbody, true);

                if (email_error == false)
                {
                    MessageBox.Show("Результат подбора успешно отправлен\nна адрес: " + to, "Письмо отправлено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch// (Exception ex)
            {
           
            }
        }
        //http://stackoverflow.com/questions/3253701/get-public-external-ip-address


        





        private void button_update_Click(object sender, EventArgs e)
        {
            label_update_wait.Visible = true;
            string file_update_distr = "";
            string file_current = "";
            string file_backup = "";
            //DateTime localDate = DateTime.Now;
            string timebackup = "-" + DateTime.Now.Year.ToString() + "-" +
                                DateTime.Now.Month.ToString()+ "-" +
                                DateTime.Now.Day.ToString() + "-" +
                                DateTime.Now.Hour.ToString() + "-" +
                                DateTime.Now.Minute.ToString() + "-" +
                                DateTime.Now.Second.ToString();
            

            if(openFileDialog_update.ShowDialog() == DialogResult.OK)
            {
                
                

                try
                {
                    file_update_distr = openFileDialog_update.FileName.ToString();
                    file_current = Application.StartupPath.ToString() + "\\current.sch";
                    file_backup = Application.StartupPath.ToString() + "\\current" + timebackup + ".bac";
                    //бекап текущих схем
                    File.Copy(file_current, file_backup, true);
                    File.Copy(file_update_distr, file_current, true);
                    //проверка на доступность БД
                    try
                    {
                        SQLiteConnection concheck = new SQLiteConnection("Data Source=current.sch;Password=password;");
                        concheck.Open();
                        concheck.ChangePassword("password");
                        concheck.Close();
                    }
                    catch
                    {
                        SQLiteConnection concheck = new SQLiteConnection("Data Source=current.sch;");
                        concheck.Open();
                        concheck.ChangePassword("password");
                        concheck.Close();
                    }
                    label_update_wait.Visible = false;
                    //удаляем все бекапы, кроме последнего
                    try
                    {
                        FileInfo[] path = new DirectoryInfo(Application.StartupPath).GetFiles("*.bac", SearchOption.TopDirectoryOnly);
                        foreach (FileInfo file in path)
                        {
                            if(file.FullName != file_backup)
                            {
                                File.Delete(file.FullName);
                            }

                        }
                    }
                    catch { }

                    MessageBox.Show("Программа будет перезапущена", "Обновление схем завершено",MessageBoxButtons.OK,MessageBoxIcon.Information);
                    // MessageBox.Show("Обновление не удалось", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Restart();  
                }
                catch// (Exception ex)
                {
                    File.Copy(file_backup, file_current, true);
                    File.Delete(file_backup);
                    //удаление всех бекапов
                    try
                    {
                        FileInfo[] path = new DirectoryInfo(Application.StartupPath).GetFiles("*.bac", SearchOption.TopDirectoryOnly);
                        foreach(FileInfo file in path)
                        {
                            File.Delete(file.FullName);
                        }
                    }
                    catch { }

                    label_update_wait.Visible = false;
                    //MessageBox.Show("Обновление не удалось " + ex.ToString(), "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    //MessageBox.Show("Обновление не удалось\nпроверьте, что у Вас есть права\nлокального администратора на этот компьютер", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show("Обновление не удалось. Программа будет перезапущена.\n"+
                                    "Возможные причины:\n"+
                                    "* файл обновления испорчен -> скачайте файл снова и попробуйте ещё раз \n"+
                                    "* нет прав на изменение файлов в директории, где установлена программа\n"+
                                    "   -> получите права локального администратора на этот компьютер или\n"+
                                    "   -> переустановите программу в такое место, где эти права у Вас есть",
                                    "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.Restart();
                }


            }
            else
            {
                label_update_wait.Visible = false;
            }
        }

        private void button_searh_fig_on_map_Click(object sender, EventArgs e)
        {
            webBrowser1.Focus();
            SendKeys.Send("^f");
            SendKeys.Send("^F");
            SendKeys.Send("^А");
            SendKeys.Send("^а");
            SendKeys.Send("{BACKSPACE}");
            SendKeys.Send("{BACKSPACE}");
            SendKeys.Send("{BACKSPACE}");
            SendKeys.Send("{BACKSPACE}");
            SendKeys.Send("{BACKSPACE}");
            SendKeys.Send("{BACKSPACE}");

            SendKeys.Send(current_fig_linktable);//.ToString());
            SendKeys.Send("{TAB}");
            SendKeys.Send(" ");
           
        }

        private void Form1_SizeChanged(object sender, EventArgs e)
        {
            /*
            if (bigsize == false)
            {
                try
                {
                    webBrowser1.Width = Form1.ActiveForm.Width - 30;
                    webBrowser1.Height = Form1.ActiveForm.Height - 540;
                    bigsize = true;
                }
                catch { }
            }
            else
            {
                webBrowser1.Width = 1069;
                webBrowser1.Height = 359;
                bigsize = false;
            }
            */
            try
            {
                webBrowser1.Width = Form1.ActiveForm.Width - 40;//30
                webBrowser1.Height = Form1.ActiveForm.Height - 540;
                bigsize = true;
            }
            catch { }


        }

        private void button_feedback_Click(object sender, EventArgs e)
        {
            pictureBox1.Visible = false;
            if (panel_feedback.Visible == true)
            {
                panel_feedback.Visible = false;
                
            }
            else
            {
                panel_feedback.Visible = true;
                panel_reverse_search.Visible = false;
            }
        }

        private void button_feedback_Cancel_Click(object sender, EventArgs e)
        {
            panel_feedback.Visible = false;
        }

        private void send_email_ssl(string User, string Password, string To, string Subject, string TextBody, bool ShowError)
        {
            //https://www.emailarchitect.net/easendmail/ex/c/13.aspx
            //[C# - Send Email using Gmail Account over Implicit SSL on 465 Port]
            //https://www.emailarchitect.net/easendmail/kb/csharp.aspx?cat=2

            EASendMail.SmtpMail oMail = new EASendMail.SmtpMail("TryIt");
            EASendMail.SmtpClient oSmtp = new EASendMail.SmtpClient();
            oMail.To = To; 
            oMail.Subject = Subject;// "test email from gmail account";
            oMail.TextBody = TextBody;// "this is a test email sent from c# project with gmail.";
            EASendMail.SmtpServer oServer = new EASendMail.SmtpServer("smtp.gmail.com");
            oServer.Port = 465;
            oServer.ConnectType = EASendMail.SmtpConnectType.ConnectSSLAuto;
            oServer.User = User; 
            oServer.Password = Password;
            try
            {
                oSmtp.SendMail(oServer, oMail);
                email_error = false;
            }
            catch// (Exception ex)
            {
                if (ShowError == true)
                {
                    email_error = true;
                    //MessageBox.Show(ex.ToString());
                    MessageBox.Show("Проверьте:\n" +
                                    "* настроена ли почта Google\n" +
                                    "* есть ли соединение с интернетом\n" +
                                    "* не блокирует ли антивирус или Firewall 465 порт", "Ошибка отправки почты", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private void button_feedback_OK_Click(object sender, EventArgs e)
        {
            //textBox_feedback_name
            //textBox_feedback_email
            //textBox_feedback_phone
            //richTextBox_feedback_message
            try
            {

                string Subject = "Обратная связь от " + textBox_feedback_name.Text + ", e-mail: " + textBox_feedback_email.Text + ", тел: " + textBox_feedback_phone.Text;
                string TextBody = richTextBox_feedback_message.Text;

                if(richTextBox_feedback_message.Text == "Кто разработчик?")
                {
                    richTextBox_feedback_message.Text = "Котов Александр, e-mail: ak****k@gmail.com";
                    return;
                }
                if (TextBody.Length > 5) //сообщение должно быть не пустым
                {
                    TextBody = TextBody + "\n\n\nотправлено с " + ip + "\n" + current_version + "\n\nPCD:\n";
                    TextBody = TextBody + "Ф.И.О.: " + textBox_email_to_FIO.Text + "\n";
                    TextBody = TextBody + "Должность: " + textBox_email_to_job_title.Text + "\n";
                    TextBody = TextBody + "Компания: " + textBox_email_to_company.Text + "\n";
                    TextBody = TextBody + "тел.: " + textBox_email_to_phone.Text + " доб. " + textBox_email_to_phone_ext.Text + "\n";
                    TextBody = TextBody + "e-mail: " + textBox_email_to_EMAIL.Text + "@" + textBox_email_to_EMAIL_domain.Text;

                    send_email_ssl(GMailUser, GMailPass, GMailUserOMEGA, Subject, TextBody, true);
                    if (email_error == false)
                    {
                        MessageBox.Show("Благодарим за отзыв! \nВаши предложения обязательно будут рассмотрены.", "Сообщение отправлено", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                        panel_feedback.Visible = false;
                }
                else
                {
                    MessageBox.Show("Даже в слове Привет целых 6 букв :)", "Нечего отправлять");
                }

            }
            catch (Exception ex)// (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }

        private void button_ws_Click(object sender, EventArgs e)
        {
            if (listView_order.Items.Count > 0)
            {
                //авторизация
                panel_ws_login.Visible = true;
                panel_reverse_search.Visible = false;
                panel_email_to.Visible = false;
                // MessageBox.Show("Обращение к веб сервисам");
            }
        }

        private void button_ws_Cancel_Click(object sender, EventArgs e)
        {
            panel_ws_login.Visible = false;
            //textBox_ws_username.Text = "";
            //textBox_ws_password.Text = "";
        }

        private void button_ws_OK_Click(object sender, EventArgs e)
        {
            string PartNumber = "";
            string strmsgbody = "";
            int count_request = 0;

            //ещё раз защита от пустого результата подбора
            if (listView_order.Items.Count > 0)
            {
                //авторизация
                SecurityClient.SecurityClient client = new SecurityClient.SecurityClient();

                string hashSession = client.Logon(textBox_ws_username.Text, textBox_ws_password.Text);
                if (hashSession == "<root><error_message>Произошла ошибка: Object reference not set to an instance of an object.</error_message></root>")
                {
                    MessageBox.Show("Неправильное имя пользователя или пароль", "Ошибка авторизации", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
                //MessageBox.Show("Вы авторизованы на сайте как " + textBox_ws_username.Text);

                //цикл по строкам подбора
                label_ws_count.Visible = true;

                for (int j = 0; j < listView_order.Items.Count; j++)
                {
                    PartNumber = listView_order.Items[j].SubItems[1].Text;
                    if (PartNumber.Length > 5)//на всякий случай
                    {
                        if (check_exist(PartNumber, hashSession) == true)
                        {
                            listView_order.Items[j].SubItems[4].Text = "на складе";
                            listView_order.Items[j].UseItemStyleForSubItems = false;
                            listView_order.Items[j].SubItems[4].ForeColor = Color.Green;
                        }
                        else
                        {
                            listView_order.Items[j].SubItems[4].Text = "-";// "ожидается";
                            listView_order.Items[j].UseItemStyleForSubItems = false;
                            listView_order.Items[j].SubItems[4].ForeColor = Color.Red;
                        }
                        count_request += 1;
                        label_ws_count.Text = "Запрос " + count_request.ToString();// + " по Euroricambi";
                        label_ws_count.Update();
                    }
                }

                for (int j = 0; j < listView_order.Items.Count; j++)
                {
                    PartNumber = listView_order.Items[j].SubItems[2].Text;
                    if (PartNumber.Length > 5)//на всякий случай
                    {
                        if (check_exist(PartNumber, hashSession) == true)
                        {
                            listView_order.Items[j].SubItems[4].Text = "на складе";
                            listView_order.Items[j].UseItemStyleForSubItems = false;
                            listView_order.Items[j].SubItems[4].ForeColor = Color.Green;
                        }
                        
                        else if(listView_order.Items[j].SubItems[4].Text != "на складе")
                        {
                            listView_order.Items[j].SubItems[4].Text = "-";// "ожидается";
                            listView_order.Items[j].UseItemStyleForSubItems = false;
                            listView_order.Items[j].SubItems[4].ForeColor = Color.Red;
                        }
                        
                        count_request += 1;
                        label_ws_count.Text = "Запрос " + count_request.ToString();// + " по " + current_brandname;
                        label_ws_count.Update();
                        panel_ws_login.Refresh();
                    }
                }

                //отсылаем на на нашу почту
                //send_email_ssl(to, Subject, TextBody);
                try
                {
                    string msgstr = "Обращение к web сервисам под логином: " + textBox_ws_username.Text;
                    for (int j = 0; j < listView_order.Items.Count; j++)
                    {
                        //strmsgbody = strmsgbody + listView_order.Items[j].SubItems[3].Text + "\n";
                        //номер Eurorecambi, Eurorecambi, количество 1, код товара АДТС
                        strmsgbody = strmsgbody + listView_order.Items[j].SubItems[1].Text + " - Euroricambi/KS; " +
                                                  listView_order.Items[j].SubItems[2].Text + " - " +
                                                  current_brandname.ToString() + "; " +
                                                  //listView_order.Items[j].SubItems[3].Text + " - код товара в АДТС; " +
                                                  listView_order.Items[j].SubItems[4].Text + "\n";//" - Euroricambi; " +
                                                  //listView_order.Items[j].SubItems[5].Text + " - " + current_brandname + "\n";
                    }
                    // msg.Body = strmsgbody + "\n\n\nотправлено с " + ip;
                    strmsgbody = strmsgbody + "\n\n\nотправлено с " + ip + "\n"+ current_version + "\n\nPCD:\n";
                    strmsgbody = strmsgbody + "Ф.И.О.: " + textBox_email_to_FIO.Text + "\n";
                    strmsgbody = strmsgbody + "Должность: " + textBox_email_to_job_title.Text + "\n";
                    strmsgbody = strmsgbody + "Компания: " + textBox_email_to_company.Text + "\n";
                    strmsgbody = strmsgbody + "тел.: " + textBox_email_to_phone.Text + " доб. " + textBox_email_to_phone_ext.Text + "\n";
                    strmsgbody = strmsgbody + "e-mail: " + textBox_email_to_EMAIL.Text + "@" + textBox_email_to_EMAIL_domain.Text;


                    send_email_ssl(GMailUserOMEGA, GMailPassOMEGA, GMailUserOMEGA, msgstr, strmsgbody, false);
                }
                catch { }
                label_ws_count.Visible = false;
                panel_ws_login.Visible = false;
                //textBox_ws_username.Text = "";
                //textBox_ws_password.Text = "";
            }
        }

        private bool check_exist(string PartNumber, string hashSession)
        {
            // Получаем список деталей с признаком наличия на складе
            SearchClient.SearchClient client2 = new SearchClient.SearchClient();
            string results = client2.SearchBasic(PartNumber, hashSession);
            XDocument xdoc = XDocument.Parse(results);

            // Ведем поиск номера в полученной информации (внимание, номера в полученной информации и искомый номер требуют нормализации)
            if (xdoc.Descendants("part").Count() > 0)
                foreach (var res in xdoc.Descendants("part"))
                    //if (res.Element("unique_number").Value == PartNumber && res.Element("is_sklad").Value == "1")
                    if (res.Element("is_sklad").Value == "1")
                        return true;
            return false;
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            string fname = Application.StartupPath.ToString() + "\\l.****";
            string fname_gmail = Application.StartupPath.ToString() + "\\l.gmail";

            if (checkBox_save_login.Checked == true)
            {
                System.IO.StreamWriter f = new System.IO.StreamWriter(fname,false);
                f.WriteLine(textBox_ws_username.Text);
                f.WriteLine(textBox_ws_password.Text);
                f.Close();
            }
            else
            {
                try
                {
                    File.Delete(fname);
                }
                catch { }
            }

            if(checkBox_email_sender_save_login.Checked == true)
            {
                System.IO.StreamWriter fg = new System.IO.StreamWriter(fname_gmail, false);
                fg.WriteLine(textBox_email_sender.Text);
                fg.WriteLine(textBox_email_sender_password.Text);
                fg.Close();
            }
            else
            {
                try
                {
                    File.Delete(fname_gmail);
                }
                catch { }
            }

        }

        private void check_login_file()
        {
            string fname = Application.StartupPath.ToString() + "\\l.****";
            if (File.Exists(fname))
            {
               // MessageBox.Show("файл логина найден");
                string[] lines = System.IO.File.ReadAllLines(fname);
                textBox_ws_username.Text = lines[0];
                textBox_ws_password.Text = lines[1];
                checkBox_save_login.Checked = true;
            }
        }

        private void check_gmail_login_file()
        {
            string fname_gmail = Application.StartupPath.ToString() + "\\l.gmail";
            if (File.Exists(fname_gmail))
            {
                string[] lines = System.IO.File.ReadAllLines(fname_gmail);
                textBox_email_sender.Text = lines[0];
                textBox_email_sender_password.Text = lines[1];
                checkBox_email_sender_save_login.Checked = true;
            }
        }

        private void check_pcd_file()
        {
            string fname = Application.StartupPath.ToString() + "\\pcd.txt";
            if (File.Exists(fname))
            {
                // MessageBox.Show("файл с контактными данными найден");
                string[] lines = System.IO.File.ReadAllLines(fname);
                textBox_email_to_FIO.Text = lines[0];
                textBox_email_to_job_title.Text = lines[1];
                textBox_email_to_company.Text = lines[2];
                textBox_email_to_phone.Text = lines[3];
                textBox_email_to_phone_ext.Text = lines[4];
                textBox_email_to_EMAIL.Text = lines[5];
                textBox_email_to_EMAIL_domain.Text = lines[6];
            }

        }

        //http://stackoverflow.com/questions/6620165/how-can-i-parse-json-with-c
        private void check_config_file()
        {
            string fname = Application.StartupPath.ToString() + "\\config.json";
            if (File.Exists(fname))
            {
                //MessageBox.Show("Файл конфигурации config.json найден");
                string line = System.IO.File.ReadAllText(fname);
                JsonConfig deserializedJson = JsonConvert.DeserializeObject<JsonConfig>(line);

                if (deserializedJson.SplashText != "")
                {
                    current_version = deserializedJson.SplashText;
                }

                if (deserializedJson.ToolTipButtonHelp != "")
                {
                    //MessageBox.Show(deserializedJson.ToolTipButtonHelp);
                    toolTipJSON.SetToolTip(button_help, deserializedJson.ToolTipButtonHelp);
                }

                if(deserializedJson.ToolTipButtonReverseSearch != "")
                {
                    toolTipJSON.SetToolTip(button_reverse_search, deserializedJson.ToolTipButtonReverseSearch);
                }

                if(deserializedJson.ToolTipButtonWs != "")
                {
                    toolTipJSON.SetToolTip(button_ws, deserializedJson.ToolTipButtonWs);
                }

                if (deserializedJson.ToolTipListViewOrder != "")
                {
                    toolTipJSON.SetToolTip(listView_order, deserializedJson.ToolTipListViewOrder);
                }

                if (deserializedJson.FormText != "")
                {
                    this.Text = deserializedJson.FormText;
                }

                if(deserializedJson.LabelReverseSearchInfo != "")
                {
                    label_reverse_search_info.Text = deserializedJson.LabelReverseSearchInfo;
                }

                if(deserializedJson.GroupBoxWsLogin != "")
                {
                    groupBox_ws_login.Text = deserializedJson.GroupBoxWsLogin;
                }

                if(deserializedJson.HeaderOurNo != "")
                {
                    header_ourno = deserializedJson.HeaderOurNo;
                    //listView_order.Columns[1].Text = header_ourno;
                }

                if(deserializedJson.HeaderPartNo != "")
                {
                    header_partno = deserializedJson.HeaderPartNo;
                }

                if(deserializedJson.HeaderItemCod != "")
                {
                    header_itemcod = deserializedJson.HeaderItemCod;
                }

                if(deserializedJson.HeaderType != "")
                {
                    header_type = deserializedJson.HeaderType;
                }

                if (deserializedJson.Olga == "1")
                {
                    copyCellData = true;
                    button_csv_visible = true;
                    pcd_omega = true;
                }

                if (deserializedJson.softutc == "1")
                {
                    softutc = true;
                }
            }
        }


        private void textBox_ws_password_KeyUp(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {               
                button_ws_OK.Focus();
                panel_ws_login.Refresh();
                button_ws_OK.PerformClick();
            }
        }

        private void button_reverse_search_Click(object sender, EventArgs e)
        {
            if (panel_reverse_search.Visible == false)
            {
                //textBox_reverse_search.Text = "введите номер детали";
                panel_reverse_search.Visible = true;
                pictureBox1.Visible = false;
                panel_email_to.Visible = false;
                panel_feedback.Visible = false;
            }
            else
            {
                panel_reverse_search.Visible = false;
            }
        }

        private void button_reverse_search_Cancel_Click(object sender, EventArgs e)
        {
            panel_reverse_search.Visible = false;
        }

        private void textBox_reverse_search_Enter(object sender, EventArgs e)
        {
            //textBox_reverse_search.Text = "";
        }

        private void button_test_Click(object sender, EventArgs e)
        {
            if (textBox_reverse_search.Text != " " && textBox_reverse_search.Text.Length > 3)
            {
                reverse_search(textBox_reverse_search.Text);// ("1 315 202 055");
            }
        }
        private void reverse_search(string reverse_any_no)
        {

            //http://www.dotnetperls.com/datatable-select

            //string reverse_any_no = "1 315 202 055";//"95535210"; //"1 315 202 055";
            //string addtmp = "";

            try
            {
                dt_reverse.Rows.Clear();
            }
            catch { }
            int reverse_count = 0;
            int rcn = 0;

            DataRow[] ds = dt_linktable.Select("ourno = '" + reverse_any_no + "' OR partno = '" + reverse_any_no + "'");
            DataRow[] dsgb = null;

            //количество найденных схем и их номера
            reverse_count = ds.Length;
            label_found.Text = "моделей: " + reverse_count.ToString();
            //reverse_nums = new int[reverse_count];
            for (int i = 0; i < reverse_count; i++)
            {
                rcn = Int32.Parse(ds[i][0].ToString());
                //dt_reverse.Rows.Add(ds[i][0], "ZF", "GEARBOX: 12 AS 1210 TO (1336 033 xxx)", "10,37-0,81 R. 10,56", "AS Tronic mid Truck (12 speeds)");
                dsgb = dt_gearbox.Select("num = '" + rcn.ToString() + "'");
                dt_reverse.Rows.Add(rcn, dsgb[0][2].ToString(), dsgb[0][5].ToString(), dsgb[0][3].ToString(), dsgb[0][4].ToString());
            }

            dataGridView_reverse_search.DataSource = dt_reverse;

            //грид dataGridView_reverse_search
            dataGridView_reverse_search.Columns[0].HeaderText = "№ схемы";
            dataGridView_reverse_search.Columns[0].Width = 78;
            dataGridView_reverse_search.Columns[1].HeaderText = "производитель";
            dataGridView_reverse_search.Columns[1].Width = 100;
            dataGridView_reverse_search.Columns[2].HeaderText = "семейство (группа)";
            dataGridView_reverse_search.Columns[2].Width = 180;
            dataGridView_reverse_search.Columns[3].HeaderText = "модель";
            dataGridView_reverse_search.Columns[3].Width = 265;
            dataGridView_reverse_search.Columns[4].HeaderText = " ";
            dataGridView_reverse_search.Columns[4].Width = 215;

        }

        private void textBox_reverse_search_TextChanged(object sender, EventArgs e)
        {
            string s = "";
            string snorm_ZF_xxxx_xxx_xx = ""; //нормализация ZF
            string snorm_Renault_xx_xx_xxx_xx = "";//нормализация Renault
            string snorm_MAN_0x_xxxxx_xxxx = "";//нормализация MAN при длине 10
            string snorm_MAN_dot_xxx_xxx_xxxx = "";//нормализация MAN при длине 11 и первом символе точке
            string snorm_MAN_xx_xxxxx_xxxx = "";//нормализация MAN при длине 11
            string snorm_MAN_S_xx_xxxxx_xxxx = "";//нормализация MAN при длине 12 и первом символе S
            string snorm_DAF_0xx_xxx = "";//нормализация DAF при длине 5
            string snorm_DAF_xxx_xxx = "";//нормализация DAF при длине 6
            string snorm_DAF_xxxx_xxx = "";//нормализация DAF при длине 7

            //bool normalization = false;
            s = textBox_reverse_search.Text;

            if (s != " " && s.Length > 3)
            {
                if(s.Length == 5)
                {
                    snorm_DAF_0xx_xxx = "0" + s.Substring(0, 2) + " " + s.Substring(2, 3);
                }
                else if (s.Length == 6)
                {
                    snorm_DAF_xxx_xxx = s.Substring(0, 3) + " " + s.Substring(3, 3);
                }
                else if (s.Length == 7)
                {
                    snorm_DAF_xxxx_xxx = s.Substring(0, 4) + " " + s.Substring(4, 3);
                }

                else if (s.Length == 10)
                {
                    snorm_ZF_xxxx_xxx_xx = s.Substring(0,4) + " " + s.Substring(4, 3) + " " + s.Substring(7,3);
                    snorm_Renault_xx_xx_xxx_xx = s.Substring(0, 2) + " " + s.Substring(2, 2) + " " + s.Substring(4, 3) + " " + s.Substring(7, 3);
                    snorm_MAN_0x_xxxxx_xxxx = "0" + s.Substring(0, 1) + " " + s.Substring(1, 5) + " " + s.Substring(6,4);
                }
                else if(s.Length == 11 && s.Substring(0,1) == ".")
                {
                    snorm_MAN_dot_xxx_xxx_xxxx = ". " + s.Substring(1, 3) + " " + s.Substring(4, 3) + " " + s.Substring(7,4);
                }
                else if (s.Length == 11 && s.Substring(0, 1) != ".")
                {
                    snorm_MAN_xx_xxxxx_xxxx = s.Substring(0, 2) + " " + s.Substring(2, 5) + " " + s.Substring(7, 4);
                }
                else if(s.Length == 12 && s.Substring(0, 1) == "S")
                {
                    snorm_MAN_S_xx_xxxxx_xxxx = "S " + s.Substring(1, 2) + " " + s.Substring(3, 5) + " " + s.Substring(8, 4);
                }

                try
                {
                    normal_search_string = s;
                    reverse_search(s);// ("1 315 202 055");
                }
                catch { }               

                if (dataGridView_reverse_search.Rows.Count > 1)
                {
                    //normalization = false;
                    return;
                }

                //normalization = true;
                button_reverse_search_OK.Text = "Перейти к схеме ???";
                if (s.Length == 10)
                {
                    //попытка по ZF
                    try
                    {
                        normal_search_string = snorm_ZF_xxxx_xxx_xx;
                        reverse_search(snorm_ZF_xxxx_xxx_xx);
                    }
                    catch { }
                    if (dataGridView_reverse_search.Rows.Count == 1)
                    {
                        //попытка по Renault
                        try
                        {
                            normal_search_string = snorm_Renault_xx_xx_xxx_xx;
                            reverse_search(snorm_Renault_xx_xx_xxx_xx);
                        }
                        catch { }
                    }
                    if (dataGridView_reverse_search.Rows.Count == 1)
                    {
                        //попытка по MAN с длиной 10
                        try
                        {
                            normal_search_string = snorm_MAN_0x_xxxxx_xxxx;
                            reverse_search(snorm_MAN_0x_xxxxx_xxxx);
                        }
                        catch { }
                    }

                }

                //проверка
                if (dataGridView_reverse_search.Rows.Count > 1)
                {
                    return;
                }
                //нормализация MAN при длине 11 и первом символе точке
                button_reverse_search_OK.Text = "Перейти к схеме ???";
                if (s.Length == 11 && s.Substring(0, 1) == ".")
                {
                    try
                    {
                        normal_search_string = snorm_MAN_dot_xxx_xxx_xxxx;
                        reverse_search(snorm_MAN_dot_xxx_xxx_xxxx);
                    }
                    catch { }
                }

                //проверка
                if (dataGridView_reverse_search.Rows.Count > 1)
                {
                    return;
                }
                //нормализация MAN при длине 11 и первом символе НЕ точке
                button_reverse_search_OK.Text = "Перейти к схеме ???";
                if (s.Length == 11 && s.Substring(0, 1) != ".")
                {
                    try
                    {
                        normal_search_string = snorm_MAN_xx_xxxxx_xxxx;
                        reverse_search(snorm_MAN_xx_xxxxx_xxxx);
                    }
                    catch { }
                }

                //проверка
                if (dataGridView_reverse_search.Rows.Count > 1)
                {
                    return;
                }
                //нормализация MAN при длине 12 и первом символе S
                button_reverse_search_OK.Text = "Перейти к схеме ???";
                if (s.Length == 12 && s.Substring(0, 1) == "S")
                {
                    try
                    {
                        normal_search_string = snorm_MAN_S_xx_xxxxx_xxxx;
                        reverse_search(snorm_MAN_S_xx_xxxxx_xxxx);
                    }
                    catch { }
                }

                //проверка
                if (dataGridView_reverse_search.Rows.Count > 1)
                {
                    return;
                }
                //нормализация DAF при длине 5
                button_reverse_search_OK.Text = "Перейти к схеме ???";
                if (s.Length == 5)
                {
                    try
                    {
                        normal_search_string = snorm_DAF_0xx_xxx;
                        reverse_search(snorm_DAF_0xx_xxx);
                    }
                    catch { }
                }

                //проверка
                if (dataGridView_reverse_search.Rows.Count > 1)
                {
                    return;
                }
                //нормализация DAF при длине 6
                button_reverse_search_OK.Text = "Перейти к схеме ???";
                if (s.Length == 6)
                {
                    try
                    {
                        normal_search_string = snorm_DAF_xxx_xxx;
                        reverse_search(snorm_DAF_xxx_xxx);
                    }
                    catch { }
                }

                //проверка
                if (dataGridView_reverse_search.Rows.Count > 1)
                {
                    return;
                }
                //нормализация DAF при длине 7
                button_reverse_search_OK.Text = "Перейти к схеме ???";
                if (s.Length == 7)
                {
                    try
                    {
                        normal_search_string = snorm_DAF_xxxx_xxx;
                        reverse_search(snorm_DAF_xxxx_xxx);
                    }
                    catch { }
                }


            }
        }

        private void dataGridView_reverse_search_SelectionChanged(object sender, EventArgs e)
        {
            int c = dataGridView_reverse_search.CurrentCell.ColumnIndex;
            int r = dataGridView_reverse_search.CurrentCell.RowIndex;
            //int current_num_gb_reverse = 0;
            //current_brandname
            //current_gearboxgroup

            int rowcount = dataGridView_reverse_search.Rows.Count; //RowCount;
            if (r < rowcount -1 && rowcount > 0)
            {
                current_num_gb_reverse = Int32.Parse((dataGridView_reverse_search[0, r].Value.ToString()));
                current_brandname_reverse = dataGridView_reverse_search[1, r].Value.ToString();
                current_gearboxgroup_reverse = dataGridView_reverse_search[2, r].Value.ToString();

                //Перейти к схеме
                button_reverse_search_OK.Text = "Перейти к схеме " + current_num_gb_reverse.ToString();
            }

        }

        private void button_reverse_search_OK_Click(object sender, EventArgs e)
        {
            //dataGridView_gearbox_select[1, r];
            //current_num_gb = current_num_gb_reverse;
            //current_brandname = current_brandname_reverse;
            //current_gearboxgroup = current_gearboxgroup_reverse;

            if (current_gearboxgroup_reverse != "0")
            {
                //выделяем нужную строку в dataGridView_gearbox_select
                foreach (DataGridViewRow r in dataGridView_gearbox_select.Rows)
                    if ((string)r.Cells["gearboxgroup"].Value == current_gearboxgroup_reverse)
                    {
                        dataGridView_gearbox_select.Rows[r.Index].Selected = true;
                        break;
                    }

                //фильтруем dataGridView_gearbox
                bindingSource_gearbox.Filter = "brandname = '" + current_brandname_reverse + "' AND gearboxgroup = '" + current_gearboxgroup_reverse + "'";

                //выделяем нужную строку в dataGridView_gearbox
                foreach (DataGridViewRow rs in dataGridView_gearbox.Rows)
                    //if ((Int32)rs.Cells["num"].Value == current_num_gb_reverse)
                    if (Int32.Parse(rs.Cells["num"].Value.ToString()) == current_num_gb_reverse)
                    {

                        dataGridView_gearbox.CurrentCell = dataGridView_gearbox[3, rs.Index];
                        dataGridView_gearbox.Rows[rs.Index].Selected = true;
                        break;
                    }

                //getmap();

                listView_order.Items.Clear();
                //rcn = Int32.Parse(ds[i][0].ToString());

                //по выбранной схеме и номеру Euroricambi или оригинальному вернуть номер позиции из dt_linktable

                DataRow[] dsourno = dt_linktable.Select("num = '" + current_num_gb_reverse + "' AND ourno = '" +
                                      normal_search_string + "'");
                DataRow[] dspartno = dt_linktable.Select("num = '" + current_num_gb_reverse + "' AND partno = '" +
                                      normal_search_string + "'");
                try
                {
                    if (dsourno.Count() > 0)
                    {
                        //current_fig_reverse = (Int32)dsourno[0].ItemArray[1];
                        current_fig_reverse = Int32.Parse(dsourno[0].ItemArray[1].ToString());
                        //MessageBox.Show("по Euroricambi, позиция "+current_fig_reverse.ToString());
                    }
                    else
                    {
                        //current_fig_reverse = (Int32)dspartno[0].ItemArray[1];
                        current_fig_reverse = Int32.Parse(dspartno[0].ItemArray[1].ToString());
                        //MessageBox.Show("по оригинальному, позиция " + current_fig_reverse.ToString());
                    }
                    panel_reverse_search.Visible = false;
                }
                catch { }// { MessageBox.Show("Ошибка с current_fig_reverse"); }

                //выделяем строку в dataGridView_linktable
                foreach (DataGridViewRow rs2 in dataGridView_linktable.Rows)
                    //if ((Int32)rs.Cells["num"].Value == current_num_gb_reverse)
                    if (Int32.Parse(rs2.Cells["fig"].Value.ToString()) == current_fig_reverse)
                    {

                        dataGridView_linktable.CurrentCell = dataGridView_linktable[1, rs2.Index];
                        dataGridView_linktable.Rows[rs2.Index].Selected = true;
                        break;
                    }

            }

        }

        void select_fig_after_reverse()
        {
            //string s = "";
            //current_fig_reverse - номер позиции известен из 
            //private void dataGridView_reverse_search_SelectionChanged(object sender, EventArgs e)
            if(current_fig_reverse > 0)
            {
                /*
                s = current_fig_reverse.ToString();
                MessageBox.Show("Здесь будет подсветка позиции " + current_fig_reverse.ToString());
                mshtml.IHTMLDocument2 document = webBrowser1.Document.DomDocument as IHTMLDocument2;
                IHTMLSelectionObject currentSelection = document.selection;
                IHTMLTxtRange range = currentSelection.createRange() as IHTMLTxtRange;

                if (range.findText(s, s.Length, 2))
                {
                    range.select();
                }
                */
                webBrowser1.Focus();
                SendKeys.Send("^f");
                SendKeys.Send("^F");
                SendKeys.Send("^А");
                SendKeys.Send("^а");
                SendKeys.Send("{BACKSPACE}");
                SendKeys.Send("{BACKSPACE}");
                SendKeys.Send(current_fig_reverse.ToString());
                SendKeys.Send("{TAB}");
                SendKeys.Send(" ");


                //после подсветки скидываем в ноль
                current_fig_reverse = 0;
            }
        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            select_fig_after_reverse();
        }

        private void dataGridView_gearbox_select_MouseClick(object sender, MouseEventArgs e)
        {
            textBox_filter_series.Text = "";
            clear_background = true;
        }

        private void checkBox_All_models_CheckedChanged(object sender, EventArgs e)
        {
            string n = "";
            if (checkBox_All_models.Checked == false)
             {
                //номер текущей схемы берем из dataGridView_linktable
                n = dataGridView_linktable[0, 0].FormattedValue.ToString();
                //dataGridView_gearbox_select.Visible = true;
                panel_no_groups.Visible = false;
                bindingSource_gearbox.Filter = "brandname = '" + current_brandname + "' AND gearboxgroup = '" + current_gearboxgroup + "'";
                //выбор правильной строки в dataGridView_gearbox_select
                //http://skillcoding.com/Default.aspx?id=151

                string group = "";
                group = label_no_groups_group.Text;
                for (int i=0; i < dataGridView_gearbox_select.RowCount; i++)
                    if(dataGridView_gearbox_select[1,i].FormattedValue.ToString().Contains(group))
                    {
                        dataGridView_gearbox_select.CurrentCell = dataGridView_gearbox_select[1, i];
                        break;
                    }
                //текущую схемы берем из dataGridView_linktable
                //http://www.cyberforum.ru/windows-forms/thread435177.html
                bindingSource_gearbox.Position = bindingSource_gearbox.Find("num",n);
            }
            else
             {
                //номер текущей схемы берем из dataGridView_linktable
                n = dataGridView_linktable[0, 0].FormattedValue.ToString();
                bindingSource_gearbox.Filter = "";
                //dataGridView_gearbox_select.Visible = false;
                panel_no_groups.Visible = true;
                dataGridView_gearbox.Focus();
                bindingSource_gearbox.Position = bindingSource_gearbox.Find("num", n);
            }
        }

        private void listView_order_MouseDoubleClick(object sender, MouseEventArgs e)
        {

            ListViewHitTestInfo hit = listView_order.HitTest(e.Location);
            string selectedn = "";
            foreach (ListViewItem item in listView_order.SelectedItems)
            {
                if (radioButton_type_partno.Checked == true)
                {
                    selectedn = item.SubItems[2].Text;
                    if (selectedn == "*****" || selectedn.Length < 3)
                    {
                        //selectedn = item.SubItems[1].Text;
                        MessageBox.Show("Номер не определён, попробуйте по " + header_ourno);
                        return;
                    }
                }
                else
                {
                    selectedn = item.SubItems[1].Text;
                    if (selectedn.Length < 3)
                    {
                        MessageBox.Show("Номер не определён, попробуйте по " + header_partno);
                        return;
                    }

                }
            }
            //selectedn = listView_order.SelectedItems.ToString;
            //MessageBox.Show(selectedn);
            //http://www.****.ru/search.aspx?text=95570722
            try
            {
                System.Diagnostics.Process.Start("http://www.****.ru/search.aspx?text=" + selectedn); 
            }
            catch { }
            

            //https://habrahabr.ru/post/170015/
            //webControl1.Show();
           
        }


        private void button_ws_show_password_MouseDown(object sender, MouseEventArgs e)
        {
            textBox_ws_password.PasswordChar = '\0';
            //imageList_password
            button_ws_show_password.Image = imageList_password.Images[1];
        }

        private void button_ws_show_password_MouseUp(object sender, MouseEventArgs e)
        {
            textBox_ws_password.PasswordChar = '*';
            button_ws_show_password.Image = imageList_password.Images[0];
            textBox_ws_password.Focus();
        }


        private void dataGridView_linktable_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            int currow = dataGridView_linktable.CurrentCell.RowIndex;
            current_fig_linktable = dataGridView_linktable[1, currow].Value.ToString();
            if (no_notes_on_sorting == false)
            {
                if (dataGridView_linktable.CurrentCell.ColumnIndex == 6 && dataGridView_linktable.CurrentCell.Value.ToString().Length > 0)
                {
                    //MessageBox.Show("заметка");
                    buttonNotes.PerformClick();
                }
            }
            else
            {
                no_notes_on_sorting = false;
            }
        }

        private void dataGridView_linktable_CellContentDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            int currow = dataGridView_linktable.CurrentCell.RowIndex;
            string fig = dataGridView_linktable[1, currow].Value.ToString();
            current_fig_linktable = fig;
            string ourno = dataGridView_linktable[2, currow].Value.ToString();
            string partno = dataGridView_linktable[3, currow].Value.ToString();
            string description = dataGridView_linktable[5, currow].Value.ToString();

            listView_order.BeginUpdate();
            ListViewItem item = new ListViewItem(fig);
            item.SubItems.Add(ourno);
            item.SubItems.Add(partno);

            item.SubItems.Add(description);
            item.SubItems.Add(" ");//пустое поле для наличия ///Euroricambi
                                   //item.SubItems.Add(" ");//пустое поле для наличия оригинального
            listView_order.Items.AddRange(new ListViewItem[] { item, });
            listView_order.Columns[0].Width = 70;
            listView_order.Columns[1].Width = 130;
            listView_order.Columns[2].Width = 130;
            //listView_order.Columns[3].Width = 140;
            listView_order.Columns[3].Width = 540;
            listView_order.Columns[4].Width = 115;
            //listView_order.Columns[5].Width = 115;
            listView_order.Items[listView_order.Items.Count - 1].EnsureVisible();

            listView_order.EndUpdate();

        }

        private void dataGridView_linktable_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if(copyCellData == true)
            {
                try
                {
                    Clipboard.SetText(dataGridView_linktable.CurrentCell.Value.ToString());
                }
                catch { }
            }
            

        }

        private void dataGridView_linktable_ColumnHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            if (clear_background == false)
            {
                dataGridView_linktable_cell_background();
            }
            no_notes_on_sorting = true;
            panelNotes.Visible = false;
        }

        //http://alexeyworld.com/blog/mouseeventhandler_rightclick.36.aspx
        void listView_order_right_click(object sender, MouseEventArgs e)
        {
            //button_csv_visible
            if (button_csv_visible == true)
            {
                if (e.Button == MouseButtons.Right)
                {
                    if (button_csv.Visible == false)
                    {
                        button_csv.Visible = true;
                    }
                    else
                    {
                        button_csv.Visible = false;
                    }
                }
                else if (e.Button == MouseButtons.Left)
                {
                    button_csv.Visible = false;
                }
            }
        }

        private void buttonNotes_Click(object sender, EventArgs e)
        {
            if(panelNotes.Visible == false && current_notes !="")
            {
                richTextBoxNotes.Text = current_notes;
                pictureBox1.Visible = false;
                panelNotes.Visible = true;
            }
            else
            {
                panelNotes.Visible = false;
            }
        }

        private void button_email_sender_show_password_MouseDown(object sender, MouseEventArgs e)
        {
            textBox_email_sender_password.PasswordChar = '\0';
            button_email_sender_show_password.Image = imageList_password.Images[1];
        }

        private void button_email_sender_show_password_MouseUp(object sender, MouseEventArgs e)
        {
            textBox_email_sender_password.PasswordChar = '*';
            button_email_sender_show_password.Image = imageList_password.Images[0];
            textBox_email_sender_password.Focus();
        }

        private void linkLabel_new_Gmail_Click(object sender, EventArgs e)
        {
            try
            {
                System.Diagnostics.Process.Start("https://accounts.google.com/signup"); //
            }
            catch { }
        }
    }
}
