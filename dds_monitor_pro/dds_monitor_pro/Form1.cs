using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Odbc;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using static dds_monitor_pro.Program;
namespace dds_monitor_pro
{
    public partial class Form1 : Form
    {
        public static DataTable log_data;
        public static DataTable log_data_no_access;
        public static DataTable log_data_system;
        DataSet dr_card_log = new DataSet();
        DataSet not_return_card = new DataSet();

        public static Boolean is_update_photo = false;
        public static string curr_photo_path = "";


        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            CultureInfo culture = (CultureInfo)CultureInfo.CurrentCulture.Clone();
            culture.DateTimeFormat.ShortDatePattern = @"yyyy-MM-dd";
            culture.DateTimeFormat.LongTimePattern = @"HH:mm:ss";
            Thread.CurrentThread.CurrentCulture = culture;

            // pictureBox1.ImageLocation = @"C:\Program Files (x86)\Amadeus5\Media\Portraits\22673.jpg";
            log_data = new DataTable();
            log_data.Clear();
            log_data.Columns.Add("Date");
            log_data.Columns.Add("Description");
            log_data.Columns.Add("Photo");
            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.DataSource = log_data;
            dataGridView1.Refresh();


            log_data_no_access = new DataTable();
            log_data_no_access.Clear();
            log_data_no_access.Columns.Add("Date");
            log_data_no_access.Columns.Add("Description");
            log_data_no_access.Columns.Add("Photo");

            dataGridView2.AutoGenerateColumns = false;
            dataGridView2.DataSource = log_data_no_access;
            dataGridView2.Refresh();


            log_data_system = new DataTable();
            log_data_system.Clear();
            log_data_system.Columns.Add("Date");
            log_data_system.Columns.Add("Description");

            dataGridView3.AutoGenerateColumns = false;
            dataGridView3.DataSource = log_data_system;
            dataGridView3.Refresh();
            this.Text = this.Text + " Ver: " + tversion;

            string c_datetime = get_server_time_conn(ConnectionString);

            DateTime dateTime1 = DateTime.ParseExact(c_datetime, "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture).AddDays(-1);
            txtdate1.Text = dateTime1.Date.ToShortDateString();
        }
        private void load_pic(PictureBox pic_object, string p_file_path)
        {
            if (!string.IsNullOrEmpty(p_file_path.Trim()))
            {
                pic_object.ImageLocation = photopath.Trim() + p_file_path.Trim();
            }
            else
            {
                pic_object.ImageLocation = "";
            }
            pic_object.Refresh();
        }
        private void action_access()
        {
            Console.WriteLine("action_access");

            string tsql = "select TOP 150 l.*, e.Caption as event_name from LOGt l left join Log_Events e on l.Trn_Type =e.id  where Trn_Type = 1 or Trn_Type = 2";
            //tsql = tsql = " where convert(varchar, Date, 120) >= '" + get_data_date.ToString("yyyy-MM-dd HH:mm:ss") + "'"
            tsql = tsql + " order by date desc";
            using (OdbcConnection connection = new OdbcConnection(ConnectionString))
            {




                OdbcCommand cmd = new OdbcCommand(tsql, connection);
                OdbcDataAdapter da = new OdbcDataAdapter(cmd);

                DataTable dt = new DataTable();
                da.Fill(dt);
                Boolean is_frist = true;
                Boolean is_diff = false;
                foreach (DataRow dr in dt.Rows)
                {

                    int get_id = int.Parse(dr["ID"].ToString());
                    if (is_frist == true)
                    {
                        is_frist = false;
                        string get_data_date_word = dr["Date"].ToString().Trim();
                        Console.WriteLine("get_data_date_word: " + get_data_date_word);
                        Console.WriteLine("get_data_date_word: " + get_data_date_word);
                        get_data_date = DateTime.ParseExact(get_data_date_word, "yyyy-MM-dd HH:mm:ss", null);
                        Console.WriteLine("ID:" + dr["ID"].ToString());
                        Console.WriteLine("get_log_id:" + get_log_id.ToString());
                        if (get_log_id != int.Parse(dr["ID"].ToString()))
                        {
                            get_log_id = int.Parse(dr["ID"].ToString());
                            is_diff = true;
                        }
                        else
                        {
                            return;
                        }
                    }
                }

                if (is_diff == true)
                {
                    is_diff = false;

                    Console.WriteLine("Reload Data 1");
                    log_data.Rows.Clear();
                    int tno = 0;
                    foreach (DataRow dr in dt.Rows)
                    {
                        tno = tno + 1;

                        int log_id = int.Parse(dr["ID"].ToString());
                        string get_data_date_word = dr["Date"].ToString().Trim();
                        Console.WriteLine("get_data_date_word: " + get_data_date_word);

                        DateTime log_date = DateTime.ParseExact(get_data_date_word, "yyyy-MM-dd HH:mm:ss", null);

                        string Trn_Type = dr["Trn_Type"].ToString().Trim();
                        string event_name = dr["event_name"].ToString().Trim();
                        string status = event_name;
                        string from_name = dr["From_name"].ToString().Trim();
                        string desc3 = (dr["Desc3"].ToString()).Trim();

                        string description = (dr["Desc3"].ToString()).Trim();

                        if (!string.IsNullOrEmpty(status))
                        {
                            description = " [ " + status + " ] " + description;
                        }


                        Console.WriteLine("desc3: " + desc3);


                        int tat = desc3.LastIndexOf(" ");
                        string number = "";
                        string tmp_name = "";
                        string query = "";
                        if (tat < 0)
                        {
                            tmp_name = desc3.Trim();
                            query = "select TOP 1 h.*, c.Code FROM CRDHLD h left join card c on h.id= c.Owner where rtrim(ltrim((h.last_name  + h.First_Name))) = '" + tmp_name + "'";
                            Console.WriteLine("query: " + query);
                        }
                        else
                        {
                            number = desc3.Substring(tat).Trim();
                            query = "select TOP 1 h.*, c.Code FROM CRDHLD h left join card c on h.id= c.Owner where h.num = '" + number + "'";

                        }
                        Console.WriteLine("number:" + number);

                        OdbcCommand cmd2 = new OdbcCommand(query, connection);
                        OdbcDataAdapter da2 = new OdbcDataAdapter(cmd2);
                        DataTable dt2 = new DataTable();
                        da2.Fill(dt2);
                        string photo = "";
                        foreach (DataRow dr2 in dt2.Rows)
                        {

                            photo = dr2["photo"].ToString().Trim();
                            description = description + " {Card: " + dr2["Code"].ToString().Trim() + " }";
                        }
                        description = description + " " + from_name;
                        DataRow _Row = log_data.NewRow();
                        _Row["Date"] = log_date;
                        _Row["Description"] = description;

                        if (tno == 1)
                        {
                            load_pic(pictureBox1, photo);
                        }
                        if (tno == 2)
                        {
                            load_pic(pictureBox2, photo);
                        }
                        if (tno == 3)
                        {
                            load_pic(pictureBox3, photo);
                        }
                        if (tno == 4)
                        {
                            load_pic(pictureBox4, photo);
                        }
                        if (tno == 5)
                        {
                            load_pic(pictureBox5, photo);
                        }
                        if (tno == 6)
                        {
                            load_pic(pictureBox6, photo);
                        }
                        if (tno == 7)
                        {
                            load_pic(pictureBox7, photo);
                        }
                        if (tno == 8)
                        {
                            load_pic(pictureBox8, photo);
                        }
                        if (tno == 9)
                        {
                            load_pic(pictureBox9, photo);
                        }
                        if (tno == 10)
                        {
                            load_pic(pictureBox10, photo);
                        }

                        log_data.Rows.Add(_Row);
                    }
                }

                dt.Dispose();
                da.Dispose();
                cmd.Dispose();
                connection.Dispose();

            }



        }
        private void action_no_access()
        {
            Console.WriteLine("action_no_access");
            string tsql = "select TOP 150 l.*, e.Caption as event_name from LOGt l left join Log_Events e on l.Trn_Type =e.id where Trn_Type = 3 or Trn_Type = 4 or Trn_Type = 63";
            //tsql = tsql = " where convert(varchar, Date, 120) >= '" + get_data_date.ToString("yyyy-MM-dd HH:mm:ss") + "'"
            tsql = tsql + " order by date desc";
            using (OdbcConnection connection = new OdbcConnection(ConnectionString))
            {
                OdbcCommand cmd = new OdbcCommand(tsql, connection);
                OdbcDataAdapter da = new OdbcDataAdapter(cmd);

                DataTable dt = new DataTable();
                da.Fill(dt);
                Boolean is_frist = true;
                Boolean is_diff = false;
                foreach (DataRow dr in dt.Rows)
                {
                    int get_id = int.Parse(dr["ID"].ToString());
                    if (is_frist == true)
                    {
                        is_frist = false;
                        string get_data_date_word = dr["Date"].ToString().Trim();
                        Console.WriteLine("get_data_date_word: " + get_data_date_word);

                        get_data_date = DateTime.ParseExact(get_data_date_word, "yyyy-MM-dd HH:mm:ss", null);
                        Console.WriteLine("ID:" + dr["ID"].ToString());
                        Console.WriteLine("get_log_id:" + get_log_id.ToString());
                        if (get_log_id_no_access != int.Parse(dr["ID"].ToString()))
                        {
                            get_log_id_no_access = int.Parse(dr["ID"].ToString());
                            is_diff = true;
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                if (is_diff == true)
                {
                    is_diff = false;

                    Console.WriteLine("Reload Data 2");
                    log_data_no_access.Rows.Clear();

                    int tno = 0;

                    foreach (DataRow dr in dt.Rows)
                    {
                        tno = tno + 1;

                        int log_id = int.Parse(dr["ID"].ToString());
                        string get_data_date_word = dr["Date"].ToString().Trim();
                        Console.WriteLine("get_data_date_word: " + get_data_date_word);

                        DateTime log_date = DateTime.ParseExact(get_data_date_word, "yyyy-MM-dd HH:mm:ss", null);

                        string Trn_Type = dr["Trn_Type"].ToString().Trim();
                        string event_name = dr["event_name"].ToString().Trim();

                        string status = event_name;
                        string from_name = dr["From_name"].ToString().Trim();
                        string description = (dr["Desc3"].ToString()).Trim();
                        string desc3 = (dr["Desc3"].ToString()).Trim();



                        if (!string.IsNullOrEmpty(status))
                        {
                            description = " [ " + status + " ] " + description;
                        }



                        DataRow _Row = log_data_no_access.NewRow();
                        _Row["Date"] = log_date;

                        Console.WriteLine(description);



                        //

                        int tat = desc3.LastIndexOf(" ");
                        string number = "";
                        string tmp_name = "";
                        string query = "";
                        if (tat < 0)
                        {
                            tmp_name = desc3.Trim();
                            query = "select TOP 1 h.*, c.Code FROM CRDHLD h left join card c on h.id= c.Owner where rtrim(ltrim((h.last_name  + h.First_Name))) = '" + tmp_name + "'";
                            Console.WriteLine("query: " + query);
                        }
                        else
                        {
                            number = desc3.Substring(tat).Trim();
                            query = "select TOP 1 h.*, c.Code FROM CRDHLD h left join card c on h.id= c.Owner where h.num = '" + number + "'";

                        }


                        OdbcCommand cmd2 = new OdbcCommand(query, connection);
                        OdbcDataAdapter da2 = new OdbcDataAdapter(cmd2);
                        DataTable dt2 = new DataTable();
                        da2.Fill(dt2);
                        string photo = "";
                        foreach (DataRow dr2 in dt2.Rows)
                        {

                            photo = dr2["photo"].ToString().Trim();
                            description = description + " {Card: " + dr2["Code"].ToString().Trim() + " }";
                        }

                        if (!string.IsNullOrEmpty(description))
                        {
                            description = description + " at ";
                        }
                        description = description + from_name;
                        _Row["Description"] = description;

                        if (tno == 1)
                        {
                            load_pic(pictureBox11, photo);
                        }
                        if (tno == 2)
                        {
                            load_pic(pictureBox12, photo);
                        }
                        if (tno == 3)
                        {
                            load_pic(pictureBox13, photo);
                        }
                        if (tno == 4)
                        {
                            load_pic(pictureBox14, photo);
                        }
                        if (tno == 5)
                        {
                            load_pic(pictureBox15, photo);
                        }
                        if (tno == 6)
                        {
                            load_pic(pictureBox16, photo);
                        }
                        if (tno == 7)
                        {
                            load_pic(pictureBox17, photo);
                        }
                        if (tno == 8)
                        {
                            load_pic(pictureBox18, photo);
                        }
                        if (tno == 9)
                        {
                            load_pic(pictureBox19, photo);
                        }
                        if (tno == 10)
                        {
                            load_pic(pictureBox20, photo);
                        }
                        //

                        log_data_no_access.Rows.Add(_Row);
                    }
                }
                dt.Dispose();
                da.Dispose();
                cmd.Dispose();
                connection.Dispose();
            }


        }
        private void action_system()
        {
            Console.WriteLine("action_system");


            string tsql = "select TOP 150 l.*, e.Caption as event_name from LOGt l left join Log_Events e on l.Trn_Type =e.id where Trn_Type > 4 and Trn_Type <> 63";
            //tsql = tsql = " where convert(varchar, Date, 120) >= '" + get_data_date.ToString("yyyy-MM-dd HH:mm:ss") + "'"
            tsql = tsql + " order by date desc";
            using (OdbcConnection connection = new OdbcConnection(ConnectionString))
            {
                OdbcCommand cmd = new OdbcCommand(tsql, connection);
                OdbcDataAdapter da = new OdbcDataAdapter(cmd);

                DataTable dt = new DataTable();
                da.Fill(dt);
                Boolean is_frist = true;
                Boolean is_diff = false;
                foreach (DataRow dr in dt.Rows)
                {

                    int get_id = int.Parse(dr["ID"].ToString());
                    if (is_frist == true)
                    {
                        is_frist = false;
                        string get_data_date_word = dr["Date"].ToString().Trim();
                        Console.WriteLine("get_data_date_word: " + get_data_date_word);

                        get_data_date = DateTime.ParseExact(get_data_date_word, "yyyy-MM-dd HH:mm:ss", null);
                        Console.WriteLine("ID:" + dr["ID"].ToString());
                        Console.WriteLine("get_log_id:" + get_log_id.ToString());
                        if (get_log_id_system != int.Parse(dr["ID"].ToString()))
                        {
                            get_log_id_system = int.Parse(dr["ID"].ToString());
                            is_diff = true;
                        }
                        else
                        {
                            return;
                        }
                    }
                }
                if (is_diff == true)
                {
                    is_diff = false;

                    Console.WriteLine("Reload Data 2");
                    log_data_system.Rows.Clear();
                    foreach (DataRow dr in dt.Rows)
                    {
                        int log_id = int.Parse(dr["ID"].ToString());
                        string get_data_date_word = dr["Date"].ToString().Trim();
                        Console.WriteLine("get_data_date_word: " + get_data_date_word);

                        DateTime log_date = DateTime.ParseExact(get_data_date_word, "yyyy-MM-dd HH:mm:ss", null);

                        string Trn_Type = dr["Trn_Type"].ToString().Trim();
                        string event_name = dr["event_name"].ToString().Trim();

                        string status = event_name;
                        string from_name = dr["From_name"].ToString().Trim();
                        string description = (dr["Desc3"].ToString()).Trim();
                        string desc1 = (dr["Desc1"].ToString()).Trim();


                        if (!string.IsNullOrEmpty(description))
                        {
                            description = description + " at ";
                        }
                        if (Trn_Type == "10" || Trn_Type == "11")
                        {

                            if (desc1 == "0")
                            {
                                description = "Forced open " + description;

                            }
                            if (desc1 == "1")
                            {
                                description = "Held open " + description;

                            }
                            if (desc1 == "2")
                            {
                                description = "Return normal " + description;

                            }
                        }

                        if (!string.IsNullOrEmpty(status))
                        {
                            description = " [ " + status + " ] " + description;
                        }

                        description = description + from_name;


                        DataRow _Row = log_data_system.NewRow();
                        _Row["Date"] = log_date;
                        _Row["Description"] = description;

                        Console.WriteLine(description);
                        log_data_system.Rows.Add(_Row);
                    }
                }
                dt.Dispose();
                da.Dispose();
                cmd.Dispose();
                connection.Dispose();
            }

        }
        private void timer1_Tick(object sender, EventArgs e)
        {

            timer1.Enabled = false;
            action_access();
            action_no_access();
            action_system();

            timer1.Enabled = true;

        }



        private void pictureBox1_Click(object sender, EventArgs e)
        {









        }

        private void pictureBox7_Click(object sender, EventArgs e)
        {

        }

        private void tabControl1_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Console.WriteLine("tabControl1.SelectedTab.Tab: " + tabControl1.SelectedTab.Tag.ToString());
            switch (tabControl1.SelectedTab.Tag.ToString().Trim())
            {
                case "1":
                    timer1.Enabled = true;
                    timer_get_card_no.Enabled = false;

                    //                timer_get_c_card_no.Enabled = false;
                    //                timer_get_r_card_no.Enabled = false;
                    Console.WriteLine("tab 1");

                    break;
                case "2":
                    timer1.Enabled = false;

                    //                 timer_get_card_no.Enabled = true;
                    //                 timer_get_c_card_no.Enabled = false;
                    //                 timer_get_r_card_no.Enabled = false;
                    Console.WriteLine("tab 2");
                    txtcard_no.Select();


                    break;

                case "3":
                    timer1.Enabled = false;
                    //                timer_get_card_no.Enabled = false;
                    //                timer_get_c_card_no.Enabled = true;
                    //               timer_get_r_card_no.Enabled = false;

                    Console.WriteLine("tab 3");
                    txtfind.Select();

                    break;
                case "4":
                    timer1.Enabled = false;
                    //                 timer_get_card_no.Enabled = false;
                    //             timer_get_c_card_no.Enabled = false;
                    //             timer_get_r_card_no.Enabled = true;
                    txtr_card_no.Select();

                    Console.WriteLine("tab 4");
                    break;

                case "5":

                    txtl_card_no.Select();

                    Console.WriteLine("tab 5");
                    break;

                case "6":

                    txtfind_data.Select();

                    Console.WriteLine("tab 6");
                    break;
                case "7":
                    find_not_return_card();
                    txtfind_data2.Select();

                    break;
                default:
                    timer1.Enabled = true;
                    timer_get_card_no.Enabled = false;
                    timer_get_c_card_no.Enabled = false;


                    break;


            }
            //     if (tabControl1.SelectedTab == tabControl1.TabPages["tabPage1"])            { }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            cam cam = new cam();

            var result = cam.ShowDialog();
            if (result == DialogResult.OK)
            {
                string val = cam.ruturn_value;           //values preserved after close

                get_photopath = val;
                pic_photo.ImageLocation = get_photopath;


            }



            //      cam.Show();
            //     cam.Activate();
            //    cam.FormBorderStyle = FormBorderStyle.Sizable;

        }

        private void tabPage2_Click(object sender, EventArgs e)
        {

        }
        private void clean_all()
        {



            get_photopath = "";

            txtLastName.Text = "";
            txtFirstName.Text = "";
            txtstaff_no.Text = "";
            txtstaff_date.Text = "";
            txtidcard.Text = "";
            txtidcard_code.Text = "";
            txtpassport_no.Text = "";
            txtpassport_date.Text = "";
            txtother.Text = "";
            comdept.Text = "";
            pic_photo.ImageLocation = "";
            txtcard_no.Text = "";
            txtphone.Text = "";
            comdept.BackColor = SystemColors.Window;


            txtc_card_no.Text = "";
            txtc_name.Text = "";
            txtc_staff_no.Text = "";
            txtc_staff_date.Text = "";
            txtc_idcard.Text = "";

            txtc_passport_no.Text = "";
            txtc_passport_date.Text = "";
            txtc_other.Text = "";
            txtc_dept.Text = "";
            txtc_num.Text = "";
            txtc_id.Text = "";

            txtc_photo_file.Text = "";
            txtc_phone.Text = "";

            txtfind.Text = "";
            pic_c_photo.ImageLocation = "";
            txtc_dept.BackColor = SystemColors.Control;

            txtr_num.Text = "";
            txtr_id.Text = "";

            txtr_card_no.Text = "";
            txtr_name.Text = "";
            txtr_staff_no.Text = "";
            txtr_staff_date.Text = "";
            txtr_idcard.Text = "";
            txtr_passport_no.Text = "";
            txtr_passport_date.Text = "";
            txtr_other.Text = "";
            txtr_dept.Text = "";
            txtr_phone.Text = "";
            txtr_dept.BackColor = SystemColors.Control;

            txtl_num.Text = "";
            txtl_id.Text = "";
            txtl_card_no.Text = "";
            txtl_name.Text = "";
            txtl_staff_no.Text = "";
            txtl_idcard.Text = "";
            txtl_passport_no.Text = "";
            txtl_passport_date.Text = "";
            txtl_other.Text = "";
            txtl_dept.Text = "";
            txtl_phone.Text = "";
            txtl_dept.BackColor = SystemColors.Control;


            pic_r_photo.ImageLocation = "";

            txtstatus.BackColor = SystemColors.Control;
            txtstatus.ForeColor = SystemColors.Control;


        }

        private void button2_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "jpg files (*.jpg)|*.jpg";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                get_photopath = openFileDialog1.FileName;
                pic_photo.ImageLocation = get_photopath;


            }
        }

        private void comdept_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comdept.Text.Trim().ToUpper() == "FTC")
            {
                comdept.BackColor = Color.FromArgb(255, 128, 0);
            }
            else
            {
                if (comdept.Text.Trim().ToUpper() == "DDW")
                {
                    comdept.BackColor = Color.FromArgb(255, 204, 1);
                }
                else
                {
                    if (comdept.Text.Trim().ToUpper() == "ADC")
                    {
                        comdept.BackColor = Color.FromArgb(255, 128, 255);
                    }
                    else
                    {
                        if (comdept.Text.Trim().ToUpper() == "PH3 ENTRANCE" || comdept.Text.Trim().ToUpper() == "PH3" || comdept.Text.Trim().ToUpper() == "PHASE3 ENTRANCE")
                        {
                             comdept.BackColor = Color.FromArgb(255, 192, 255);
                          

                        }
                        else
                        {

                            if (comdept.Text.Trim().ToUpper() == "SFC")
                            {
                                comdept.BackColor = Color.FromArgb(135, 206, 250);
                            }
                            else
                            {
                                //   comdept.BackColor = SystemColors.Window;
                            }
                        }
                    }
                }
            }
        }



        private string read_card()
        {
            string c_datetime = get_server_time_conn(ConnectionString);
            string card_no = "";

            DateTime oDate = DateTime.ParseExact(c_datetime, "yyyy-MM-dd HH:mm:ss", null);
            oDate = oDate.AddSeconds(-15);

            //  action_time = DateTime.Now.AddSeconds(-10).ToString("yyyy-MM-dd HH:mm:ss");
            string check_datetime = oDate.ToString("yyyy-MM-dd HH:mm:ss");
            using (OdbcConnection connection = new OdbcConnection(ConnectionString))
            {
                string query = "select TOP 1 * FROM LOGt where (trn_type = 1 or trn_type = 3 or trn_type = 61 or trn_type = 63)  and [From_Name] = '" + card_reader_name.Trim() + "' and date >= '" + check_datetime + "' and date > '" + card_record_dt + "' order by  id desc";

                Console.WriteLine("query: " + query);




                OdbcCommand cmd = new OdbcCommand(query, connection);
                OdbcDataAdapter da = new OdbcDataAdapter(cmd);

                DataTable dt = new DataTable();

                da.Fill(dt);
                string desc3 = "";
                string number = "";
                string name = "";

                int trn_type = 0;
                if (dt.Rows.Count <= 0)
                {
                    return "";
                }
                foreach (DataRow dr in dt.Rows)
                {


                    int card_record_id = int.Parse(dr["ID"].ToString());

                    card_record_dt = dr["Date"].ToString().Trim();

                    trn_type = int.Parse(dr["Trn_type"].ToString());
                    Console.WriteLine("trn_type: " + trn_type.ToString());
                    if (trn_type == 61 || trn_type == 63)
                    {

                        //card_no = (int.Parse(dr["desc3"].ToString().Trim()) - 6291456).ToString().Trim().PadLeft(6, '0');
                        card_no = (int.Parse(dr["desc3"].ToString().Trim())).ToString().Trim().PadLeft(8, '0');
                        Console.WriteLine("card_no:" + card_no.ToString());
                    }
                    else
                    {
                        desc3 = dr["desc3"].ToString().Trim();
                        int find_at = desc3.LastIndexOf(' ');

                        number = desc3.Substring(find_at + 1, (desc3.Length - find_at - 1));
                        //  System.Diagnostics.Debug.WriteLine(number);
                        Console.WriteLine("number: " + number);
                        query = "select TOP 1 h.last_name,c.code,free_field1 as gender FROM CRDHLD h, card c where h.id = c.owner and h.num = '" + number + "'";
                        //       MessageBox.Show(query);
                        OdbcCommand cmd3 = new OdbcCommand(query, connection);
                        OdbcDataAdapter da3 = new OdbcDataAdapter(cmd3);
                        DataTable dt3 = new DataTable();
                        da3.Fill(dt3);

                        foreach (DataRow dr3 in dt3.Rows)
                        {
                            //card_no = (int.Parse(dr3["Code"].ToString()) - 6291456).ToString().Trim().PadLeft(6, '0');
                            card_no = (int.Parse(dr3["Code"].ToString())).ToString().Trim().PadLeft(8, '0');
                            Console.WriteLine("card_no:" + card_no.ToString());


                        }
                    }





                }

                dt.Dispose();
                da.Dispose();
                cmd.Dispose();
                connection.Dispose();
            }
            return card_no;


        }
        private void button5_Click(object sender, EventArgs e)
        {
            string c_datetime = get_server_time_conn(ConnectionString);

            DateTime oDate = DateTime.ParseExact(c_datetime, "yyyy-MM-dd HH:mm:ss", null);
            // oDate = oDate.AddSeconds(-5);

            //  action_time = DateTime.Now.AddSeconds(-10).ToString("yyyy-MM-dd HH:mm:ss");

            string check_datetime = oDate.ToString("yyyy-MM-dd HH:mm:ss");
            using (OdbcConnection connection = new OdbcConnection(ConnectionString))
            {
                string query = "select TOP 1 * FROM LOGt where (trn_type = 1 or trn_type = 61 or trn_type = 63)  and [From_Name] = '" + card_reader_name.Trim() + "' and date >= '" + check_datetime + "' and date > '" + card_record_dt + "' order by  id desc";
                Console.WriteLine("query: " + query);
                //string query = "select TOP 1 * FROM LOGt where (trn_type = 1 or trn_type = 61 or trn_type = 63) and date >= '" + check_datetime + "' and date > '" + card_record_dt + "' order by  id desc";
                //   string query = "select TOP 1 * FROM LOGt where (trn_type = 1 or trn_type = 61 or trn_type = 63) order by  id desc";

                OdbcCommand cmd = new OdbcCommand(query, connection);
                OdbcDataAdapter da = new OdbcDataAdapter(cmd);

                DataTable dt = new DataTable();

                da.Fill(dt);
                string desc3 = "";
                string number = "";
                string name = "";
                string card_no = "";

                int trn_type = 0;
                if (dt.Rows.Count <= 0)
                {
                    return;
                }
                foreach (DataRow dr in dt.Rows)
                {


                    int card_record_id = int.Parse(dr["ID"].ToString());

                    card_record_dt = dr["Date"].ToString().Trim();

                    trn_type = int.Parse(dr["Trn_type"].ToString());
                    Console.WriteLine("trn_type: " + trn_type.ToString());
                    if (trn_type == 61 || trn_type == 63)
                    {

                        //card_no = (int.Parse(dr["desc3"].ToString().Trim()) - 6291456).ToString().Trim().PadLeft(6, '0');
                        card_no = (int.Parse(dr["desc3"].ToString().Trim())).ToString().Trim().PadLeft(8, '0');
                        txtc_card_no.Text = card_no;
                        Console.WriteLine("card_no:" + card_no.ToString());
                    }
                    else
                    {
                        desc3 = dr["desc3"].ToString().Trim();
                        int find_at = desc3.LastIndexOf(' ');

                        number = desc3.Substring(find_at + 1, (desc3.Length - find_at - 1));
                        //  System.Diagnostics.Debug.WriteLine(number);
                        query = "select TOP 1 h.last_name,c.code,free_field1 as gender FROM CRDHLD h, card c where h.id = c.owner and h.num = '" + number + "'";
                        //  MessageBox.Show(query);
                        OdbcCommand cmd3 = new OdbcCommand(query, connection);
                        OdbcDataAdapter da3 = new OdbcDataAdapter(cmd3);
                        DataTable dt3 = new DataTable();
                        da3.Fill(dt3);

                        foreach (DataRow dr3 in dt3.Rows)
                        {
                            //card_no = (int.Parse(dr3["Code"].ToString()) - 6291456).ToString().Trim().PadLeft(6, '0');
                            card_no = (int.Parse(dr3["Code"].ToString())).ToString().Trim().PadLeft(8, '0');
                            txtc_card_no.Text = card_no;
                            Console.WriteLine("card_no:" + card_no.ToString());


                        }
                        dt3.Dispose();
                        da3.Dispose();
                        cmd3.Dispose();
                        connection.Dispose();
                    }






                }
                dt.Dispose();
                da.Dispose();
                cmd.Dispose();
                connection.Dispose();
            }

        }

        private void button6_Click(object sender, EventArgs e)
        {

        }
        private Boolean check_pass()
        {
            Boolean treturn = true;

            if (string.IsNullOrEmpty(txtLastName.Text.Trim()))
            {
                treturn = false;
                MessageBox.Show("必需輸入 {姓氏}");
            }
            if (string.IsNullOrEmpty(txtFirstName.Text.Trim()))
            {
                treturn = false;
                MessageBox.Show("必需輸入 {名稱}");
            }
            if (string.IsNullOrEmpty(comdept.Text.Trim()))
            {
                treturn = false;
                MessageBox.Show("必需輸入 {部門}");
            }
            if (string.IsNullOrEmpty(txtidcard.Text.Trim()))
            {
                treturn = false;
                MessageBox.Show("必需輸入 {身份證號碼}");
            }
            if (string.IsNullOrEmpty(txtcard_no.Text.Trim()))
            {
                treturn = false;
                MessageBox.Show("必需 {抇卡} 或 輸入 {卡號}");
            }
            if (string.IsNullOrEmpty(txtstaff_no.Text.Trim()))
            {
                treturn = false;
                MessageBox.Show("必需輸入 {員工編號}");
            }

            return treturn;
        }

        private void button3_Click(object sender, EventArgs e)
        {


            string card_no = txtcard_no.Text.Trim().ToUpper();
            if (!string.IsNullOrEmpty(card_no))
            {
                card_no = card_no.PadLeft(8, '0');
            }
            string lastname = txtLastName.Text.Trim();
            string firstname = txtFirstName.Text.Trim();
            string staff_no = txtstaff_no.Text.Trim();
            string staff_date = txtstaff_date.Text.Trim();
            string c_from_date = txtstaff_date.Text.Trim();
            string phone = txtphone.Text.Trim();

            string passport_no = txtpassport_no.Text.Trim().ToUpper();
            string passport_date = txtpassport_date.Text.Trim().ToUpper();
            string other = txtother.Text.Trim();
            string dept = comdept.Text.Trim();
            string card_reader = card_reader_name.Trim();

            string idcard = txtidcard.Text.Trim().ToUpper();
            string idcard_code = txtidcard_code.Text.Trim().ToUpper();
            string department = "";
            if (string.IsNullOrEmpty(idcard_code.Trim()) == false)
            {

                idcard = idcard.Trim() + "(" + idcard_code.Trim() + ")";

            }

            if (check_pass() == false)
            {
                return;
            }
            // check 

            string c_datetime = get_server_time_conn(ConnectionString);

            using (OdbcConnection connection = new OdbcConnection(ConnectionString))
            {
                string query = "SELECT c.Code , h.*  FROM CRDHLD h , Card c where  h.id = c.Owner and c.code = '" + card_no + "'";


                OdbcCommand cmd = new OdbcCommand(query, connection);

                OdbcDataAdapter da = new OdbcDataAdapter(cmd);

                DataTable dt = new DataTable();
                da.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        MessageBox.Show("卡號: { " + dr["code"].ToString().Trim() + " } 已屬於 " + (dr["first_name"].ToString().Trim() + " " + dr["last_name"].ToString().Trim()).Trim(), "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }


                    return;
                }

                dt.Dispose();
                da.Dispose();
                cmd.Dispose();
                connection.Dispose();

            }

            using (OdbcConnection conn1 = new OdbcConnection(ConnectionString))
            {
                string query = "SELECT c.Code , h.*  FROM CRDHLD h , Card c where  h.id = c.Owner and h.idcard = '" + idcard + "'";

                OdbcCommand cmd = new OdbcCommand(query, conn1);

                OdbcDataAdapter da = new OdbcDataAdapter(cmd);

                DataTable dt = new DataTable();
                da.Fill(dt);
                if (dt.Rows.Count > 0)
                {

                    foreach (DataRow dr in dt.Rows)
                    {
                        MessageBox.Show("這身份証: { " + dr["idcard"].ToString().Trim() + " } 已擁有卡號: {" + dr["code"].ToString().Trim() + "} ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
                    return;
                }

                dt.Dispose();
                da.Dispose();
                cmd.Dispose();
                conn1.Dispose();
            }

            using (OdbcConnection conn1 = new OdbcConnection(ConnectionString))
            {
                string query = "SELECT h.*  FROM CRDHLD h where h.idcard = '" + idcard + "'";

                OdbcCommand cmd = new OdbcCommand(query, conn1);

                OdbcDataAdapter da = new OdbcDataAdapter(cmd);

                DataTable dt = new DataTable();
                da.Fill(dt);
                if (dt.Rows.Count > 0)
                {

                    foreach (DataRow dr in dt.Rows)
                    {
                        MessageBox.Show("這身分證:" + " { " + idcard.Trim() + "} 系統已登記於 " + (dr["first_name"].ToString().Trim() + " " + dr["last_name"].ToString().Trim()).Trim(), "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
                    return;
                }
                dt.Dispose();
                da.Dispose();
                cmd.Dispose();
                conn1.Dispose();


            }



            // check end

            using (OdbcConnection conn2 = new OdbcConnection(ConnectionString2))
            {



                string query = "";




                query = "INSERT INTO dr_card_log ([event_dt],[borrow_return],[card_reader],[idcard],[name],[staff_no],[staff_date],[phone],[passport_no],[passport_date],[other],[dept],[card_no])";

                query = query + " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)";

                //
                //

                OdbcCommand command = new OdbcCommand(query, conn2);





                command.Parameters.AddWithValue("event_dt", c_datetime);
                command.Parameters.AddWithValue("borrow_return", "Add_Borrow");

                command.Parameters.AddWithValue("card_reader", card_reader);
                command.Parameters.AddWithValue("idcard", idcard);


                command.Parameters.AddWithValue("name", lastname);

                command.Parameters.AddWithValue("staff_no", staff_no);

                command.Parameters.AddWithValue("staff_date", staff_date.ToString().Trim());

                command.Parameters.AddWithValue("phone", phone);


                command.Parameters.AddWithValue("passport_no", passport_no.ToString().Trim());

                command.Parameters.AddWithValue("passport_date", passport_date);
                command.Parameters.AddWithValue("other", other);
                command.Parameters.AddWithValue("dept", dept);
                command.Parameters.AddWithValue("card_no", card_no);
                conn2.Open();
                int result = command.ExecuteNonQuery();

                if (result < 0)
                {
                    Console.WriteLine("加不到借卡記錄, 身份證為:" + idcard.ToString() + " !");
                    MessageBox.Show("錯誤! 加不到借卡記錄, 身份證為: " + idcard.ToString() + " !");
                }
                else
                {
                    Console.WriteLine("加不到借卡記錄, 身份證為:" + idcard.ToString() + " !");

                };
                //
                command.Dispose();
                conn2.Dispose();


            }




            int tnext_no = 1;
            Boolean add_card = false;

            string c_rev_date = get_server_time_conn(connection1);

            using (OdbcConnection conn2 = new OdbcConnection(connection2))
            {
                conn2.Open();
                string query = "select next_no FROM sys_next_no where sys_name = 'CARDHLD' ";
                OdbcCommand cmd = new OdbcCommand(query, conn2);




                var reader = cmd.ExecuteReader();
                int f_next_no = reader.GetOrdinal("next_no");
                if (!reader.Read())
                    throw new InvalidOperationException("No records were returned.");

                tnext_no = reader.GetInt32(f_next_no);
                Console.WriteLine("get next_no: " + tnext_no.ToString());

                cmd.Dispose();
                conn2.Dispose();
            }





            using (OdbcConnection conn1 = new OdbcConnection(connection1))
            {
                conn1.Open();
                OdbcCommand command1;
                OdbcDataAdapter da3;
                DataTable dt3 = new DataTable();
                while (true)
                {



                    string c_tnext_no = tnext_no.ToString();

                    string query3 = "select h.num from CRDHLD h where  h.num = '" + c_tnext_no + "'";

                    command1 = new OdbcCommand(query3, conn1);
                    da3 = new OdbcDataAdapter(command1);
                    dt3.Clear();


                    da3.Fill(dt3);
                    if (dt3.Rows.Count > 0)
                    {

                        tnext_no = tnext_no + 1;

                    }
                    else
                    {
                        break;
                    }


                }
                Console.WriteLine("check and get next_no: " + tnext_no.ToString());
                dt3.Dispose();
                da3.Dispose();
                command1.Dispose();


                conn1.Dispose();
            }

            using (OdbcConnection conn2 = new OdbcConnection(connection2))
            {
                conn2.Open();
                string query4 = "update sys_next_no set next_no = " + (tnext_no + 1) + " where sys_name = 'CARDHLD' ";
                OdbcCommand cmd4 = new OdbcCommand(query4, conn2);




                var reader4 = cmd4.ExecuteNonQuery();

                Console.WriteLine("update next_no: " + (tnext_no + 1).ToString());

                cmd4.Dispose();
                conn2.Dispose();
            }

            string num = tnext_no.ToString().Trim().ToUpper();
            string c_to_date = "";

            using (OdbcConnection conn1 = new OdbcConnection(connection1))
            {

                conn1.Open();



                string group_right = comdept.Text.Trim();
                if (group_right.Equals("DDW"))
                {
                    group_right = "Hub-Badge(Vis)";
                }
                if (group_right.Equals("ADC"))
                {
                    group_right = "ADECCO";
                }
                if (group_right.Equals("SFC"))
                {
                    group_right = "Con Cleaning";
                }
                string company = comdept.Text.Trim();
                if (string.IsNullOrEmpty(c_from_date) == true)
                {
                    c_from_date = c_rev_date;
                }

                DateTime oDate = DateTime.ParseExact(c_from_date, "yyyy-MM-dd HH:mm:ss", null);
                oDate = oDate.AddYears(2);

                c_to_date = oDate.ToString("yyyy-MM-dd HH:mm:ss");

                string dep_name = comdept.Text.Trim();
                if (dep_name.Equals("ADC"))
                {
                    department = "Adecco";
                }


                string tphotopath = "";
                string image_filename = "";
                if (!string.IsNullOrEmpty(get_photopath))
                {
                    File.Copy(get_photopath, photopath + @"" + num + ".jpg", true);
                    tphotopath = photopath + @"" + num + ".jpg";
                    image_filename = System.IO.Path.GetFileName(tphotopath);
                }

                Console.WriteLine("photopath: " + photopath);
                Console.WriteLine("tphotopath: " + tphotopath);



                string query5 = "INSERT INTO QueueMSGAPI ([DateCreated], [ServerName], [Cmd], [Msg], [Status], [Result]) VALUES(?,?,?,?,?,?)";



                OdbcCommand command5 = new OdbcCommand(query5, conn1);




                String tmsg = "<query>";
                tmsg = tmsg + "<Number>" + num.ToString() + "</Number>";
                tmsg = tmsg + "<Last_Name>" + lastname + "</Last_Name>";
                tmsg = tmsg + "<First_Name>" + firstname + "</First_Name>";
                if (string.IsNullOrEmpty(card_no.Trim()) == false)
                {
                    tmsg = tmsg + "<Badge>" + card_no + "</Badge>";
                    tmsg = tmsg + "<BorrowCardDate>" + c_datetime + "</BorrowCardDate>";
                }
                tmsg = tmsg + "<From_Date>" + c_from_date + "</From_Date>";
                tmsg = tmsg + "<To_Date>" + c_to_date + "</To_Date>";

                tmsg = tmsg + "<Company>" + dep_name + "</Company>";


                tmsg = tmsg + "<Access_Group>" + group_right + "</Access_Group>";

                string type = "1";
                tmsg = tmsg + "<type>" + type + "</type>";
                tmsg = tmsg + "<Technology>3</Technology>";
                tmsg = tmsg + "<Description>1</Description>";
                if (!String.IsNullOrEmpty(image_filename))
                {
                    tmsg = tmsg + "<Photo>" + image_filename + "</Photo>";


                }
                phone = String.IsNullOrEmpty(staff_no) ? "00000" : staff_no;
                tmsg = tmsg + "<Office_Phone>" + phone + "</Office_Phone>";


                tmsg = tmsg + "<idcard>" + idcard + "</idcard>";
                tmsg = tmsg + "<StaffDate>" + c_from_date + "</StaffDate>";
                tmsg = tmsg + "<Passport>" + passport_no + "</Passport>";
                tmsg = tmsg + "<PassportDate>" + passport_date + "</PassportDate>";
                tmsg = tmsg + "<Company>" + dept + "</Company>";

                if (!String.IsNullOrEmpty(department))
                {
                    tmsg = tmsg + "<Department>" + department + "</Department>";
                }
                tmsg = tmsg + "<other>" + other + "</other>";

                tmsg = tmsg + "<Quit>0</Quit>";
                tmsg = tmsg + "<Isolate>0</Isolate>";


                tmsg = tmsg + "</query>";


                command5.Parameters.AddWithValue("DateCreated ", c_rev_date);
                command5.Parameters.AddWithValue("ServerName", servername);
                command5.Parameters.AddWithValue("Cmd", "ImportOneCardHolderXML");
                command5.Parameters.AddWithValue("Msg", tmsg);
                command5.Parameters.AddWithValue("Status", 0);

                command5.Parameters.AddWithValue("Result", 0);



                int result5 = command5.ExecuteNonQuery();
                // Check Error

                if (result5 < 0)
                {
                    Console.WriteLine("error insert id:" + num.ToString() + " data into DDS!");



                }
                else
                {
                    Console.WriteLine("insert id:" + num.ToString() + " data into DDS!");
                    MessageBox.Show("系統編號: " + tnext_no.ToString().Trim() + "已加人: " + txtstaff_no.Text.ToString().Trim() + " " + txtLastName.Text.ToString().Trim() + " " + txtFirstName.Text.ToString().Trim());
                    clean_all();

                };
                command5.Dispose();
                conn1.Close();

            }

            txtcard_no.Select();
        }
        private void find_c()
        {
            txtc_num.Text = "";
            txtc_id.Text = "";

            if (string.IsNullOrEmpty(txtfind.Text.Trim().ToUpper()) == true)
            {
                MessageBox.Show("請輸入 {身份證} 或 {名稱} 搜尋");
                return;
            }
            string find = txtfind.Text.Trim().ToUpper();
            using (OdbcConnection conn1 = new OdbcConnection(connection1))
            {
                conn1.Open();

                string query3 = "select h.*, c.code from CRDHLD h left join card c on h.id = c.owner where  (h.idcard like '%" + find + "%' or h.last_name like '%" + find + "%'or h.first_name like '%" + find + "%') and h.type <>3";
                query3 = query3 + " order by h.last_name, h.first_name, h.idcard ";
                OdbcCommand command1 = new OdbcCommand(query3, conn1);
                OdbcDataAdapter da3 = new OdbcDataAdapter(command1);

                //  DataTable dt3 = new DataTable();

                find_data.Clear();


                da3.Fill(find_data, "find_data");
                if (find_data.Tables["find_data"].Rows.Count >= 2)
                {
                    using (var form2 = new Form2())
                    {
                        var result = form2.ShowDialog();
                        if (result == DialogResult.OK)
                        {
                            string val = form2.ReturnValue;
                            query3 = "select h.*, c.code from CRDHLD h left join card c on h.id = c.owner where h.type <>3 and h.id = '" + val + "' ";
                            query3 = query3 + " order by h.last_name, h.first_name, h.idcard ";

                            command1 = new OdbcCommand(query3, conn1);
                            da3 = new OdbcDataAdapter(command1);
                            find_data.Clear();
                            da3.Fill(find_data, "find_data");



                        }
                    }
                }
                if (find_data.Tables["find_data"].Rows.Count == 1)
                {
                    foreach (DataRow dr in find_data.Tables["find_data"].Rows)
                    {
                        txtc_num.Text = dr["Num"].ToString();
                        txtc_id.Text = dr["Id"].ToString();

                        txtc_name.Text = dr["Last_Name"].ToString() + " " + dr["First_Name"].ToString();
                        txtc_staff_no.Text = dr["Office_Phone"].ToString();
                        txtc_staff_date.Text = dr["StaffDate"].ToString();
                        string idcard = dr["idcard"].ToString();
                        if (idcard.Length > 6)
                        {
                            txtc_idcard.Text ="****" + idcard.Substring(idcard.Length - 6);
                        }
                        else
                        {
                            txtc_idcard.Text = idcard;
                        }
                        txtc_dept.Text = dr["Company"].ToString();
                        if (txtc_dept.Text.Trim().ToUpper() == "FTC")
                        {
                            txtc_dept.BackColor = Color.FromArgb(255, 128, 0);
                        }
                        else
                        {
                            if (txtc_dept.Text.Trim().ToUpper() == "DDW")
                            {
                                txtc_dept.BackColor = Color.FromArgb(255, 204, 1);
                            }
                            else
                            {
                                if (txtc_dept.Text.Trim().ToUpper() == "ADC")
                                {
                                    txtc_dept.BackColor = Color.FromArgb(255, 128, 255);
                                }
                                else
                                {
                                    if (txtc_dept.Text.Trim().ToUpper() == "PH3 ENTRANCE" || txtc_dept.Text.Trim().ToUpper() == "PH3" || txtc_dept.Text.Trim().ToUpper() == "PHASE3 ENTRANCE")
                                    {
                                        txtc_dept.BackColor = Color.FromArgb(255, 192, 255);

                                    }
                                    else
                                    {
                                        if (txtc_dept.Text.Trim().ToUpper() == "SFC")
                                        {
                                            txtc_dept.BackColor = Color.FromArgb(135, 206, 250);
                                        }
                                        else
                                        {
                                            txtc_dept.BackColor = SystemColors.Window;
                                        }
                                    }
                                }
                            }
                        }
                        string c_to_date = dr["to_date"].ToString().Trim();

                        if (string.IsNullOrEmpty(c_to_date) == true)
                        {
                            c_to_date = "2999-12-31 23:59:59";

                        }
                        DateTime to_date = DateTime.ParseExact(c_to_date, "yyyy-MM-dd HH:mm:ss", null);

                        string c_datetime = get_server_time_conn(ConnectionString);

                        DateTime NowDate = DateTime.ParseExact(c_datetime, "yyyy-MM-dd HH:mm:ss", null);

                        String tstatus = "";
                        Boolean status_pass = true;
                        if (to_date <= NowDate)
                        {
                            tstatus = "過期";
                            status_pass = false;
                        }

                        if (dr["Valid"].ToString().Trim().ToUpper() == "FALSE")
                        {
                            if (!string.IsNullOrEmpty(tstatus))
                            {
                                tstatus = tstatus + " 和 ";

                            }
                            tstatus = "停用";
                            status_pass = false;

                        }


                        if (!string.IsNullOrEmpty(dr["code"].ToString().Trim()))
                        {
                            tstatus = "已有卡號:" + dr["code"].ToString().Trim();
                            status_pass = false;

                        }


                        if (status_pass == true)
                        {
                            tstatus = "有效";
                            txtstatus.BackColor = Color.White;
                            txtstatus.ForeColor = Color.Black;

                        }
                        else
                        {
                            txtstatus.BackColor = Color.Black;
                            txtstatus.ForeColor = Color.White;

                        }

                        txtstatus.Text = tstatus;



                        txtc_phone.Text = dr["Office_Phone"].ToString();

                        txtc_photo_file.Text = dr["Photo"].ToString().Trim();
                        txtc_other.Text = dr["other"].ToString().Trim();

                        txtc_passport_no.Text = dr["Passport"].ToString().Trim();
                        txtc_passport_date.Text = dr["PassportDate"].ToString().Trim();


                        pic_c_photo.ImageLocation = photopath + dr["Photo"].ToString().Trim();


                        txtc_holdcard.Text = dr["code"].ToString().Trim();


                    }
                }

                else
                {
                    if (find_data.Tables["find_data"].Rows.Count == 0)
                    {
                        MessageBox.Show("找不到 {身份證}");
                    }
                    return;
                }



                //    Console.WriteLine("check and get next_no: " + tnext_no.ToString());
                da3.Dispose();
                command1.Dispose();
                conn1.Dispose();
            }
            txtc_card_no.Select();

        }

        private void button6_Click_1(object sender, EventArgs e)
        {
            find_c();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string card_no = txtc_card_no.Text.Trim().ToUpper();
            if (string.IsNullOrEmpty(card_no) == true)
            {
                MessageBox.Show("未輸入卡號!");
                return;
            }
            card_no = card_no.PadLeft(8, '0');
            string name = txtc_name.Text.Trim().ToUpper();
            string staff_no = txtc_staff_no.Text.Trim();
            string staff_date = txtc_staff_date.Text.Trim();
            string c_from_date = txtc_staff_date.Text.Trim();
            string phone = txtc_phone.Text.Trim();

            string passport_no = txtc_passport_no.Text.Trim().ToUpper();
            string passport_date = txtc_passport_date.Text.Trim().ToUpper();
            string other = txtc_other.Text.Trim();
            string dept = txtc_dept.Text.Trim();
            string group_right = dept;
            if (group_right.Equals("DDW"))
            {
                group_right = "Hub-Badge(Vis)";
            }

            string card_reader = card_reader_name.Trim();

            string idcard = txtc_idcard.Text.Trim().ToUpper();


            string dep_name = txtc_dept.Text.Trim();

            string num = txtc_num.Text.Trim();

            int id = 0;
            Int32.TryParse(txtc_id.Text.Trim(), out id);

            if (string.IsNullOrEmpty(num))
            {
                if (id == 0)
                {
                    MessageBox.Show("沒有資料, 請按<搜尋>再尋找");
                    return;
                }
                num = put_new_num(id);
            }

            //string id = txtc_id.Text.Trim();

            string image_filename = txtc_photo_file.Text.Trim();

            Boolean treturn = true;
            if (string.IsNullOrEmpty(card_no.Trim()))
            {
                treturn = false;
                MessageBox.Show("必需 {抇卡} 或 輸入 {卡號}");
                return;
            }

            // check 

            string c_datetime = get_server_time_conn(ConnectionString);
            string c_to_date = "";

            using (OdbcConnection connection = new OdbcConnection(ConnectionString))
            {
                string query = "SELECT c.Code , h.*  FROM CRDHLD h , Card c where  h.id = c.Owner and c.code = '" + card_no + "'";


                OdbcCommand cmd = new OdbcCommand(query, connection);

                OdbcDataAdapter da = new OdbcDataAdapter(cmd);

                DataTable dt = new DataTable();
                da.Fill(dt);
                if (dt.Rows.Count > 0)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        MessageBox.Show("卡號: { " + dr["code"].ToString().Trim() + " } 已屬於 " + (dr["first_name"].ToString().Trim() + " " + dr["last_name"].ToString().Trim()).Trim(), "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }


                    return;
                }
                dt.Dispose();
                cmd.Dispose();
                da.Dispose();
                connection.Dispose();
            }



            using (OdbcConnection conn1 = new OdbcConnection(ConnectionString))
            {
                string query = "SELECT c.Code , h.*  FROM CRDHLD h , Card c where  h.id = c.Owner and h.num = '" + num + "'";

                OdbcCommand cmd = new OdbcCommand(query, conn1);

                OdbcDataAdapter da = new OdbcDataAdapter(cmd);

                DataTable dt = new DataTable();
                da.Fill(dt);
                if (dt.Rows.Count > 0)
                {

                    foreach (DataRow dr in dt.Rows)
                    {
                        MessageBox.Show("這人: { " + (dr["first_name"].ToString().Trim() + " " + dr["last_name"].ToString().Trim()).Trim() + " } 已擁有卡號: {" + dr["code"].ToString().Trim() + "} ", "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        break;
                    }
                    return;
                }

                dt.Dispose();
                da.Dispose();
                cmd.Dispose();
                conn1.Dispose();
            }


            using (OdbcConnection conn1 = new OdbcConnection(ConnectionString))
            {
                string query = "SELECT h.*  FROM CRDHLD h where h.num = '" + num + "'";

                OdbcCommand cmd = new OdbcCommand(query, conn1);

                OdbcDataAdapter da = new OdbcDataAdapter(cmd);

                DataTable dt = new DataTable();
                da.Fill(dt);
                if (dt.Rows.Count > 0)
                {

                    foreach (DataRow dr in dt.Rows)
                    {


                        //

                        c_to_date = dr["TO_Date"].ToString();
                        Console.WriteLine("ac_to_date: " + c_to_date);

                        if (string.IsNullOrEmpty(c_to_date) == true)
                        {
                            c_to_date = "2999-12-31 23:59:59";

                        }


                        DateTime to_date = DateTime.ParseExact(c_to_date, "yyyy-MM-dd HH:mm:ss", null);


                        DateTime NowDate = DateTime.ParseExact(c_datetime, "yyyy-MM-dd HH:mm:ss", null);

                        String tstatus = "";
                        Boolean status_pass = true;
                        if (to_date <= NowDate)
                        {
                            tstatus = "過期";
                            status_pass = false;
                        }

                        if (dr["Valid"].ToString().Trim().ToUpper() == "FALSE")
                        {
                            if (!string.IsNullOrEmpty(tstatus))
                            {
                                tstatus = tstatus + " 和 ";

                            }
                            tstatus = "停用";
                            status_pass = false;

                        }
                        if (status_pass == true)
                        {
                            tstatus = "有效";


                        }
                        else
                        {
                        }

                        //
                        if (status_pass == false)
                        {
                            MessageBox.Show("這人已 { " + tstatus + " }", "Information", MessageBoxButtons.OK, MessageBoxIcon.Error);

                            return;
                        }

                    }
                }

                dt.Dispose();
                da.Dispose();
                cmd.Dispose();
                conn1.Dispose();

            }




            // check end




            // add log

            string tphotopath = "";
            image_filename = "";
            if (is_update_photo == true)
            {
                get_photopath = pic_c_photo.ImageLocation;
                pic_c_photo.ImageLocation = "";
                if (!string.IsNullOrEmpty(get_photopath))
                {
                    File.Copy(get_photopath, photopath + @"" + num + ".jpg", true);
                    tphotopath = photopath + @"" + num + ".jpg";
                    image_filename = System.IO.Path.GetFileName(tphotopath);
                }
            }





            using (OdbcConnection conn2 = new OdbcConnection(ConnectionString2))
            {



                string query = "";




                query = "INSERT INTO dr_card_log ([event_dt],[borrow_return],[card_reader],[idcard],[name],[staff_no],[staff_date],[phone],[passport_no],[passport_date],[other],[dept],[card_no])";

                query = query + " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)";

                //
                //

                OdbcCommand command = new OdbcCommand(query, conn2);





                command.Parameters.AddWithValue("event_dt", c_datetime);
                command.Parameters.AddWithValue("borrow_return", "Borrow");

                command.Parameters.AddWithValue("card_reader", card_reader);
                command.Parameters.AddWithValue("idcard", idcard);


                command.Parameters.AddWithValue("name", name);

                command.Parameters.AddWithValue("staff_no", staff_no);

                command.Parameters.AddWithValue("staff_date", staff_date.ToString().Trim());

                command.Parameters.AddWithValue("phone", phone);


                command.Parameters.AddWithValue("passport_no", passport_no.ToString().Trim());

                command.Parameters.AddWithValue("passport_date", passport_date);
                command.Parameters.AddWithValue("other", other);
                command.Parameters.AddWithValue("dept", dept);
                command.Parameters.AddWithValue("card_no", card_no);
                conn2.Open();
                int result = command.ExecuteNonQuery();

                if (result < 0)
                {
                    Console.WriteLine("加不到借卡記錄, 身份證為:" + idcard.ToString() + " !");
                    MessageBox.Show("錯誤! 加不到借卡記錄, 身份證為: " + idcard.ToString() + " !");
                }
                else
                {
                    Console.WriteLine("加不到借卡記錄, 身份證為:" + idcard.ToString() + " !");

                };
                //
                command.Dispose();
                conn2.Dispose();
            }
            // add log end

            using (OdbcConnection conn1 = new OdbcConnection(connection1))
            {

                conn1.Open();




                string c_rev_date = get_server_time_conn(ConnectionString2);







                string query5 = "INSERT INTO QueueMSGAPI ([DateCreated], [ServerName], [Cmd], [Msg], [Status], [Result]) VALUES(?,?,?,?,?,?)";



                OdbcCommand command5 = new OdbcCommand(query5, conn1);




                String tmsg = "<query>";
                tmsg = tmsg + "<Number>" + num.ToString() + "</Number>";
                tmsg = tmsg + "<Last_Name>" + name + "</Last_Name>";
                tmsg = tmsg + "<First_Name>" + staff_no + "</First_Name>";
                if (string.IsNullOrEmpty(card_no.Trim()) == false)
                {
                    tmsg = tmsg + "<Badge>" + card_no + "</Badge>";
                    tmsg = tmsg + "<BorrowCardDate>" + c_datetime + "</BorrowCardDate>";

                }
                if (c_to_date.Equals("2999-12-31 23:59:59"))
                {
                    c_to_date = "";
                }
                //    tmsg = tmsg + "<From_Date>" + c_from_date + "</From_Date>";
                if (!string.IsNullOrEmpty(c_to_date))
                {
                    //    tmsg = tmsg + "<To_Date>" + c_to_date + "</To_Date>";
                }
                tmsg = tmsg + "<Company>" + dep_name + "</Company>";

                //
                //             tmsg = tmsg + "<Access_Group>" + group_right + "</Access_Group>";

                //            string type = "1";
                //            tmsg = tmsg + "<type>" + type + "</type>";
                //            tmsg = tmsg + "<Technology>3</Technology>";
                tmsg = tmsg + "<Description>1</Description>";


                if (!String.IsNullOrEmpty(image_filename))
                {
                    tmsg = tmsg + "<Photo>" + image_filename + "</Photo>";


                }
                //           tmsg = tmsg + "<Office_Phone>" + phone + "</Office_Phone>";


                //             tmsg = tmsg + "<idcard>" + idcard + "</idcard>";
                //             tmsg = tmsg + "<StaffDate>" + staff_date + "</StaffDate>";
                //            tmsg = tmsg + "<Passport>" + passport_no + "</Passport>";
                //            tmsg = tmsg + "<PassportDate>" + passport_date + "</PassportDate>";
                //        tmsg = tmsg + "<dept>" + dept + "</dept>";
                //            tmsg = tmsg + "<Company>" + dept + "</Company>";
                //            tmsg = tmsg + "<other>" + other + "</other>";


                tmsg = tmsg + "</query>";


                command5.Parameters.AddWithValue("DateCreated ", c_rev_date);
                command5.Parameters.AddWithValue("ServerName", servername);
                command5.Parameters.AddWithValue("Cmd", "ImportOneCardHolderXML");
                command5.Parameters.AddWithValue("Msg", tmsg);
                command5.Parameters.AddWithValue("Status", 0);

                command5.Parameters.AddWithValue("Result", 0);



                int result5 = command5.ExecuteNonQuery();
                // Check Error

                if (result5 < 0)
                {
                    Console.WriteLine("error insert id:" + num.ToString() + " data into DDS!");

                    MessageBox.Show("錯誤 加入 id:" + num.ToString() + " 去 DDS!");


                }
                else
                {
                    Console.WriteLine("insert id:" + num.ToString() + " data into DDS!");
                    MessageBox.Show("系統編號: " + num.ToString() + " 已加卡: " + card_no + "在" + txtc_staff_no.Text.ToString().Trim() + " " + txtc_name.Text.ToString().Trim());
                    clean_all();

                };
                command5.Dispose();

                conn1.Close();
                conn1.Dispose();
            }

            is_update_photo = false;
            labphoto_status.Text = "現在";
            labphoto_status.ForeColor = Color.FromArgb(0, 0, 0);
            txtfind.Select();
        }

        private void txtc_card_no_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtc_card_no_DoubleClick(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            clean_all();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            clean_all();

        }
        private void getdata_by_card()
        {

            string card_no = txtr_card_no.Text;
            if (string.IsNullOrEmpty(card_no))
            {
                return;
            }

            using (OdbcConnection conn1 = new OdbcConnection(connection1))
            {
                conn1.Open();




                string query4 = "select h.*, c.code as card_no from CRDHLD h, card c where c.owner = h.id and c.code = '" + card_no + "'";

                OdbcCommand command1 = new OdbcCommand(query4, conn1);
                OdbcDataAdapter da4 = new OdbcDataAdapter(command1);

                DataTable dt4 = new DataTable();
                da4.Fill(dt4);
                if (dt4.Rows.Count > 0)
                {
                    foreach (DataRow dr4 in dt4.Rows)
                    {
                        txtr_num.Text = dr4["Num"].ToString();
                        txtr_id.Text = dr4["Id"].ToString();

                        txtr_name.Text = dr4["Last_Name"].ToString() + " " + dr4["First_Name"].ToString();
                        txtr_staff_no.Text = dr4["Office_Phone"].ToString();
                        txtr_staff_date.Text = dr4["Start_Date"].ToString();
                        string idcard = dr4["idcard"].ToString();
                        if (idcard.Length > 6)
                        {
                            txtr_idcard.Text = "****" + idcard.Substring(idcard.Length - 6);
                        }
                        else
                        {
                            txtr_idcard.Text = idcard;
                        }
                        txtr_phone.Text = dr4["Office_Phone"].ToString();
                        txtr_dept.Text = dr4["Company"].ToString();

                        txtr_passport_no.Text = dr4["Passport"].ToString();
                        txtr_passport_date.Text = dr4["PassportDate"].ToString();

                        txtr_other.Text = dr4["other"].ToString();

                        if (txtr_dept.Text.Trim().ToUpper() == "FTC")
                        {
                            txtr_dept.BackColor = Color.FromArgb(255, 128, 0);
                        }
                        else
                        {
                            if (txtr_dept.Text.Trim().ToUpper() == "DDW")
                            {
                                txtr_dept.BackColor = Color.FromArgb(255, 204, 1);
                            }
                            else
                            {
                                if (txtr_dept.Text.Trim().ToUpper() == "ADC")
                                {
                                    txtr_dept.BackColor = Color.FromArgb(255, 128, 255);
                                }
                                else
                                {
                                    if (txtr_dept.Text.Trim().ToUpper() == "PH3 ENTRANCE" || txtr_dept.Text.Trim().ToUpper() == "PH3" || txtr_dept.Text.Trim().ToUpper() == "PHASE3 ENTRANCE")
                                    {
                                        txtr_dept.BackColor = Color.FromArgb(255, 192, 255);
                                    }
                                    else
                                    {
                                        if (txtr_dept.Text.Trim().ToUpper() == "SFC")
                                        {
                                            txtr_dept.BackColor = Color.FromArgb(135, 206, 250);
                                        }
                                        else
                                        {
                                            txtr_dept.BackColor = SystemColors.Window;
                                        }
                                    }
                                }
                            }
                        }
                        txtr_photo_file.Text = dr4["Photo"].ToString().Trim();
                        pic_r_photo.ImageLocation = photopath + dr4["Photo"].ToString().Trim();





                    }
                    dt4.Dispose();
                    command1.Dispose();
                    conn1.Dispose();
                }

                else
                {
                    MessageBox.Show("這卡號: " + card_no + " 沒有登記人");
                    txtr_card_no.Text = "";
                    clean_all();

                    return;
                }

                dt4.Dispose();
                da4.Dispose();
                command1.Dispose();
                conn1.Dispose();

                //    Console.WriteLine("check and get next_no: " + tnext_no.ToString());
            }
        }
        private void getdata_by_card2()
        {

            string card_no = txtl_card_no.Text;
            if (string.IsNullOrEmpty(card_no))
            {
                return;
            }

            using (OdbcConnection conn1 = new OdbcConnection(connection1))
            {
                conn1.Open();




                string query4 = "select h.*, c.code as card_no from CRDHLD h, card c where c.owner = h.id and c.code = '" + card_no + "'";

                OdbcCommand command1 = new OdbcCommand(query4, conn1);
                OdbcDataAdapter da4 = new OdbcDataAdapter(command1);

                DataTable dt4 = new DataTable();
                da4.Fill(dt4);
                if (dt4.Rows.Count > 0)
                {
                    foreach (DataRow dr4 in dt4.Rows)
                    {
                        txtl_num.Text = dr4["Num"].ToString();
                        txtl_id.Text = dr4["Id"].ToString();

                        txtl_name.Text = dr4["Last_Name"].ToString();
                        txtl_staff_no.Text = dr4["First_Name"].ToString();
                        txtl_idcard.Text = dr4["idcard"].ToString();
                        txtl_phone.Text = dr4["Office_Phone"].ToString();
                        txtl_passport_no.Text = dr4["passport"].ToString();
                        txtl_passport_date.Text = dr4["passportdate"].ToString();
                        txtl_other.Text = dr4["other"].ToString();
                        txtl_dept.Text = dr4["company"].ToString();
                        txtl_staff_date.Text = dr4["StaffDate"].ToString();
                        if (txtl_dept.Text.Trim().ToUpper() == "FTC")
                        {
                            txtl_dept.BackColor = Color.FromArgb(255, 128, 0);
                        }
                        else
                        {
                            if (txtl_dept.Text.Trim().ToUpper() == "DDW")
                            {
                                txtl_dept.BackColor = Color.FromArgb(255, 204, 1);
                            }
                            else
                            {
                                if (txtl_dept.Text.Trim().ToUpper() == "ADC")
                                {
                                    txtl_dept.BackColor = Color.FromArgb(255, 128, 255);
                                }
                                else
                                {
                                    if (txtl_dept.Text.Trim().ToUpper() == "PH3 ENTRANCE" || txtl_dept.Text.Trim().ToUpper() == "PH3" || txtl_dept.Text.Trim().ToUpper() == "PHASE3 ENTRANCE")
                                    {
                                        txtl_dept.BackColor = Color.FromArgb(255, 192, 255);
                                    }
                                    else
                                    {
                                        if (txtl_dept.Text.Trim().ToUpper() == "SFC")
                                        {
                                            txtl_dept.BackColor = Color.FromArgb(135, 206, 250);
                                        }
                                        else
                                        {
                                            txtl_dept.BackColor = SystemColors.Window;
                                        }
                                    }
                                }
                            }
                        }

                    }
                    dt4.Dispose();
                    da4.Dispose();
                    command1.Dispose();
                    conn1.Dispose();
                }

                else
                {
                    MessageBox.Show("這卡號: " + card_no + " 沒有登記人");
                    txtr_card_no.Text = "";
                    clean_all();

                    return;
                }



                //    Console.WriteLine("check and get next_no: " + tnext_no.ToString());
            }
        }
        private void button8_Click(object sender, EventArgs e)

        {
            string c_datetime = get_server_time_conn(ConnectionString);

            DateTime oDate = DateTime.ParseExact(c_datetime, "yyyy-MM-dd HH:mm:ss", null);
            oDate = oDate.AddSeconds(-5);

            //  action_time = DateTime.Now.AddSeconds(-10).ToString("yyyy-MM-dd HH:mm:ss");

            string check_datetime = oDate.ToString("yyyy-MM-dd HH:mm:ss");
            using (OdbcConnection connection = new OdbcConnection(ConnectionString))
            {
                string query = "select TOP 1 * FROM LOGt where (trn_type = 1 or trn_type = 61 or trn_type = 63)  and [From_Name] = '" + card_reader_name.Trim() + "' and date >= '" + check_datetime + "' and date > '" + card_record_dt + "' order by  id desc";
                //string query = "select TOP 1 * FROM LOGt where (trn_type = 1 or trn_type = 61 or trn_type = 63)  and [From_Name] = '" + card_reader_name.Trim() + "' and date >= '" + check_datetime + "' order by date desc";
                OdbcCommand cmd = new OdbcCommand(query, connection);
                OdbcDataAdapter da = new OdbcDataAdapter(cmd);

                DataTable dt = new DataTable();

                da.Fill(dt);
                string desc3 = "";
                string number = "";
                string name = "";
                string card_no = "";

                int trn_type = 0;
                if (dt.Rows.Count <= 0)
                {
                    return;
                }
                foreach (DataRow dr in dt.Rows)
                {


                    int card_record_id = int.Parse(dr["ID"].ToString());

                    card_record_dt = dr["Date"].ToString().Trim();

                    trn_type = int.Parse(dr["Trn_type"].ToString());
                    Console.WriteLine("trn_type: " + trn_type.ToString());
                    if (trn_type == 61 || trn_type == 63)
                    {

                        //card_no = (int.Parse(dr["desc3"].ToString().Trim()) - 6291456).ToString().Trim().PadLeft(6, '0');
                        card_no = (int.Parse(dr["desc3"].ToString().Trim())).ToString().Trim().PadLeft(8, '0');
                        txtr_card_no.Text = card_no;

                    }
                    else
                    {
                        desc3 = dr["desc3"].ToString().Trim();
                        int find_at = desc3.LastIndexOf(' ');

                        number = desc3.Substring(find_at + 1, (desc3.Length - find_at - 1));
                        //  System.Diagnostics.Debug.WriteLine(number);
                        query = "select TOP 1 h.last_name,c.code,free_field1 as gender FROM CRDHLD h, card c where h.id = c.owner and h.num = '" + number + "'";
                        //  MessageBox.Show(query);
                        OdbcCommand cmd3 = new OdbcCommand(query, connection);
                        OdbcDataAdapter da3 = new OdbcDataAdapter(cmd3);
                        DataTable dt3 = new DataTable();
                        da3.Fill(dt3);

                        foreach (DataRow dr3 in dt3.Rows)
                        {
                            //card_no = (int.Parse(dr3["Code"].ToString()) - 6291456).ToString().Trim().PadLeft(6, '0');
                            card_no = (int.Parse(dr3["Code"].ToString())).ToString().Trim().PadLeft(8, '0');
                            txtr_card_no.Text = card_no;


                        }
                    }

                    //




                    getdata_by_card();




                }

                dt.Dispose();
                da.Dispose();
                cmd.Dispose();
                connection.Dispose();
            }

        }

        private void button11_Click(object sender, EventArgs e)
        {
            string c_datetime = get_server_time_conn(ConnectionString);

            string num = txtr_num.Text.Trim();



            int id = 0;
            Int32.TryParse(txtr_id.Text.Trim(), out id);

            if (string.IsNullOrEmpty(num))
            {
                if (id == 0)
                {
                    MessageBox.Show("沒有資料, 請按<搜尋>再尋找");
                    return;
                }
                num = put_new_num(id);
            }

            string card_no = txtr_card_no.Text.Trim().ToUpper();
            if (string.IsNullOrEmpty(card_no))
            {
                card_no = card_no.PadLeft(8, '0');
            }

            string image_filename = txtr_photo_file.Text.Trim();
            string name = txtr_name.Text.Trim().ToUpper();
            string staff_no = txtr_staff_no.Text.Trim();

            string dep_name = txtr_dept.Text.Trim();
            string idcard = txtr_idcard.Text.Trim();

            string c_rev_date = get_server_time_conn(ConnectionString2);
            string card_reader = card_reader_name.Trim();

            string staff_date = txtr_staff_date.Text.Trim();

            string passport_no = txtr_passport_no.Text.Trim();

            string passport_date = txtr_passport_date.Text.Trim();

            string phone = txtr_phone.Text.Trim();

            string other = txtr_other.Text.Trim();

            string dept = txtr_dept.Text.Trim();

            using (OdbcConnection conn1 = new OdbcConnection(connection1))
            {


                conn1.Open();

                string query5 = "INSERT INTO QueueMSGAPI ([DateCreated], [ServerName], [Cmd], [Msg], [Status], [Result]) VALUES(?,?,?,?,?,?)";



                OdbcCommand command5 = new OdbcCommand(query5, conn1);




                String tmsg = "<query>";
                tmsg = tmsg + "<Number>" + num.ToString() + "</Number>";

                tmsg = tmsg + "<Badge>" + "</Badge>";
                tmsg = tmsg + "<BorrowCardDate>.</BorrowCardDate>";

                tmsg = tmsg + "</query>";


                command5.Parameters.AddWithValue("DateCreated ", c_rev_date);
                command5.Parameters.AddWithValue("ServerName", servername);
                command5.Parameters.AddWithValue("Cmd", "ImportOneCardHolderXML");
                command5.Parameters.AddWithValue("Msg", tmsg);
                command5.Parameters.AddWithValue("Status", 0);

                command5.Parameters.AddWithValue("Result", 0);



                int result5 = command5.ExecuteNonQuery();
                // Check Error

                if (result5 < 0)
                {
                    Console.WriteLine("error insert id:" + num.ToString() + " data into DDS!");



                }
                else
                {
                    Console.WriteLine("insert id:" + num.ToString() + " data into DDS!");
                    MessageBox.Show("系統編號: " + num.ToString() + " 已消證: " + card_no + "在" + txtr_staff_no.Text.ToString().Trim() + " " + txtr_name.Text.ToString().Trim());
                    clean_all();

                };
                command5.Dispose();
                conn1.Close();
                conn1.Dispose();
            }

            // add log
            using (OdbcConnection conn2 = new OdbcConnection(ConnectionString2))
            {



                string query = "";




                query = "INSERT INTO dr_card_log ([event_dt],[borrow_return],[card_reader],[idcard],[name],[staff_no],[staff_date],[phone],[passport_no],[passport_date],[other],[dept],[card_no])";

                query = query + " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)";

                //
                //

                OdbcCommand command = new OdbcCommand(query, conn2);





                command.Parameters.AddWithValue("event_dt", c_datetime);
                command.Parameters.AddWithValue("borrow_return", "Return");

                command.Parameters.AddWithValue("card_reader", card_reader);
                command.Parameters.AddWithValue("idcard", idcard);


                command.Parameters.AddWithValue("name", name);

                command.Parameters.AddWithValue("staff_no", staff_no);

                command.Parameters.AddWithValue("staff_date", staff_date.ToString().Trim());

                command.Parameters.AddWithValue("phone", phone);


                command.Parameters.AddWithValue("passport_no", passport_no.ToString().Trim());

                command.Parameters.AddWithValue("passport_date", passport_date);
                command.Parameters.AddWithValue("other", other);
                command.Parameters.AddWithValue("dept", dept);
                command.Parameters.AddWithValue("card_no", card_no);
                conn2.Open();
                int result = command.ExecuteNonQuery();

                if (result < 0)
                {
                    Console.WriteLine("加不到借卡記錄, 身份證為:" + idcard.ToString() + " !");
                    MessageBox.Show("錯誤! 加不到借卡記錄, 身份證為: " + idcard.ToString() + " !");
                }
                else
                {
                    Console.WriteLine("加不到借卡記錄, 身份證為:" + idcard.ToString() + " !");

                };
                //
                command.Dispose();
                conn2.Dispose();
            }

            // add log end




            txtr_card_no.Select();

        }
        private void find_r()
        {
            txtr_num.Text = "";
            txtr_id.Text = "";
            if (txtr_card_no.ReadOnly == false)
            {
                if (string.IsNullOrEmpty(txtr_card_no.Text) == true)
                {

                    return;
                }
                if (!string.IsNullOrEmpty(txtr_card_no.Text))
                {
                    txtr_card_no.Text = txtr_card_no.Text.PadLeft(8, '0');
                }
                getdata_by_card();
            }
        }
        private void txtr_card_no_Validated(object sender, EventArgs e)
        {
            find_r();
        }

        private void txtr_card_no_DoubleClick(object sender, EventArgs e)
        {
            //       if (txtr_card_no.ReadOnly == true)
            //       {
            //           txtr_card_no.ReadOnly = false;
            //       }
            //       else
            //      {
            //          txtr_card_no.ReadOnly = true;
            //      }
        }

        private void tabPage4_Click(object sender, EventArgs e)
        {

        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }
        private string put_new_num(int id)
        {
            if (!(id > 0))
            {
                return "";
            }
            int new_num = get_new_num();
            using (OdbcConnection conn1 = new OdbcConnection(connection1))
            {
                conn1.Open();
                string query = "update CRDHLD set num = ? where id = ?";

                OdbcCommand command = new OdbcCommand(query, conn1);


                command.Parameters.AddWithValue("num", new_num);
                command.Parameters.AddWithValue("id", id);


                int result = command.ExecuteNonQuery();

                if (result < 0)
                {
                    MessageBox.Show("Error, can not save");
                }
                command.Dispose();
                conn1.Dispose();
            }

            return new_num.ToString();
        }

        private int get_new_num()
        {

            int tnext_no = 1;
            Boolean add_card = false;

            string c_rev_date = get_server_time_conn(connection1);

            using (OdbcConnection conn2 = new OdbcConnection(connection2))
            {
                conn2.Open();
                string query = "select next_no FROM sys_next_no where sys_name = 'CARDHLD' ";
                OdbcCommand cmd = new OdbcCommand(query, conn2);




                var reader = cmd.ExecuteReader();
                int f_next_no = reader.GetOrdinal("next_no");
                if (!reader.Read())
                    throw new InvalidOperationException("No records were returned.");

                tnext_no = reader.GetInt32(f_next_no);
                Console.WriteLine("get next_no: " + tnext_no.ToString());
                cmd.Dispose();
                conn2.Dispose();
            }





            using (OdbcConnection conn1 = new OdbcConnection(connection1))
            {
                conn1.Open();

                OdbcCommand command1;
                OdbcDataAdapter da3;
                DataTable dt3 = new DataTable();
                while (true)
                {



                    string c_tnext_no = tnext_no.ToString();

                    string query3 = "select h.num from CRDHLD h where  h.num = '" + c_tnext_no + "'";

                    command1 = new OdbcCommand(query3, conn1);
                    da3 = new OdbcDataAdapter(command1);

                    dt3.Clear();
                    da3.Fill(dt3);
                    if (dt3.Rows.Count > 0)
                    {

                        tnext_no = tnext_no + 1;

                    }
                    else
                    {
                        break;
                    }


                }
                dt3.Dispose();
                da3.Dispose();
                command1.Dispose();
                conn1.Dispose();
                Console.WriteLine("check and get next_no: " + tnext_no.ToString());
            }
            using (OdbcConnection conn2 = new OdbcConnection(connection2))
            {
                conn2.Open();
                string query4 = "update sys_next_no set next_no = " + (tnext_no + 1) + " where sys_name = 'CARDHLD' ";
                OdbcCommand cmd4 = new OdbcCommand(query4, conn2);




                var reader4 = cmd4.ExecuteNonQuery();

                Console.WriteLine("update next_no: " + (tnext_no + 1).ToString());

                cmd4.Dispose();
                conn2.Dispose();
            }

            string num = tnext_no.ToString().Trim().ToUpper();
            return tnext_no;


        }
        private void button9_Click(object sender, EventArgs e)
        {
            clean_all();
            is_update_photo = false;

            labphoto_status.Text = "現在";
            labphoto_status.ForeColor = Color.FromArgb(0, 0, 0);

        }

        private void timer_get_c_card_no_Tick(object sender, EventArgs e)
        {
            Console.WriteLine("timer_get_c_card_no_Tick");

            string c_card_no = read_card();

            if (string.IsNullOrEmpty(c_card_no) == false)
            {
                if (string.IsNullOrEmpty(txtc_card_no.Text.Trim()) == true)
                {
                    Console.WriteLine("NOT Focused");


                    txtc_card_no.Text = c_card_no.Trim();
                }
                else
                {
                    Console.WriteLine("txtc_card_no.Focused:" + txtc_card_no.Focused.ToString());
                    if (txtc_card_no.Focused == true)
                    {
                        Console.WriteLine("Focused");

                        txtc_card_no.Text = c_card_no.Trim();
                    }

                }


            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void timer_get_card_no_Tick(object sender, EventArgs e)
        {
            Console.WriteLine("timer_get_card_no_Tick");
            string card_no = read_card();
            Console.WriteLine("get card_no: " + card_no);
            if (string.IsNullOrEmpty(card_no) == false)
            {
                if (string.IsNullOrEmpty(txtcard_no.Text.Trim()) == true)
                {
                    Console.WriteLine("NOT Focused");


                    txtcard_no.Text = card_no.Trim();
                }
                else
                {
                    if (txtc_card_no.Focused == true)
                    {
                        Console.WriteLine("Focused");

                        txtcard_no.Text = card_no.Trim();
                    }

                }


            }
        }

        private void timer_get_r_card_no_Tick(object sender, EventArgs e)
        {
            Console.WriteLine("timer_get_r_card_no_Tick");

            string r_card_no = read_card();

            if (string.IsNullOrEmpty(r_card_no) == false)
            {
                if (string.IsNullOrEmpty(txtr_card_no.Text.Trim()) == true)
                {
                    Console.WriteLine("NOT Focused");


                    txtr_card_no.Text = r_card_no.Trim();
                    getdata_by_card();
                }
                else
                {
                    Console.WriteLine("txtr_card_no.Focused:" + txtr_card_no.Focused.ToString());
                    if (txtr_card_no.Focused == true)
                    {
                        Console.WriteLine("Focused");

                        txtr_card_no.Text = r_card_no.Trim();
                        getdata_by_card();
                    }

                }


            }
        }

        private void txtcard_no_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtidcard_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Check if text is selected in the TextBox
            if (txtidcard.SelectionLength > 0)
            {
                return; // Allow keypress event for selected text
            }

            // Ensure the first character is a letter
            if (txtidcard.Text.Length == 0 && !char.IsLetter(e.KeyChar))
            {
                e.Handled = true; // Reject the keypress event
                return;
            }

            // Ensure the remaining characters are digits
            if (txtidcard.Text.Length > 0 && !char.IsDigit(e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true; // Reject the keypress event
                return;
            }

            // Limit the input to a maximum of seven digits
            if (txtidcard.Text.Length >= 7 && e.KeyChar != '\b')
            {
                e.Handled = true; // Reject the keypress event
                return;
            }
        }

        private void txtidcard_code_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Check if text is selected in the TextBox
            if (txtidcard_code.SelectionLength > 0)
            {
                return; // Allow keypress event for selected text
            }

            // only accept digit
            if (!char.IsDigit(e.KeyChar) && !char.IsControl(e.KeyChar))
            {
                e.Handled = true; // Reject the keypress event
            }

            // only accept one digit
            if (txtidcard_code.Text.Length > 0 && e.KeyChar != '\b')
            {
                e.Handled = true; // Reject the keypress event
                return;
            }
        }

        private void pic_photo_Click(object sender, EventArgs e)
        {

        }

        private void button5_Click_1(object sender, EventArgs e)
        {
            string c_datetime = get_server_time_conn(ConnectionString);


            string card_no = txtl_card_no.Text.Trim().ToUpper();
            string num = txtl_num.Text.Trim();


            if (string.IsNullOrEmpty(card_no))
            {
                card_no = card_no.PadLeft(8, '0');
            }

            int id = 0;
            Int32.TryParse(txtl_id.Text.Trim(), out id);

            if (string.IsNullOrEmpty(num))
            {
                if (id == 0)
                {
                    MessageBox.Show("沒有資料, 請按<搜尋>再尋找");
                    return;
                }
                num = put_new_num(id);
            }

            string image_filename = txtr_photo_file.Text.Trim();
            string name = txtl_name.Text.Trim().ToUpper();
            string staff_no = txtl_staff_no.Text.Trim();

            string idcard = txtl_idcard.Text;

            string c_rev_date = get_server_time_conn(ConnectionString2);
            string card_reader = card_reader_name.Trim();
            string phone = txtl_phone.Text.Trim();
            string passport_no = txtl_passport_no.Text.Trim();
            string passport_date = txtl_passport_date.Text.Trim();
            string dept = txtl_dept.Text.Trim();
            string staff_date = txtl_staff_date.Text.Trim();

            string other = txtl_other.Text.Trim();








            using (OdbcConnection conn1 = new OdbcConnection(connection1))
            {


                conn1.Open();

                string query5 = "INSERT INTO QueueMSGAPI ([DateCreated], [ServerName], [Cmd], [Msg], [Status], [Result]) VALUES(?,?,?,?,?,?)";



                OdbcCommand command5 = new OdbcCommand(query5, conn1);




                String tmsg = "<query>";
                tmsg = tmsg + "<Number>" + num.ToString() + "</Number>";

                tmsg = tmsg + "<Badge>" + "</Badge>";
                tmsg = tmsg + "<BorrowCardDate>.</BorrowCardDate>";

                tmsg = tmsg + "</query>";


                command5.Parameters.AddWithValue("DateCreated ", c_rev_date);
                command5.Parameters.AddWithValue("ServerName", servername);
                command5.Parameters.AddWithValue("Cmd", "ImportOneCardHolderXML");
                command5.Parameters.AddWithValue("Msg", tmsg);
                command5.Parameters.AddWithValue("Status", 0);

                command5.Parameters.AddWithValue("Result", 0);



                int result5 = command5.ExecuteNonQuery();
                // Check Error

                if (result5 < 0)
                {
                    Console.WriteLine("error insert id:" + num.ToString() + " data into DDS!");



                }
                else
                {
                    Console.WriteLine("insert id:" + num.ToString() + " data into DDS!");


                };
                command5.Dispose();
                conn1.Close();
                conn1.Dispose();
            }



            // add log
            using (OdbcConnection conn2 = new OdbcConnection(ConnectionString2))
            {



                string query = "";




                query = "INSERT INTO dr_card_log ([event_dt],[borrow_return],[card_reader],[idcard],[name],[staff_no],[staff_date],[phone],[passport_no],[passport_date],[other],[dept],[card_no],[is_lost])";

                query = query + " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?)";

                //
                // log

                OdbcCommand command = new OdbcCommand(query, conn2);





                command.Parameters.AddWithValue("event_dt", c_datetime);
                command.Parameters.AddWithValue("borrow_return", "Lost");

                command.Parameters.AddWithValue("card_reader", card_reader);
                command.Parameters.AddWithValue("idcard", idcard);


                command.Parameters.AddWithValue("name", name);

                command.Parameters.AddWithValue("staff_no", staff_no);

                command.Parameters.AddWithValue("staff_date", staff_date.ToString().Trim());

                command.Parameters.AddWithValue("phone", phone);


                command.Parameters.AddWithValue("passport_no", passport_no.ToString().Trim());

                command.Parameters.AddWithValue("passport_date", passport_date);
                command.Parameters.AddWithValue("other", other);
                command.Parameters.AddWithValue("dept", dept);
                command.Parameters.AddWithValue("card_no", card_no);
                command.Parameters.AddWithValue("is_lost", true);

                conn2.Open();
                int result = command.ExecuteNonQuery();

                if (result < 0)
                {
                    Console.WriteLine("加不到借卡記錄, 身份證為:" + idcard.ToString() + " !");
                    MessageBox.Show("錯誤! 加不到借卡記錄, 身份證為: " + idcard.ToString() + " !");
                }
                else
                {

                    MessageBox.Show("已記錄報失及消號 卡號: " + idcard.ToString() + " !");
                    clean_all();
                };
                //

                command.Dispose();
                conn2.Dispose();
            }
            // add log end
            //


            //





            //

            txtl_card_no.Select();
        }
        private void find_l()
        {
            txtl_num.Text = "";
            txtl_id.Text = "";

            if (!string.IsNullOrEmpty(txtl_card_no.Text))
            {
                txtl_card_no.Text = txtl_card_no.Text.PadLeft(8, '0');
            }
            string card_no = txtl_card_no.Text;
            getdata_by_card2();
        }
        private void txtl_card_no_Validated(object sender, EventArgs e)
        {
            find_l();
        }
        public static bool IsValidDate(string inputString)
        {
            string[] formats = { "yyyy/M/d", "yyyy/MM/d", "yyyy/M/dd", "yyyy/MM/dd", "yyyy-M-d", "yyyy-M-dd", "yyyy-MM-d", "yyyy-MM-dd" };
            DateTime parsedDate;
            var isValidFormat = DateTime.TryParseExact(inputString, formats, new CultureInfo("en-US"), DateTimeStyles.None, out parsedDate);

            if (isValidFormat)
            {
                string.Format("{0:yyyy/MM/d}", parsedDate);
                return true;

            }
            else
            {
                return false;
            }
        }
        private void button8_Click_1(object sender, EventArgs e)
        {

            string twhere = "";
            if (!String.IsNullOrEmpty(txtfind_data.Text.Trim()))
            {


                if (!String.IsNullOrEmpty(twhere.Trim()))
                {
                    twhere = twhere + " and ";
                }

                twhere = twhere + " (";
                if (IsValidDate(txtfind_data.Text))
                {

                    //      MessageBox.Show(IsValidDate(txtfind.Text).ToString());

                    DateTime fromdate = DateTime.Parse(txtfind_data.Text.Trim());
                    string fromdatetime = fromdate.ToString("yyyy-MM-dd");

                    DateTime todate = DateTime.Parse(txtfind_data.Text.Trim()).AddDays(1);
                    string todatetime = todate.ToString("yyyy-MM-dd");

                    twhere = twhere + " (event_dt >= '" + fromdate + "' and event_dt <  '" + todatetime + "' ) or ";


                }


                twhere = twhere + " idcard like '%" + txtfind_data.Text.Trim() + "%' ";
                twhere = twhere + " or ";
                twhere = twhere + " name like '%" + txtfind_data.Text.Trim() + "%' ";
                twhere = twhere + " or ";
                twhere = twhere + " card_no like '%" + txtfind_data.Text.Trim() + "%' ";
                twhere = twhere + " or ";
                twhere = twhere + " phone like '%" + txtfind_data.Text.Trim() + "%' ";
                twhere = twhere + " or ";
                twhere = twhere + " passport_no like '%" + txtfind_data.Text.Trim() + "%' ";
                twhere = twhere + " or ";
                twhere = twhere + " other like '%" + txtfind_data.Text.Trim() + "%' ";


                twhere = twhere + " ) ";


            }

            if (!String.IsNullOrEmpty(txtdate1.Text.Trim()))
            {
                if (!String.IsNullOrEmpty(twhere.Trim()))
                {
                    twhere = twhere + " and ";
                }


                twhere = twhere + " (event_dt >=  '" + txtdate1.Text.Trim() + "') ";


            }
            if (!String.IsNullOrEmpty(txtdate2.Text.Trim()))
            {
                if (!String.IsNullOrEmpty(twhere.Trim()))
                {
                    twhere = twhere + " and ";
                }
                // string cdatetime = (DateTime.Parse(txtdate2.Text.Trim(), "yyyy-MM-dd HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture).AddDays(1)).ToShortDateString();
                DateTime tdate = DateTime.Parse(txtdate2.Text.Trim()).AddDays(1);
                string cdatetime = tdate.ToString("yyyy-MM-dd");
                twhere = twhere + " (event_dt <  '" + cdatetime + "') ";

            }


            if (!string.IsNullOrEmpty(comtype.Text.Trim()))
            {
                if (comtype.Text.Equals("加人"))
                {
                    if (!String.IsNullOrEmpty(twhere.Trim()))
                    {
                        twhere = twhere + " and ";
                    }
                    twhere = twhere + " borrow_return = 'Add_Borrow'";
                }
                if (comtype.Text.Equals("加證"))
                {
                    if (!String.IsNullOrEmpty(twhere.Trim()))
                    {
                        twhere = twhere + " and ";
                    }
                    twhere = twhere + " borrow_return = 'Borrow'";
                }
                if (comtype.Text.Equals("消證"))
                {
                    if (!String.IsNullOrEmpty(twhere.Trim()))
                    {
                        twhere = twhere + " and ";
                    }
                    twhere = twhere + " borrow_return = 'Return'";
                }
                if (comtype.Text.Equals("報失"))
                {
                    if (!String.IsNullOrEmpty(twhere.Trim()))
                    {
                        twhere = twhere + " and ";
                    }
                    twhere = twhere + " borrow_return = 'Lost'";
                }
                if (comtype.Text.Equals("停用"))
                {
                    if (!String.IsNullOrEmpty(twhere.Trim()))
                    {
                        twhere = twhere + " and ";
                    }
                    twhere = twhere + " borrow_return = 'StopCard'";
                }
                if (comtype.Text.Equals("生效"))
                {
                    if (!String.IsNullOrEmpty(twhere.Trim()))
                    {
                        twhere = twhere + " and ";
                    }
                    twhere = twhere + " borrow_return = 'ActCard'";
                }
            }

            string query = "";
            if (!String.IsNullOrEmpty(twhere.Trim()))
            {
                twhere = " where " + twhere;
            }
            query = "SELECT  * FROM dr_card_log " + twhere + " order by event_dt desc";
            //    MessageBox.Show(query);
            Console.WriteLine(query);
            using (OdbcConnection connection = new OdbcConnection(ConnectionString2))
            {
                connection.Open();

                OdbcCommand command = new OdbcCommand(query, connection);


                OdbcDataAdapter DA = new OdbcDataAdapter(command);

                dr_card_log.Clear();
                DA.Fill(dr_card_log, "dr_card_log");
                g_dr_card.AutoGenerateColumns = false;
                g_dr_card.DataSource = dr_card_log;
                g_dr_card.DataMember = "dr_card_log";
                foreach (DataGridViewRow dr in g_dr_card.Rows)
                {
                    if (dr.Cells["Column6"].Value.ToString().Trim().ToUpper().Equals("ADD_BORROW"))
                    {

                        dr.Cells["Column6"].Value = "加人";
                    }
                    if (dr.Cells["Column6"].Value.ToString().Trim().ToUpper().Equals("BORROW"))
                    {

                        dr.Cells["Column6"].Value = "加證";
                    }
                    if (dr.Cells["Column6"].Value.ToString().Trim().ToUpper().Equals("RETURN"))
                    {

                        dr.Cells["Column6"].Value = "消證";
                    }
                    if (dr.Cells["Column6"].Value.ToString().Trim().ToUpper().Equals("LOST"))
                    {

                        dr.Cells["Column6"].Value = "報失";
                    }
                    if (dr.Cells["Column6"].Value.ToString().Trim().ToUpper().Equals("STOPCARD"))
                    {

                        dr.Cells["Column6"].Value = "停用";
                    }
                    if (dr.Cells["Column6"].Value.ToString().Trim().ToUpper().Equals("ACTCARD"))
                    {

                        dr.Cells["Column6"].Value = "生效";
                    }
                    if (dr.Cells["Column8"].Value.ToString().Trim().ToUpper().Equals("TRUE"))
                    {

                        dr.Cells["Column7"].Value = "是";
                    }
                    else
                    {
                        dr.Cells["Column7"].Value = "否";

                    }

                }
                DA.Dispose();
                command.Dispose();
                connection.Dispose();
            }

        }
        private void free_memory()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
        private void tabPage6_Click(object sender, EventArgs e)
        {

        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void button12_Click(object sender, EventArgs e)
        {
            clean_all();
        }

        private void txtr_card_no_KeyUp(object sender, KeyEventArgs e)
        {

            if (e.KeyCode == Keys.Enter)
            {

                if (sender is TextBox)
                {
                    find_r();
                }
            }
        }

        private void txtfind_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                if (sender is TextBox)
                {
                    find_c();
                }
            }
        }

        private void txtfind_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtLastName_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtFirstName_TextChanged(object sender, EventArgs e)
        {

        }

        private void button13_Click(object sender, EventArgs e)
        {
            find_r();
        }

        private void txtr_card_no_TextChanged(object sender, EventArgs e)
        {

        }

        private void button14_Click(object sender, EventArgs e)
        {
            find_l();
        }

        private void txtl_card_no_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {

                if (sender is TextBox)
                {
                    find_l();
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void button15_Click(object sender, EventArgs e)
        {
            //      MessageBox.Show(  g_dr_card.Rows[0].Cells[0].ToString());
            //         MessageBox.Show(g_dr_card.Rows[0].Cells[1].ToString());
            //         MessageBox.Show(g_dr_card.Rows[0].Cells[2].ToString());
            //         MessageBox.Show(g_dr_card.Rows[0].Cells[3].ToString());

        }
        private void find_not_return_card()
        {
            string twhere2 = "";

            if (!String.IsNullOrEmpty(comvalid.Text.Trim()))
            {
                if (!String.IsNullOrEmpty(twhere2.Trim()))
                {
                    twhere2 = twhere2 + " and ";
                }
                if (comvalid.Text.Trim().Equals("有效"))
                {

                    twhere2 = twhere2 + " h.Valid = 1  ";
                }
                if (comvalid.Text.Trim().Equals("冇效"))
                {

                    twhere2 = twhere2 + " h.Valid = 0  ";
                }

            }
            if (!String.IsNullOrEmpty(txtfind_data2.Text.Trim()))
            {

                if (!String.IsNullOrEmpty(twhere2.Trim()))
                {
                    twhere2 = twhere2 + " and ";
                }

                twhere2 = twhere2 + " (";

                if (IsValidDate(txtfind_data2.Text))
                {
                    DateTime fromdate = DateTime.Parse(txtfind_data2.Text.Trim());
                    string fromdatetime = fromdate.ToString("yyyy-MM-dd");

                    DateTime todate = DateTime.Parse(txtfind_data2.Text.Trim()).AddDays(1);
                    string todatetime = todate.ToString("yyyy-MM-dd");

                    twhere2 = twhere2 + " (Start_date >= '" + fromdate + "' and Start_date <  '" + todatetime + "'  ) or ";


                }


                twhere2 = twhere2 + " h.idcard like '%" + txtfind_data2.Text.Trim() + "%' ";
                twhere2 = twhere2 + " or ";
                twhere2 = twhere2 + " h.last_name like '%" + txtfind_data2.Text.Trim() + "%' ";
                twhere2 = twhere2 + " or ";
                twhere2 = twhere2 + " c.code like '%" + txtfind_data2.Text.Trim() + "%' ";
                twhere2 = twhere2 + " or ";
                twhere2 = twhere2 + " h.office_phone like '%" + txtfind_data2.Text.Trim() + "%' ";
                twhere2 = twhere2 + " or ";
                twhere2 = twhere2 + " h.passport like '%" + txtfind_data2.Text.Trim() + "%' ";
                twhere2 = twhere2 + " ) ";

            }
            if (!String.IsNullOrEmpty(txtfind_data2.Text.Trim()) || !String.IsNullOrEmpty(comvalid.Text.Trim()))
            {

                twhere2 = " and " + twhere2;

            }
            string query = "SELECT c.Code , h.*  FROM CRDHLD h , Card c where  h.id = c.Owner and (h.company = 'FTC' OR h.company = 'DDW' OR h.company = 'ADC' OR h.company = 'PH3 Entrance' OR h.company = 'Phase3 Entrance' OR h.company = 'PH3' OR h.company = 'SFC') " + twhere2 + " order BY h.BorrowCardDate desc, h.Start_date ";
            Console.WriteLine(query);

            using (OdbcConnection connection = new OdbcConnection(ConnectionString))
            {
                connection.Open();

                OdbcCommand command = new OdbcCommand(query, connection);


                OdbcDataAdapter DA = new OdbcDataAdapter(command);
                not_return_card = new DataSet();

                not_return_card.Clear();
                DA.Fill(not_return_card, "not_return_card");
                g_not_return_card.AutoGenerateColumns = false;
                g_not_return_card.DataSource = not_return_card;
                g_not_return_card.DataMember = "not_return_card";
                string c_datetime = get_server_time_conn(ConnectionString);
                DateTime NowDate = DateTime.ParseExact(c_datetime, "yyyy-MM-dd HH:mm:ss", null);

                foreach (DataGridViewRow dr in g_not_return_card.Rows)
                {
                    String tstatus = "";
                    Boolean status_pass = true;
                    if (dr.Cells["valid"].Value.ToString().Trim().ToUpper().Equals("1"))
                    {
                        tstatus = "有效";
                    }


                    string c_to_date = dr.Cells["to_date"].Value.ToString().Trim().ToUpper();
                    if (string.IsNullOrEmpty(c_to_date) == true)
                    {
                        c_to_date = "2999-12-31 23:59:59";

                    }

                    DateTime d_to_date = DateTime.ParseExact(c_to_date, "yyyy-MM-dd HH:mm:ss", null);

                    if (d_to_date <= NowDate)
                    {

                        tstatus = "過期";
                        status_pass = false;
                    }
                    if (dr.Cells["valid"].Value.ToString().Trim().ToUpper().Equals("0"))
                    {

                        tstatus = "停用";
                        status_pass = false;

                    }

                    dr.Cells["status"].Value = tstatus;

                }

                //       if (!string.IsNullOrEmpty(dr.Cells["code"].Value.ToString().Trim().ToUpper()))
                //        {
                //            if (!string.IsNullOrEmpty(tstatus))
                //           {
                //              tstatus = tstatus + " 和 ";

                //          }
                //         tstatus = "已有卡號:" + dr.Cells["code"].Value.ToString().Trim().ToUpper();
                //         status_pass = false;



                command.Dispose();
                DA.Dispose();
                connection.Dispose();

            }




        }
        private void button15_Click_1(object sender, EventArgs e)
        {
            find_not_return_card();
        }

        private void tabPage7_Click(object sender, EventArgs e)
        {

        }

        private void g_not_return_card_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void pic_c_photo_Click(object sender, EventArgs e)
        {

        }
        private void stop_card()
        {

            string num = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["num"].Value.ToString().Trim();
            string name = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["last_name"].Value.ToString().Trim();
            string staff_no = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["first_name"].Value.ToString().Trim();
            string card_no = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["code"].Value.ToString().Trim();
            string dept = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["company"].Value.ToString().Trim();
            string passport_no = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["Passport"].Value.ToString().Trim();
            string passport_date = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["Passportdate"].Value.ToString().Trim();
            string other = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["other"].Value.ToString().Trim();
            string idcard = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["idcard"].Value.ToString().Trim();
            string phone = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["Office_phone"].Value.ToString().Trim();




            string c_datetime = get_server_time_conn(ConnectionString2);
            string card_reader = card_reader_name.Trim();

            string staff_date = c_datetime;


            // add log
            using (OdbcConnection conn2 = new OdbcConnection(ConnectionString2))
            {



                string query = "";




                query = "INSERT INTO dr_card_log ([event_dt],[borrow_return],[card_reader],[idcard],[name],[staff_no],[staff_date],[phone],[passport_no],[passport_date],[other],[dept],[card_no])";

                query = query + " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)";

                //
                //

                OdbcCommand command = new OdbcCommand(query, conn2);





                command.Parameters.AddWithValue("event_dt", c_datetime);
                command.Parameters.AddWithValue("borrow_return", "StopCard");

                command.Parameters.AddWithValue("card_reader", card_reader);
                command.Parameters.AddWithValue("idcard", idcard);


                command.Parameters.AddWithValue("name", name);

                command.Parameters.AddWithValue("staff_no", staff_no);

                command.Parameters.AddWithValue("staff_date", staff_date.ToString().Trim());

                command.Parameters.AddWithValue("phone", phone);


                command.Parameters.AddWithValue("passport_no", passport_no.ToString().Trim());

                command.Parameters.AddWithValue("passport_date", passport_date);
                command.Parameters.AddWithValue("other", other);
                command.Parameters.AddWithValue("dept", dept);
                command.Parameters.AddWithValue("card_no", card_no);
                conn2.Open();
                int result = command.ExecuteNonQuery();

                if (result < 0)
                {
                    Console.WriteLine("加不到借卡記錄, 身份證為:" + idcard.ToString() + " !");
                    MessageBox.Show("錯誤! 加不到借卡記錄, 身份證為: " + idcard.ToString() + " !");
                }
                else
                {
                    Console.WriteLine("加不到借卡記錄, 身份證為:" + idcard.ToString() + " !");

                };
                //
                command.Dispose();
                conn2.Dispose();

            }
            // add log end

            using (OdbcConnection conn1 = new OdbcConnection(connection1))
            {


                conn1.Open();

                string query5 = "INSERT INTO QueueMSGAPI ([DateCreated], [ServerName], [Cmd], [Msg], [Status], [Result]) VALUES(?,?,?,?,?,?)";



                OdbcCommand command5 = new OdbcCommand(query5, conn1);




                String tmsg = "<query>";
                tmsg = tmsg + "<Number>" + num.ToString() + "</Number>";
                //     tmsg = tmsg + "<From_Date>2000-01-01 00:00:00</From_Date>";

                //        tmsg = tmsg + "<To_Date>2020-01-01 00:00:01</To_Date>";
                //tmsg = tmsg + "<Valid>0</Valid>";
                tmsg = tmsg + "<Validated>0</Validated>";
                //  tmsg = tmsg + "<BorrowCardDate>.</BorrowCardDate>";

                tmsg = tmsg + "</query>";


                command5.Parameters.AddWithValue("DateCreated ", c_datetime);
                command5.Parameters.AddWithValue("ServerName", servername);
                command5.Parameters.AddWithValue("Cmd", "ImportOneCardHolderXML");
                command5.Parameters.AddWithValue("Msg", tmsg);
                command5.Parameters.AddWithValue("Status", 0);

                command5.Parameters.AddWithValue("Result", 0);



                int result5 = command5.ExecuteNonQuery();
                // Check Error

                if (result5 < 0)
                {
                    Console.WriteLine("error insert id:" + num.ToString() + " data into DDS!");



                }
                else
                {
                    Console.WriteLine("insert id:" + num.ToString() + " data into DDS!");
                    MessageBox.Show("系統編號: " + num.ToString() + "已停證, 名稱: " + staff_no.Trim() + " " + name.Trim() + " 卡號:" + card_no);
                    clean_all();

                };
                command5.Dispose();
                conn1.Close();
                conn1.Dispose();

            }











        }

        private void activie_card()
        {

            string num = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["num"].Value.ToString().Trim();
            string name = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["last_name"].Value.ToString().Trim();
            string staff_no = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["first_name"].Value.ToString().Trim();
            string card_no = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["code"].Value.ToString().Trim();
            string dept = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["company"].Value.ToString().Trim();
            string passport_no = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["Passport"].Value.ToString().Trim();
            string passport_date = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["Passportdate"].Value.ToString().Trim();
            string other = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["other"].Value.ToString().Trim();
            string idcard = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["idcard"].Value.ToString().Trim();
            string phone = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["Office_phone"].Value.ToString().Trim();




            string c_datetime = get_server_time_conn(ConnectionString2);
            string card_reader = card_reader_name.Trim();

            string staff_date = c_datetime;


            // add log
            using (OdbcConnection conn2 = new OdbcConnection(ConnectionString2))
            {



                string query = "";




                query = "INSERT INTO dr_card_log ([event_dt],[borrow_return],[card_reader],[idcard],[name],[staff_no],[staff_date],[phone],[passport_no],[passport_date],[other],[dept],[card_no])";

                query = query + " VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?)";

                //
                //

                OdbcCommand command = new OdbcCommand(query, conn2);





                command.Parameters.AddWithValue("event_dt", c_datetime);
                command.Parameters.AddWithValue("borrow_return", "ActCard");

                command.Parameters.AddWithValue("card_reader", card_reader);
                command.Parameters.AddWithValue("idcard", idcard);


                command.Parameters.AddWithValue("name", name);

                command.Parameters.AddWithValue("staff_no", staff_no);

                command.Parameters.AddWithValue("staff_date", staff_date.ToString().Trim());

                command.Parameters.AddWithValue("phone", phone);


                command.Parameters.AddWithValue("passport_no", passport_no.ToString().Trim());

                command.Parameters.AddWithValue("passport_date", passport_date);
                command.Parameters.AddWithValue("other", other);
                command.Parameters.AddWithValue("dept", dept);
                command.Parameters.AddWithValue("card_no", card_no);
                conn2.Open();
                int result = command.ExecuteNonQuery();

                if (result < 0)
                {
                    Console.WriteLine("加不到借卡記錄, 身份證為:" + idcard.ToString() + " !");
                    MessageBox.Show("錯誤! 加不到借卡記錄, 身份證為: " + idcard.ToString() + " !");
                }
                else
                {
                    Console.WriteLine("加不到借卡記錄, 身份證為:" + idcard.ToString() + " !");

                };
                //
                command.Dispose();
                conn2.Dispose();
            }
            // add log end

            using (OdbcConnection conn1 = new OdbcConnection(connection1))
            {



                conn1.Open();

                string query5 = "INSERT INTO QueueMSGAPI ([DateCreated], [ServerName], [Cmd], [Msg], [Status], [Result]) VALUES(?,?,?,?,?,?)";



                OdbcCommand command5 = new OdbcCommand(query5, conn1);




                String tmsg = "<query>";
                tmsg = tmsg + "<Number>" + num.ToString() + "</Number>";
                //     tmsg = tmsg + "<From_Date>2000-01-01 00:00:00</From_Date>";

                //        tmsg = tmsg + "<To_Date>2020-01-01 00:00:01</To_Date>";
                //tmsg = tmsg + "<Valid>0</Valid>";
                tmsg = tmsg + "<Validated>1</Validated>";
                //  tmsg = tmsg + "<BorrowCardDate>.</BorrowCardDate>";

                tmsg = tmsg + "</query>";


                command5.Parameters.AddWithValue("DateCreated ", c_datetime);
                command5.Parameters.AddWithValue("ServerName", servername);
                command5.Parameters.AddWithValue("Cmd", "ImportOneCardHolderXML");
                command5.Parameters.AddWithValue("Msg", tmsg);
                command5.Parameters.AddWithValue("Status", 0);

                command5.Parameters.AddWithValue("Result", 0);



                int result5 = command5.ExecuteNonQuery();
                // Check Error

                if (result5 < 0)
                {
                    Console.WriteLine("error insert id:" + num.ToString() + " data into DDS!");



                }
                else
                {
                    Console.WriteLine("insert id:" + num.ToString() + " data into DDS!");
                    MessageBox.Show("系統編號: " + num.ToString() + "已生效證, 名稱: " + staff_no.Trim() + " " + name.Trim() + " 卡號:" + card_no);
                    clean_all();

                };
                command5.Dispose();
                conn1.Close();
                conn1.Dispose();



            }










        }
        private void button16_Click(object sender, EventArgs e)
        {


            string num = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["num"].Value.ToString().Trim();
            string name = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["last_name"].Value.ToString().Trim();
            string staff_id = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["first_name"].Value.ToString().Trim();
            string idcard = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["idcard"].Value.ToString().Trim();

            DialogResult dialogResult = MessageBox.Show("確定停用: " + (staff_id + name).Trim() + " 身份證: " + idcard, "停用", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (dialogResult == DialogResult.Yes)
            {
                stop_card();
            }
            else if (dialogResult == DialogResult.No)
            {
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            string num = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["num"].Value.ToString().Trim();
            string name = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["last_name"].Value.ToString().Trim();
            string staff_id = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["first_name"].Value.ToString().Trim();
            string idcard = g_not_return_card.Rows[g_not_return_card.CurrentRow.Index].Cells["idcard"].Value.ToString().Trim();

            DialogResult dialogResult = MessageBox.Show("確定生效: " + (staff_id + name).Trim() + " 身份證: " + idcard, "生效", MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button2);
            if (dialogResult == DialogResult.Yes)
            {
                activie_card();
            }
            else if (dialogResult == DialogResult.No)
            {
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
        }

        private void button19_Click(object sender, EventArgs e)
        {
            cam cam = new cam();

            var result = cam.ShowDialog();
            if (result == DialogResult.OK)
            {
                string val = cam.ruturn_value;           //values preserved after close
                curr_photo_path = pic_c_photo.ImageLocation;
                get_photopath = val;


                pic_c_photo.ImageLocation = get_photopath;
                is_update_photo = true;
                labphoto_status.Text = "新相";
                labphoto_status.ForeColor = Color.FromArgb(255, 0, 0);
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (is_update_photo == true)
            {
                pic_c_photo.ImageLocation = curr_photo_path;

            }
            is_update_photo = false;
            labphoto_status.Text = "現在";
            labphoto_status.ForeColor = Color.FromArgb(0, 0, 0);
        }

        private void button18_Click_1(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog1 = new OpenFileDialog();

            openFileDialog1.InitialDirectory = "c:\\";
            openFileDialog1.Filter = "jpg files (*.jpg)|*.jpg";
            openFileDialog1.FilterIndex = 0;
            openFileDialog1.RestoreDirectory = true;

            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {

                get_photopath = openFileDialog1.FileName;
                pic_c_photo.ImageLocation = get_photopath;
                is_update_photo = true;
                labphoto_status.Text = "新相";
                labphoto_status.ForeColor = Color.FromArgb(255, 0, 0);


            }

        }

        private void button21_Click(object sender, EventArgs e)
        {
            // put_new_num(155898);
            int aa = 0;
            int id = 123;
            Console.WriteLine("a: " + aa.ToString());
        }

        private void txtl_num_TextChanged(object sender, EventArgs e)
        {

        }

        private void timer_free_Tick(object sender, EventArgs e)
        {
            Console.WriteLine("Free");
            GC.Collect();
            GC.WaitForPendingFinalizers();
            //      MessageBox.Show("Free");

        }
    }
}
