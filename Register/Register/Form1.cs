using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Globalization;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Threading;

namespace Register
{
    public partial class Form1 : Form
    {
        // SQLite物件
        private SQLiteHelper sqlHeper;
        private List<string[]> class_info_list;
        private List<string[]> find_kidinfo_list;
        private List<string[]> class_table_list;
        
        public Form1()
        {
            InitializeComponent();
            // new SQLiteHelper物件
            this.sqlHeper = new SQLiteHelper();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.Text = "課程報名系統     開啟時間 : [" + DateTime.Now.ToString()+"]";
            DateTime myDate = DateTime.Now;
            numericUpDown1.Value = myDate.Month;
            numericUpDown2.Value = myDate.Month;
            Boolean OK = this.sqlHeper.GetClassInfo(out class_info_list);
            if (OK)
            {
                foreach (String[] s in class_info_list)
                {
                    comboBox2.Items.Add(s[0]);
                    //MessageBox.Show(s[0]+" "+s[1]);
                }
            }
            // 預設選擇課程為 0
            if (comboBox2.Items.Count >= 1) comboBox2.SelectedIndex = 0;
            if (comboBox4.Items.Count >= 1) comboBox4.SelectedIndex = 0;

            GetClass(numericUpDown1); //取得課程列表
            GetClass(numericUpDown2);
        }

        private void TabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            Search_list();
            Search_kids(textBox2, Selectclass_list);
            Search_kids(textBox3, Searchclass_list);
            GetClass(numericUpDown1); //取得課程列表
            GetClass(numericUpDown2);
            dataGridView2.Rows.Clear();
            dataGridView3.Rows.Clear();
            textBox4.Text = "";
            comboBox2.SelectedIndex = 0;
            comboBox4.SelectedIndex = 0;
            radioButton7.Checked = true;
            radioButton_male.Checked = true;
            numericUpDown3.Text = "5";
            textBox1.Text = "";
            textBox2.Text = "";
            textBox3.Text = "";

            ///選擇課程頁面
            label2.Text = "幼兒姓名: ";
            label17.Text = "幼兒性別: ";
            label22.Text = "幼兒年齡: ";
            label23.Text = "家長姓名:  ";

            ///選課查詢頁面
            label4.Text = "幼兒姓名: ";
            label5.Text = "幼兒性別: ";
            label11.Text = "幼兒年齡: ";
            label13.Text = "家長姓名: ";
        }

        #region 課程新增頁面
        private void Add_calss_button_Click(object sender, EventArgs e)
        {
            if (dateTimePicker2.Value.Year != DateTime.Now.Year)
            {
                int years = Convert.ToInt32(DateTime.Now.Year.ToString());
                MessageBox.Show("請選擇今年的年份!\n(" + years + "年)", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                years -= years;
                dateTimePicker2.Value = DateTime.Now.AddYears(years);
                return;
            }
            try
            {
                int people = Convert.ToInt32(numericUpDown3.Text.ToString());
                if (people > 100 || people < 1)
                {
                    MessageBox.Show("開課人數範圍 1 ~ 100 人!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    numericUpDown3.Text = "5";
                    return;
                }
            }
            catch (Exception)
            {
                MessageBox.Show("數字格式錯誤\n(請使用半形)!", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                textBox4.Text = "";
                numericUpDown3.Text = "5";
                return;
            }
            try
            {            
                String class_name = "";
                int class_type = Int32.Parse(class_info_list[comboBox2.SelectedIndex][1]);
                String date = String.Format("{0:yyyy-MM-dd}", dateTimePicker2.Value.Date) + " " + comboBox4.Text + ":00";
                String class_datetime = date;
                try
                {
                    DateTime result;
                    CultureInfo provider = CultureInfo.InvariantCulture;
                    result = DateTime.ParseExact(class_datetime, "yyyy-MM-dd HH:mm:00", provider);
                }
                catch (FormatException)
                {
                    MessageBox.Show("時間格式錯誤\n\n請依以下此格式(24時制)\n\n[時:分]", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    comboBox4.SelectedIndex = 0;
                    return;
                }

                int age_range = 0;

                if (radioButton6.Checked)
                {
                    age_range = 0;
                }
                else if (radioButton7.Checked)
                {
                    age_range = 1;
                }
                else if (radioButton8.Checked)
                {
                    age_range = 2;
                }
                else if (radioButton9.Checked)
                {
                    age_range = 3;
                }
                else if (radioButton10.Checked)
                {
                    age_range = 4;
                }
                else if (radioButton1.Checked)
                {
                    age_range = 5;
                }
                else
                {
                    MessageBox.Show("請選擇開課年齡!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                int class_age_range = age_range;
                int kid_max = Int32.Parse(numericUpDown3.Text);

                Boolean OK = this.sqlHeper.Add_Class_List(class_name, class_type, class_datetime, class_age_range, kid_max);
                if (OK) MessageBox.Show("新增課程成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else MessageBox.Show("課程重複，新增失敗!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            }
            catch (Exception)
            {
                MessageBox.Show("請在 [額外新增課程] 輸入", "錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
                comboBox2.SelectedIndex = 0;
            }
        }

        private void Add_extra_class_btn_Click(object sender, EventArgs e)
        {
            Boolean OK0 = JudgmentClass();
            if (OK0)
            {
                MessageBox.Show("課程重複，新增失敗!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Exclamation); 
                textBox4.Text = "";
                return;
            }
            Boolean OK1 = sqlHeper.Classadd(textBox4.Text);
            Boolean OK2;
            if (textBox4.Text == "") { MessageBox.Show("請輸入課程名稱!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); return; }
            if (OK1)
            {
                comboBox2.Items.Clear();
                OK2 = this.sqlHeper.GetClassInfo(out class_info_list);
                if (OK2)
                {
                    foreach (String[] s in class_info_list)
                    {
                        comboBox2.Items.Add(s[0]);
                    }
                    MessageBox.Show("已新增額外課程!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }

        }

        private void Backup_btn_Click(object sender, EventArgs e)
        {
            sqlHeper.BackupDB();
        }

        private void Restore_btn_Click(object sender, EventArgs e)
        {
            sqlHeper.RestoreDB();
            dataGridView3.Rows.Clear();
        }

        private void Del_Alldata_Click(object sender, EventArgs e)
        { 
            try
            {
                DialogResult myResult = MessageBox.Show("確認後會將全部的幼兒資料和選課課程刪除。\n(預設課程會保留)\n\n★建議先備份資料庫!\n\n是否刪除全部資料?", "警告!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                if (myResult == DialogResult.Yes)
                {
                    comboBox2.Items.Clear();                  
                    Boolean OK1 = sqlHeper.Del_Alldata();
                    if (OK1)
                    {
                        Boolean OK2 = this.sqlHeper.GetClassInfo(out class_info_list);
                        if (OK2)
                        {
                            foreach (String[] s in class_info_list)
                            {                        
                                comboBox2.Items.Add(s[0]);
                            }
                            MessageBox.Show("全部資料已刪除!", "警告!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        }
                    }
                }
            }
            catch (Exception exp)
            {
                exp.Message.ToString();
            }
            finally
            {
                dataGridView1.Rows.Clear();
                Selectclass_list.Items.Clear();
                checkedListBox26.Items.Clear();
                Class_list_lb.Items.Clear();
                Searchclass_list.Items.Clear();
                dataGridView3.Rows.Clear();
                comboBox2.SelectedIndex = 0;
                textBox4.Text = "";
            }
        }

        private void Del_AllClass_btn_Click(object sender, EventArgs e)
        {
            try
            {
                DialogResult myResult = MessageBox.Show("確認後會將新增的課程名單全部刪除。\n(幼兒資料和額外課程會保留)\n\n★建議先備份資料庫!\n\n是否刪除全部課程?", "警告!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);

                if (myResult == DialogResult.Yes)
                {
                    bool dd = sqlHeper.Del_AllClass();
                    if (dd) MessageBox.Show("全部課程已刪除!", "警告!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
            catch (Exception exp)
            {
                exp.Message.ToString();
            }
            finally
            {
                dataGridView1.Rows.Clear();
                Selectclass_list.Items.Clear();
                checkedListBox26.Items.Clear();
                Class_list_lb.Items.Clear();
                Searchclass_list.Items.Clear();
                dataGridView3.Rows.Clear();
            }
        }
        #endregion

        #region 個人建檔頁面
        private void Bt1_add_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(txt2_bname.Text))
                {
                    MessageBox.Show("請輸入幼兒姓名!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                String KidName = txt2_bname.Text.ToString();
                String ParentsName = txt1_Aname.Text.ToString();
                String ParentsPhone = txt3_phone.Text.ToString();
                Boolean sex = true;
                String birthday = dateTimePicker1.Text.ToString();

                Age kid_age = CalculateAge(dateTimePicker1.Value, DateTime.Now);
                float agekid = kid_age.Years + (kid_age.Months * 0.1f);

                String bd = dateTimePicker1.Value.ToString("yyyy-MM-dd HH:mm:ss");
                if (radioButton_female.Checked) sex = false;
                #region 註解
                /**
                //float agekid = DateAndTime.DateDiff(DateInterval.Year,//計算年齡
                //dateTimePicker1.Value, DateTime.Now, FirstDayOfWeek.Sunday, FirstWeekOfYear.Jan1);

                //float agekid = DateAndTime.DateDiff(DateInterval.Day,//計算年齡
                //dateTimePicker1.Value, DateTime.Now, FirstDayOfWeek.Sunday, FirstWeekOfYear.Jan1);

                // y = DateAndTime.DateDiff(DateInterval.Month,//計算年齡
                //     dateTimePicker1.Value, DateTime.Now, FirstDayOfWeek.Sunday, FirstWeekOfYear.Jan1);
                // y = y / 12;
                //Double m = ((y / 12) * 0.1);
                // y = y + m;
                //MessageBox.Show(agekid.ToString() + " "+ y.ToString() + " " + m.ToString() );
                //MessageBox.Show(getDateDiff(DateTime.Now, dateTimePicker1.Value).ToString());  **/
                #endregion
                if (!(radioButton_male.Checked || radioButton_female.Checked))
                {
                    MessageBox.Show("請選擇性別!");
                    return;
                }

                Boolean OK = this.sqlHeper.Add_KidInfo(KidName, ParentsName, ParentsPhone, sex, birthday, bd);
                if (OK) { } //MessageBox.Show("新增成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                else MessageBox.Show("新增失敗，原因:已重複", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);

                #region 新增完重新整理
                Search_list();
                #endregion
            }
            catch (Exception exp) { exp.Message.ToString(); }
        }

        private void Del_kidinfo_bt_Click(object sender, EventArgs e)
        {
            try
            {
                string number = dataGridView1.CurrentRow.Cells[0].Value.ToString();//取編號
                string name = dataGridView1.CurrentRow.Cells[1].Value.ToString();//取姓名
                DialogResult myResult = MessageBox.Show("是否刪除編號 [" + number + "] " + name + " 的資料", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (myResult == DialogResult.Yes)
                {
                    bool d = sqlHeper.Del_Kidinfo(number);
                    if (d) MessageBox.Show("刪除成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception exp)
            {
                exp.Message.ToString();
            }
            finally
            {
                #region 刪除完重新整理
                Search_list();
                dataGridView2.Rows.Clear();
                #endregion
            }
        }

        private void Bt2_查詢_Click(object sender, EventArgs e)
        {
            Search_list();
        }

        #endregion

        #region 選擇課程頁面
        private void Select_class_search_btn_Click(object sender, EventArgs e)
        {
            Search_kids(textBox2, Selectclass_list);
        }

        private void Select_class_add_btn_Click(object sender, EventArgs e)
        {
            if (checkedListBox26.CheckedItems.Count == 0)
            {
                MessageBox.Show("請選擇課程!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            if (Selectclass_list.SelectedItems.Count == 0)
            {
                MessageBox.Show("請選擇幼兒!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            String[] kid_info = find_kidinfo_list[Selectclass_list.SelectedIndex];

            if (checkedListBox26.CheckedItems.Count != 0)
            {
                int count_ok = 0;
                string s = "";
                for (int i = 0; i <= checkedListBox26.Items.Count - 1; i++)
                {
                    String[] class_table = class_table_list[i];
                    CheckState cs = checkedListBox26.GetItemCheckState(i);
                    s += "Checked Item " + (i + 1).ToString() + " = " + checkedListBox26.Items[i].ToString() + "\n";
                    if (cs == CheckState.Checked)
                    {
                        String class_type_id = class_table[5];
                        String kid_nid = kid_info[6];
                        String class_name = class_table[1];
                        String class_datetime = class_table[2];
                        DateTime dt = Convert.ToDateTime(class_datetime);
                        Boolean b = this.sqlHeper.Set_Selected_class(class_type_id, kid_nid);
                        if (b)
                        {
                            count_ok += 1;
                        }
                        else
                        {
                            MessageBox.Show("新增失敗，原因: 課程重複!" + "\n\n" + "課程 [" + class_name + "]" + "\n" + "日期 [" + dt.ToString("yyyy-MM-dd") + "]" + "\n" + "時間 [" + dt.ToString("HH:mm") + "]", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                    }
                }
                if (count_ok >= 1)
                {
                    MessageBox.Show("新增完成", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GetClass(numericUpDown1); //取得課程列表
                    GetClass(numericUpDown2);
                    dataGridView2.Rows.Clear();
                    dataGridView3.Rows.Clear();
                }
            }
        }

        private void Select_class_updatclass_btn_Click(object sender, EventArgs e)
        {
            GetClass(numericUpDown1); //取得課程列表
            GetClass(numericUpDown2);
        }

        private void NumericUpDown1_ValueChanged(object sender, EventArgs e)
        {
            GetClass(numericUpDown1);
            numericUpDown2.Value = numericUpDown1.Value;
            dataGridView2.Rows.Clear();
        }

        private void Selectclass_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            Kids_Update_List(Selectclass_list, label2, label17, label22, label23);
        }

        private void CheckedListBox26_MouseUp(object sender, MouseEventArgs e)
        {
            JudgmentPeople();
        }

        private void CheckedListBox26_DoubleClick(object sender, EventArgs e)
        {
            JudgmentPeople();
        }

        private void Del_SelectClass_btn_Click(object sender, EventArgs e)
        {
            if (checkedListBox26.CheckedItems.Count == 0)
            {
                MessageBox.Show("請選擇要刪除的課程!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DialogResult myResult = MessageBox.Show("是否刪除所選課程?", "警告!", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
            try
            {
                if (myResult == DialogResult.Yes)
                {
                    string s = "";
                    for (int i = 0; i <= checkedListBox26.Items.Count - 1; i++)
                    {
                        String[] class_table = class_table_list[i];
                        CheckState cs = checkedListBox26.GetItemCheckState(i);
                        s += "Checked Item " + (i + 1).ToString() + " = " + checkedListBox26.Items[i].ToString() + "\n";
                        if (cs == CheckState.Checked)
                        {
                            //MessageBox.Show(checkedListBox26.Items[i].ToString());
                            String class_type_id = class_table[5];
                            sqlHeper.Del_Selected_class(class_type_id);
                        }
                    }

                    MessageBox.Show("刪除完成", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    GetClass(numericUpDown1); //取得課程列表
                    GetClass(numericUpDown2);
                }
            }
            catch (Exception) { MessageBox.Show("操作錯誤"); return; }
        }
        #endregion

        #region 課程列表頁面
        private void Classlist_updatelist_btn_Click(object sender, EventArgs e)
        {
            GetClass(numericUpDown1); //取得課程列表
            GetClass(numericUpDown2);
            dataGridView2.Rows.Clear();
        }

        private void NumericUpDown2_ValueChanged(object sender, EventArgs e)
        {
            GetClass(numericUpDown2);
            numericUpDown1.Value = numericUpDown2.Value;
            dataGridView2.Rows.Clear();
        }

        private void Classlist_searchkids_btn_Click(object sender, EventArgs e)
        {
            Class_Student_List();
        }

        private void Class_list_lb_KeyUp(object sender, KeyEventArgs e)
        {
            Class_Student_List();
        }

        private void Class_list_lb_Click(object sender, EventArgs e)
        {
            Class_Student_List();
        }

        private void SavePDF_btn_Click(object sender, EventArgs e)
        {
            SaveDataGridViewToPDF(dataGridView2, Class_list_lb, "點名表");
        }

        #endregion

        #region 選課查詢頁面
        private void Selectsearch_search_btn_Click(object sender, EventArgs e)
        {
            Search_kids(textBox3, Searchclass_list);
        }

        private void Selectsearch_class_btn_Click(object sender, EventArgs e)
        {
            Select_class_search();
        }

        private void Del_class_btn_Click(object sender, EventArgs e)
        {
            try
            {
                string class_nid = dataGridView3.CurrentRow.Cells[0].Value.ToString(); //取編號
                DialogResult myResult = MessageBox.Show("是否刪除 編號 [" + class_nid + "] 課程", "警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (myResult == DialogResult.Yes)
                {
                    bool d = sqlHeper.Del_Class(class_nid);
                    if (d) MessageBox.Show("刪除成功", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    Select_class_search();
                    GetClass(numericUpDown1); //取得課程列表
                    GetClass(numericUpDown2);
                    dataGridView2.Rows.Clear();
                }
            }
            catch (NullReferenceException) { MessageBox.Show("請選擇要刪除的課程", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information); }
        }

        private void Searchclass_list_SelectedIndexChanged(object sender, EventArgs e)
        {
            Kids_Update_List(Searchclass_list, label4, label5, label11, label13);
        }

        private void Searchclass_list_Click(object sender, EventArgs e)
        {
            Select_class_search();
        }

        private void Searchclass_list_KeyUp(object sender, KeyEventArgs e)
        {
            Select_class_search();
        }
        #endregion

        /// <summary>
        /// 以下是Function
        /// </summary>

        #region 取得幼兒資訊列表
        public void Search_list()
        {
            DataGridViewRowCollection rows = dataGridView1.Rows;
            dataGridView1.Columns[6].DefaultCellStyle.Format = "yyyy-MM-dd";
            rows.Clear();
            String find_kid_Name = textBox1.Text.ToString();
            Boolean OK = this.sqlHeper.FindKid_Info(find_kid_Name, out find_kidinfo_list);
            if (OK)
            {
                foreach (String[] kid_info in find_kidinfo_list)
                {
                    //foreach (String kid_data in kid_info) {
                    DateTime dt = Convert.ToDateTime(kid_info[4]);
                    Age kid_age = CalculateAge(dt, DateTime.Now);
                    float agekid = kid_age.Years + (kid_age.Months * 0.1f);
                    //MessageBox.Show(kid_info[5]);
                    rows.Add(new Object[] { kid_info[6], kid_info[0], kid_info[3], dt.ToString("yyyy-MM-dd"), agekid, kid_info[1], kid_info[2],kid_info[7], kid_info[5] });
                }
                //string[] data = list[0];
                //MessageBox.Show( data[0] );
            }
        }
        #endregion

        #region 取得課程列表
        /**
         * 取得課程列表 **/
        private void GetClass(NumericUpDown NumUpDown)
        {
            try
            {
                DateTime myDate = new DateTime(DateTime.Now.Year, (int)NumUpDown.Value, DateTime.Now.Day);
                //DateTime myDate = DateTime.Now;
                //myDate.AddMonths(myDate.Month - (int)numericUpDown1.Value);
                string now_date = myDate.ToString("yyyy-MM");
                checkedListBox26.Items.Clear();
                Class_list_lb.Items.Clear();
                Boolean OK = this.sqlHeper.GetClassList(now_date, out class_table_list);
                if (OK)
                {
                    foreach (String[] class_table in class_table_list)
                    {
                        String class_type = class_table[5]; //class_table[1]
                        String class_datetime = class_table[2];
                        String class_age_range = "無限制";       //class_table[3];
                        String kid_max = class_table[4];
                        String class_name = class_table[1]; //課程名稱
                        String kid_select = class_table[6];

                        if (kid_select == "") { kid_select = "0"; }
                        int kidmax_int = int.Parse(kid_max, NumberStyles.AllowDecimalPoint);
                        int kidselect_int = int.Parse(kid_select, NumberStyles.AllowDecimalPoint);
                        int remaining = kidmax_int - kidselect_int;

                        switch (class_table[3])
                        {
                            case "0":
                                class_age_range = "無限制";
                                break;

                            case "1":
                                class_age_range = "1~3歲";
                                break;

                            case "2":
                                class_age_range = "2~4歲";
                                break;

                            case "3":
                                class_age_range = "2~6歲";
                                break;

                            case "4":
                                class_age_range = "3~6歲";
                                break;

                            case "5":
                                class_age_range = "4~6歲";
                                break;
                        }
                        DateTime dt = Convert.ToDateTime(class_datetime);
                        String week = "";
                        switch (dt.DayOfWeek)
                        {
                            case DayOfWeek.Monday:
                                week = "(一)";
                                break;
                            case DayOfWeek.Tuesday:
                                week = "(二)";
                                break;
                            case DayOfWeek.Wednesday:
                                week = "(三)";
                                break;
                            case DayOfWeek.Thursday:
                                week = "(四)";
                                break;
                            case DayOfWeek.Friday:
                                week = "(五)";
                                break;
                            case DayOfWeek.Saturday:
                                week = "(六)";
                                break;
                            case DayOfWeek.Sunday:
                                week = "(日)";
                                break;
                        }

                        Class_list_lb.Items.Add(dt.ToString("yyyy-MM-dd") +
                            week +
                            dt.ToString(" [開課時間] HH：mm") +
                            "  [名稱]" + class_name +
                            "  [適合年齡]" + class_age_range +
                            "  [總名額]" + kid_max + " [選課人數]" + kid_select + "  [剩餘名額]" + remaining.ToString());

                        //選擇課程列表格式
                        checkedListBox26.Items.Add(dt.ToString("yyyy-MM-dd") +
                            week +
                            dt.ToString(" [開課時間] HH：mm") +
                            "  [名稱]:" + class_name +
                            "  [適合年齡]" + class_age_range +
                            "  [總名額]" + kid_max + "  [選課人數]" + kid_select + "  [剩餘名額]" + remaining.ToString());
                    }
                    JudgmentPeople();
                }
            }

            catch (ArgumentOutOfRangeException)
            {
                MessageBox.Show("超出月份", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                if (NumUpDown.Value == 13) { NumUpDown.Value = 12; }
                else if (NumUpDown.Value == 0) { NumUpDown.Value = 1; }
            }
        }
        #endregion

        #region 查詢幼兒
        private void Search_kids(TextBox Tb, ListBox Lb)
        {
            Lb.Items.Clear();
            String find_kid_Name = Tb.Text.ToString();
            Boolean OK = this.sqlHeper.FindKid_Info(find_kid_Name, out find_kidinfo_list);
            if (OK)
            {
                foreach (String[] kid_info in find_kidinfo_list)
                {
                    Lb.Items.Add(kid_info[0]);
                }
            }
        }
        #endregion

        #region 點選幼兒更新列表
        private void Kids_Update_List(ListBox Lb, Label kidname_label, Label sex_label, Label agekid_label, Label parentsname_label)
        {
            if (!(find_kidinfo_list != null) || !(find_kidinfo_list.Count >= 1)) return;
            if (Lb.SelectedIndex == -1) return;
            String[] kid_info = find_kidinfo_list[Lb.SelectedIndex];
            String kidname = kid_info[0];
            String parentsname = kid_info[1];
            String sex = kid_info[3];
            String age = kid_info[4];

            DateTime dt = Convert.ToDateTime(age);
            Age kid_age = CalculateAge(dt, DateTime.Now);
            float agekid = kid_age.Years + (kid_age.Months * 0.1f);

            kidname_label.Text = "幼兒姓名: " + kidname;
            sex_label.Text = "幼兒性別: " + sex;
            agekid_label.Text = "幼兒年齡: " + agekid + " 歲";
            parentsname_label.Text = "家長姓名: " + parentsname;
        }
        #endregion

        #region 取得某課程學生列表
        private void Class_Student_List()
        {
            if (Class_list_lb.SelectedItems.Count == 0)
            {
                MessageBox.Show("請選擇課程!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            String[] class_table = class_table_list[Class_list_lb.SelectedIndex];
            DataGridViewRowCollection rows = dataGridView2.Rows;
            rows.Clear();

            List<String[]> getSelected_kid_Class_List;
            Boolean OK = sqlHeper.GetSelected_kid_Class_List(class_table[5], out getSelected_kid_Class_List);
            if (OK)
            {
                int count = 0;
                foreach (String[] s in getSelected_kid_Class_List)
                {
                    count += 1;
                    /**
                     data[0] = kid_name;
                     data[1] = kid_cell;
                     data[2] = kid_sex;
                     data[3] = kid_birthday;
                     data[4] = kid_age; 
                    **/
                    DateTime dt = Convert.ToDateTime(s[3]);//生日
                    Age kid_age = CalculateAge(dt, DateTime.Now);
                    float agekid = kid_age.Years + (kid_age.Months * 0.1f);
                    rows.Add(count, s[0], s[2], s[3], agekid.ToString(), s[1]);
                }
            }
        }
        #endregion

        #region 計算年齡(小數)
        // 計算歲數
        public struct Age
        {
            public int Years;
            public int Months;
            public int Days;
        }
        public static Age CalculateAge(DateTime birthDate, DateTime endDate)
        {
            Age result;
            if (birthDate.Date > endDate.Date)
            {
                result.Years = -1;
                result.Months = -1;
                result.Days = -1;
                return result;
                //throw new ArgumentException("birthDate cannot be higher then endDate", "birthDate");
            }

            int years = endDate.Year - birthDate.Year;
            int months = 0;
            int days = 0;

            // Check if the last year, was a full year.
            if (endDate < birthDate.AddYears(years) && years != 0)
            {
                years--;
            }

            // Calculate the number of months.
            birthDate = birthDate.AddYears(years);

            if (birthDate.Year == endDate.Year)
            {
                months = endDate.Month - birthDate.Month;
            }
            else
            {
                months = (12 - birthDate.Month) + endDate.Month;
            }

            // Check if last month was a complete month.
            if (endDate < birthDate.AddMonths(months) && months != 0)
            {
                months--;
            }

            // Calculate the number of days.
            birthDate = birthDate.AddMonths(months);

            days = (endDate - birthDate).Days;
            result.Years = years;
            result.Months = months;
            result.Days = days;
            return result;
        }
        #endregion

        #region 選課查詢
        public void Select_class_search()
        {
            try
            {
                if (!(find_kidinfo_list != null) || !(find_kidinfo_list.Count >= 1)) return;
                if (Searchclass_list.SelectedIndex == -1) return;
                String[] kid_info = find_kidinfo_list[Searchclass_list.SelectedIndex];

                List<String[]> getkid_Selected_Class_List;
                DataGridViewRowCollection rows = dataGridView3.Rows;
                rows.Clear();

                Boolean OK = this.sqlHeper.Getkid_Selected_Class_List(kid_info[6], out getkid_Selected_Class_List);
                if (OK)
                {
                    // int count = 0;
                    foreach (String[] kid_selected_class_info in getkid_Selected_Class_List)
                    {
                        //    count += 1;
                        String class_datetime = kid_selected_class_info[0];
                        String class_name = kid_selected_class_info[1];
                        String class_nid = kid_selected_class_info[2];

                        DateTime dt = Convert.ToDateTime(class_datetime);
                        String week = "";
                        switch (dt.DayOfWeek)
                        {
                            case DayOfWeek.Monday:
                                week = "(一)";
                                break;
                            case DayOfWeek.Tuesday:
                                week = "(二)";
                                break;
                            case DayOfWeek.Wednesday:
                                week = "(三)";
                                break;
                            case DayOfWeek.Thursday:
                                week = "(四)";
                                break;
                            case DayOfWeek.Friday:
                                week = "(五)";
                                break;
                            case DayOfWeek.Saturday:
                                week = "(六)";
                                break;
                            case DayOfWeek.Sunday:
                                week = "(日)";
                                break;
                        }

                        rows.Add(class_nid, dt.ToString("yyyy-MM-dd") +
                            "  " + week + "  " +
                            dt.ToString("HH:mm") + "  " +
                            "名稱:" + class_name);
                    }
                }
            }
            catch (Exception exp) { exp.Message.ToString(); }
        }
        #endregion

        #region DataGridView儲存PDF
        private void SaveDataGridViewToPDF(DataGridView Dv, ListBox Lb, String secondcell)
        {
            try
            {
                String fontPath = System.Environment.GetEnvironmentVariable("windir") + @"\Fonts\KAIU.ttf";//標楷體
                String class_allinfo = Lb.SelectedItem.ToString();

                String _fileName = DateTime.Now.ToString("yyyy-MM-dd") + ".pdf";
                String.Format(@"{0}.pdf", DateTime.Now.ToString("yyyy-MM-dd"));

                //開啟pdf檔案，四個20代表左右上下邊界各20單位(有按照順序)
                Document document = new Document(iTextSharp.text.PageSize.A4, 20, 20, 20, 20);
                //新建儲存對話框物件
                SaveFileDialog saveFileDialog1 = new SaveFileDialog();
                String PDFName = Lb.SelectedItem.ToString();//String.Format("{0}", DateTime.Now.ToString("yyyy-MM-dd"));

                saveFileDialog1.Filter = "PDF|*.pdf";
                saveFileDialog1.Title = "選擇要儲存的資料夾";
                saveFileDialog1.FileName = PDFName; // 預設擋案名稱
                saveFileDialog1.DefaultExt = ".pdf"; // 預設擋案附檔名(extension)

                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    //建立PDF擋案
                    PdfWriter writer = PdfWriter.GetInstance(document, new FileStream(saveFileDialog1.FileName, FileMode.Create));
                    document.Open();
                    //設定字型
                    FontFactory.Register(fontPath);
                    iTextSharp.text.Font fontContentText = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, 12f, iTextSharp.text.Font.UNDEFINED, BaseColor.BLACK);
                    iTextSharp.text.Font fontContentText1 = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, 12f, iTextSharp.text.Font.UNDEFINED, BaseColor.BLUE);
                    //新增一個一欄的PdfPTable
                    PdfPTable ptbTitle = new PdfPTable(1);
                    //Title新增一行文字，字型可套用剛剛的設定或是自己設
                    PdfPCell subTitleCell = new PdfPCell(new Phrase(class_allinfo, fontContentText1));
                    //將此行文字靠左
                    subTitleCell.HorizontalAlignment = Element.ALIGN_LEFT;
                    //設定此行文字不要有框線
                    subTitleCell.BorderWidth = 0;
                    //將此行文字放到PdfPTable中
                    ptbTitle.AddCell(subTitleCell);

                    //新增第二個Cell(不用特別New)
                    subTitleCell = new PdfPCell(new Phrase(secondcell, FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, 15f, iTextSharp.text.Font.UNDEFINED, BaseColor.RED)));
                    subTitleCell.BorderWidth = 0;
                    //此Cell的高度會是100單位，注意如果同列只要設定一次，會參照最小值被撐開
                    subTitleCell.MinimumHeight = 20;
                    //將此行文字置中
                    subTitleCell.HorizontalAlignment = Element.ALIGN_CENTER;
                    ptbTitle.AddCell(subTitleCell);

                    //新增了第二個PdfPTable，有四欄(這個是根據gvData的欄數動態產生)
                    PdfPTable ptb = new PdfPTable(Dv.Columns.Count); //
                                                                     //固定每欄的比例(確定欄數才可以這樣做，不然我也不知道會發生什麼事...)
                    float[] widths = new float[] { 10, 15, 10, 28, 10, 40 };
                    ptb.SetWidths(widths);
                    subTitleCell.MinimumHeight = 10;
                    //將Dv標題循環加入PdfPTable表格標題

                    for (int j = 0; j < Dv.Columns.Count; ++j)
                    {
                        Phrase p = new Phrase(Dv.Columns[j].HeaderText, fontContentText);
                        PdfPCell cell = new PdfPCell(p);
                        cell.HorizontalAlignment = Element.ALIGN_CENTER;
                        cell.RunDirection = PdfWriter.RUN_DIRECTION_DEFAULT;
                        ptb.AddCell(cell);
                    }
                    //不特別將Cell取名，直接加入到 PdfPTable中，適用於沒有要做客製的時候
                    //ptb.AddCell(new Phrase(Dv.Columns[h].HeaderText, fontContentText));
                    //將Dv內文循環加入
                    //表格內文
                    for (int i = 0; i < Dv.Rows.Count - 1; ++i)
                    {
                        for (int j = 0; j < Dv.Columns.Count; ++j)
                        {

                            Phrase p = new Phrase(Dv.Rows[i].Cells[j].Value.ToString(), fontContentText);
                            PdfPCell cell = new PdfPCell(p);
                            cell.HorizontalAlignment = Element.ALIGN_CENTER;
                            cell.RunDirection = PdfWriter.RUN_DIRECTION_DEFAULT;
                            ptb.AddCell(cell);
                        }
                    }
                    //新增了第三個PdfPTable，頁尾部分
                    PdfPTable ptfoot = new PdfPTable(4);
                    PdfPCell contentTitle = new PdfPCell(new Phrase("缺席:", FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, 12f, iTextSharp.text.Font.UNDEFINED, BaseColor.RED)));
                    ptfoot.AddCell(contentTitle);
                    PdfPCell content = new PdfPCell(new Phrase("原因:", fontContentText));
                    content.MinimumHeight = 100;
                    //跨欄，同理還有RowSpan跨列
                    content.Colspan = 3;
                    ptfoot.AddCell(content);

                    //依序將三個PdfPTable加入document                  
                    document.Add(ptbTitle);
                    document.Add(ptb);
                    document.Add(ptfoot);

                    //設定PDF檔案相關屬性
                    document.AddTitle("課程點名表");
                    document.AddSubject("This is an Example");
                    document.AddKeywords("Metadata, iTextSharp 5.5.12 ,children");
                    document.AddCreator("iTextSharp 5.5.12");
                    document.AddAuthor("Mark_cheng");
                    document.AddHeader("Class", "KidsList");

                    //關閉PDF檔
                    document.Close();
                    writer.Close();

                    //開啟PDF檔案
                    String FilePath = saveFileDialog1.FileName;
                    System.Diagnostics.Process.Start(FilePath);
                }
            }
            catch (NullReferenceException)
            {
                MessageBox.Show("請選擇課程!", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }




        #endregion

        #region 判斷人數已滿
        public void JudgmentPeople()
        {
            int i = 0;
            foreach (String[] class_table in class_table_list)
            {
                String kid_max = class_table[4];
                String kid_select = class_table[6];
                if (kid_select == "") { kid_select = "0"; }
                int kidmax_int = int.Parse(kid_max, NumberStyles.AllowDecimalPoint);
                int kidselect_int = int.Parse(kid_select, NumberStyles.AllowDecimalPoint);
                int remaining = kidmax_int - kidselect_int;
                if (remaining == 0) { checkedListBox26.SetItemCheckState(i, CheckState.Indeterminate); }
                i++;
            }
        }

        #endregion

        #region 判斷額外課程重複
        public Boolean JudgmentClass()
        {
            Boolean b = false;
            foreach (String[] s in class_info_list)
            {
                if (s[0] == textBox4.Text)
                {                
                    b = true;
                }
            }
            return b;
        }
#endregion
    }
}


