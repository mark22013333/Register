using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading.Tasks;
using System.Data;
using System.IO;// for check file exist
using System.Data.SQLite;//for SQLite

namespace Register
{
    class SQLiteHelper
    {
        private SQLiteConnection sqlite_connect;// SQLite連接物件
        private SQLiteCommand sqlite_cmd;// SQLite SQL指令物件

        public SQLiteHelper()
        {
            Init();
        }

        #region 資料庫初始化
        private void Init()
        {
            // 判斷新增SQLite資料庫
            if (!File.Exists(Application.StartupPath + @"\kids.db"))
            {
                SQLiteConnection.CreateFile("kids.db");
                CreateTable();
            }
            // new SQLiteConnection物件
            this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
        }
        #endregion

        #region 額外新增課程
        public Boolean Classadd(string text)
        {
            Boolean b = false;
            StringBuilder sql = new StringBuilder();
            String class_num = "";
            try
            {
                this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
                this.sqlite_connect.Open(); // Open
                this.sqlite_cmd = sqlite_connect.CreateCommand(); //create command

                String max_str = "SELECT max(class_type) FROM classinfo;";
                SQLiteCommand insertSQL = new SQLiteCommand(sqlite_connect);
                insertSQL.CommandType = CommandType.Text;
                insertSQL.CommandText = max_str.ToString();
                SQLiteDataReader sqlite_datareader = insertSQL.ExecuteReader();
                while (sqlite_datareader.Read())
                {
                    class_num = sqlite_datareader["max(class_type)"].ToString();
                }
                int classnum = Convert.ToInt16(class_num) + 1;
                sql.Append("INSERT INTO classinfo (class_name , class_type) VALUES ( '" + text + "' , " + classnum + " ); ");
                this.sqlite_cmd.CommandText = sql.ToString();
                this.sqlite_cmd.ExecuteNonQuery();

                b = true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
            finally
            {
                this.sqlite_connect.Close();
            }
            return b;
        }
#endregion

        #region 建立資料庫表格
        /**
         * 建立資料庫表格 **/
        private void CreateTable()
        {
            StringBuilder sql = new StringBuilder();
            try
            {
                this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
                this.sqlite_connect.Open(); // Open
                this.sqlite_cmd = sqlite_connect.CreateCommand(); //create command

                sql.Append("CREATE TABLE IF NOT EXISTS kidinfo (");
                sql.Append("nid integer PRIMARY KEY AUTOINCREMENT ," +
                    "kidname TEXT ," +
                    "parentsname TEXT ," +
                    "parentsphone TEXT ," +
                    "sex TEXT ," +
                    "birthday DATE ," +
                    "age REAL ," +
                    "money BLOB ," +
                    "rec_date DATE DEFAULT (datetime('now','localtime')) , " +
                    "seq TEXT NOT NULL unique");
                sql.Append(")");

                this.sqlite_cmd.CommandText = sql.ToString();
                this.sqlite_cmd.ExecuteNonQuery();

                sql.Clear();
                sql.Append("CREATE TABLE IF NOT EXISTS classinfo (");
                sql.Append("nid integer PRIMARY KEY AUTOINCREMENT ,class_name TEXT ,class_type int ,rec_date DATE DEFAULT (datetime('now','localtime'))");
                sql.Append(")");

                this.sqlite_cmd.CommandText = sql.ToString();
                this.sqlite_cmd.ExecuteNonQuery();

                sql.Clear();
                sql.Append("INSERT INTO classinfo (class_name , class_type) VALUES ( '音樂' , 1 ); ");
                sql.Append("INSERT INTO classinfo (class_name , class_type) VALUES ( '樂高' , 2 ); ");
                sql.Append("INSERT INTO classinfo (class_name , class_type) VALUES ( '美術' , 3 ); ");
                sql.Append("INSERT INTO classinfo (class_name , class_type) VALUES ( '捏塑' , 4 ); ");
                sql.Append("INSERT INTO classinfo (class_name , class_type) VALUES ( 'LASY' , 5 ); ");
                sql.Append("INSERT INTO classinfo (class_name , class_type) VALUES ( '體能' , 6 ); ");
                sql.Append("INSERT INTO classinfo (class_name , class_type) VALUES ( '料理' , 7 ); ");
                sql.Append("INSERT INTO classinfo (class_name , class_type) VALUES ( '桌遊' , 8 ); ");
                sql.Append("INSERT INTO classinfo (class_name , class_type) VALUES ( '自然' , 9 ); ");
                sql.Append("INSERT INTO classinfo (class_name , class_type) VALUES ( '故事' ,10 ); ");
                sql.Append("INSERT INTO classinfo (class_name , class_type) VALUES ( '講座' ,11 ); ");
                sql.Append("INSERT INTO classinfo (class_name , class_type) VALUES ( '抓周' ,12 ); ");
                sql.Append("INSERT INTO classinfo (class_name , class_type) VALUES ( '收涎' ,13 ); ");
                sql.Append("INSERT INTO classinfo (class_name , class_type) VALUES ( '科學' ,14 ); ");

                this.sqlite_cmd.CommandText = sql.ToString();
                this.sqlite_cmd.ExecuteNonQuery();

                sql.Clear();
                sql.Append("CREATE TABLE IF NOT EXISTS classtable (");
                sql.Append("nid integer PRIMARY KEY AUTOINCREMENT ,class_name TEXT ,class_type int ,class_datetime DATE ,class_age_range TEXT ,kid_max INT ,rec_date DATE DEFAULT (datetime('now','localtime')) ,seq TEXT NOT NULL unique");
                sql.Append(")");

                this.sqlite_cmd.CommandText = sql.ToString();
                this.sqlite_cmd.ExecuteNonQuery();

                sql.Clear();
                sql.Append("CREATE TABLE IF NOT EXISTS selectclass (");
                sql.Append("nid integer PRIMARY KEY AUTOINCREMENT ,class_type int , kid_nid int ,rec_date DATE DEFAULT (datetime('now','localtime')) ,seq TEXT NOT NULL unique");
                sql.Append(")");

                this.sqlite_cmd.CommandText = sql.ToString();
                this.sqlite_cmd.ExecuteNonQuery();
                //this.sqlite_cmd.CommandText = "INSERT INTO phone VALUES (null, '" + name + "','" + phone_number + "');";
                //this.sqlite_cmd.ExecuteNonQuery();//using behind every write cmd
                //this.sqlite_cmd.CommandText = "INSERT INTO kidinfo (kidname, parentsname ,parentsphone,sex,birthday,age) VALUES ('test', 'test','5678','測試1','1989',5);";
                //this.sqlite_cmd.ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                this.sqlite_connect.Close();
            }
        }
        #endregion

        #region 新增幼兒資訊
        /**
         * 新增幼兒資訊 **/
        public Boolean Add_KidInfo(String KidName, String ParentsName, String ParentsPhone, Boolean sex, String birthday, String age)
        {
            Boolean b = false;
            try
            {
                this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
                this.sqlite_connect.Open(); // Open
                SQLiteCommand insertSQL = new SQLiteCommand("INSERT INTO kidinfo (kidname, parentsname ,parentsphone,sex,birthday,age,seq) VALUES (? ,? ,? ,? ,? ,? ,?) ", sqlite_connect);

                insertSQL.Parameters.Add(new SQLiteParameter("@kidname", KidName.ToString()));
                insertSQL.Parameters.Add(new SQLiteParameter("@parentsname", ParentsName.ToString()));
                insertSQL.Parameters.Add(new SQLiteParameter("@parentsphone", ParentsPhone.ToString()));
                insertSQL.Parameters.Add(new SQLiteParameter("@sex", sex == true ? "男" : "女"));
                insertSQL.Parameters.Add(new SQLiteParameter("@birthday", birthday.ToString()));
                insertSQL.Parameters.Add(new SQLiteParameter("@age", age.ToString()));
                insertSQL.Parameters.Add(new SQLiteParameter("@seq", KidName + ParentsPhone));

                insertSQL.ExecuteNonQuery();
                b = true;
            }
            catch (Exception)
            {
                //throw new Exception(ex.Message);
            }
            finally
            {
                this.sqlite_connect.Close();
            }
            return b;
        }
        #endregion

        #region 刪除幼兒資訊
        /**
         * 刪除幼兒資訊 **/
        public Boolean Del_Kidinfo(string delnum)
        {
            Boolean a = false;
            StringBuilder sql = new StringBuilder();
            try
            {
                this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
                this.sqlite_connect.Open(); // Open
                this.sqlite_cmd = sqlite_connect.CreateCommand();
                sql.Append("DELETE FROM selectclass WHERE kid_nid = '" + delnum + "';");
                sql.Append("DELETE FROM kidinfo WHERE nid = '" + delnum + "';");
                this.sqlite_cmd.CommandText = sql.ToString();
                this.sqlite_cmd.ExecuteNonQuery();
                a = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                this.sqlite_connect.Close();
            }
            return a;
        }
        #endregion

        #region 查詢幼兒資訊
        /**
         * 查詢幼兒資訊 **/
        public Boolean FindKid_Info(String KidName, out List<String[]> list)
        {
            Boolean b = false;
            try
            {
                this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
                this.sqlite_connect.Open(); // Open
                list = new List<String[]>();

                //SELECT * FROM kidinfo WHERE CharIndex( ? ,kidname) > 0
                //SELECT * FROM kidinfo WHERE  kidname LIKE '%"+ KidName + "%';"
                String sql = "SELECT * FROM kidinfo WHERE  kidname LIKE '%" + KidName + "%';";
                //MessageBox.Show(sql);
                SQLiteCommand insertSQL = new SQLiteCommand(sqlite_connect);
                insertSQL.CommandType = CommandType.Text;
                insertSQL.CommandText = sql.ToString();
                //insertSQL.Parameters.Add(new SQLiteParameter("@kidname", KidName.ToString()));

                // 執行查詢塞入 sqlite_datareader
                SQLiteDataReader sqlite_datareader = insertSQL.ExecuteReader();

                String kidname = "";
                String parentsname = "";
                String parentsphone = "";
                String sex = "";
                String birthday = "";
                String money = "";
                String rec_date = "";
                String nid = "";
                int count = 0;
                String[] data;
                // 一筆一筆列出查詢的資料
                while (sqlite_datareader.Read())
                {
                    //MessageBox.Show(sqlite_datareader["kidname"].ToString());
                    data = new String[8];
                    count += 1;
                    kidname = sqlite_datareader["kidname"].ToString();
                    parentsname = sqlite_datareader["parentsname"].ToString();
                    parentsphone = sqlite_datareader["parentsphone"].ToString();
                    sex = sqlite_datareader["sex"].ToString();
                    //birthday = sqlite_datareader["age"].ToString();
                    birthday = sqlite_datareader.GetString(6); ;//sqlite_datareader["birthday"].ToString();
                    rec_date = sqlite_datareader["rec_date"].ToString();
                    nid = sqlite_datareader["nid"].ToString();
                    money = sqlite_datareader["money"].ToString();

                    data[0] = kidname;
                    data[1] = parentsname;
                    data[2] = parentsphone;
                    data[3] = sex;
                    data[4] = birthday;
                    data[5] = rec_date;
                    data[6] = nid;
                    data[7] = money;
                    list.Add(data);
                    //MessageBox.Show(s);
                }
                //MessageBox.Show(count.ToString());
                if (count >= 1) b = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                this.sqlite_connect.Close();
            }
            return b;
        }
        #endregion

        #region 刪除幼兒的課程資訊
        /**
         * 刪除課程資訊 **/
        public Boolean Del_Class(string delnum)
        {
            Boolean a = false;
            try
            {
                this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
                this.sqlite_connect.Open(); // Open
                this.sqlite_cmd = sqlite_connect.CreateCommand();
                string sql = ("DELETE FROM selectclass WHERE nid = '" + delnum + "'");
                this.sqlite_cmd.CommandText = sql.ToString();
                this.sqlite_cmd.ExecuteNonQuery();

                a = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                this.sqlite_connect.Close();
            }
            return a;
        }
        #endregion

        #region 刪除選課資訊
        public void Del_Selected_class(String class_type)
        {
            //Boolean b = false;
            try
            {
                this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
                this.sqlite_connect.Open(); // Open
                this.sqlite_cmd = sqlite_connect.CreateCommand();
                string sql = ("DELETE FROM classtable WHERE nid = '" + class_type + "'");
                this.sqlite_cmd.CommandText = sql.ToString();
                this.sqlite_cmd.ExecuteNonQuery();
                
                //b = true;
            }
            catch (Exception)
            {
                // throw new Exception(ex.Message);
            }
            finally
            {
                this.sqlite_connect.Close();
            }
            //return b;
        }
        #endregion

        #region 刪除全部資料
        /**
         * 刪除全部資料 **/
        public Boolean Del_Alldata()
        {
            Boolean a = false;
            StringBuilder sql = new StringBuilder();
            try
            {
                this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
                this.sqlite_connect.Open(); // Open
                this.sqlite_cmd = sqlite_connect.CreateCommand();

                sql.Append("DROP TABLE classinfo;");
                sql.Append("DROP TABLE kidinfo;");
                sql.Append("DROP TABLE classtable;");
                sql.Append("DROP TABLE selectclass;");

                this.sqlite_cmd.CommandText = sql.ToString();
                this.sqlite_cmd.ExecuteNonQuery();

                CreateTable();
                a = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                this.sqlite_connect.Close();
            }
            return a;
        }
        #endregion

        #region 刪除全部課程
        public Boolean Del_AllClass()
        {
            Boolean a = false;
            StringBuilder sql = new StringBuilder();
            try
            {
                this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
                this.sqlite_connect.Open(); // Open
                this.sqlite_cmd = sqlite_connect.CreateCommand();
              
                sql.Append("DROP TABLE classtable;");
                sql.Append("DROP TABLE selectclass;");

                this.sqlite_cmd.CommandText = sql.ToString();
                this.sqlite_cmd.ExecuteNonQuery();

                sql.Clear();
                sql.Append("CREATE TABLE IF NOT EXISTS classtable (");
                sql.Append("nid integer PRIMARY KEY AUTOINCREMENT ,class_name TEXT ,class_type int ,class_datetime DATE ,class_age_range TEXT ,kid_max INT ,rec_date DATE DEFAULT (datetime('now','localtime')) ,seq TEXT NOT NULL unique");
                sql.Append(")");

                this.sqlite_cmd.CommandText = sql.ToString();
                this.sqlite_cmd.ExecuteNonQuery();

                sql.Clear();
                sql.Append("CREATE TABLE IF NOT EXISTS selectclass (");
                sql.Append("nid integer PRIMARY KEY AUTOINCREMENT ,class_type int , kid_nid int ,rec_date DATE DEFAULT (datetime('now','localtime')) ,seq TEXT NOT NULL unique");
                sql.Append(")");

                this.sqlite_cmd.CommandText = sql.ToString();
                this.sqlite_cmd.ExecuteNonQuery();

                a = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                this.sqlite_connect.Close();
            }
            return a;
        }
        #endregion

        #region 新增課程
        /**
         * 新增課程 **/
        public Boolean Add_Class_List(String class_name, int class_type, String class_datetime, int class_age_range, int kid_max)
        {
            Boolean b = false;
            try
            {
                this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
                this.sqlite_connect.Open(); // Open


                SQLiteCommand insertSQL = new SQLiteCommand("INSERT INTO classtable (class_name ,class_type ,class_datetime ,class_age_range ,kid_max ,seq) VALUES (? ,? ,? ,? ,? ,?) ", sqlite_connect);

                insertSQL.Parameters.Add(new SQLiteParameter("@class_name", class_name.ToString()));
                insertSQL.Parameters.Add(new SQLiteParameter("@class_type", class_type.ToString()));
                insertSQL.Parameters.Add(new SQLiteParameter("@class_datetime", class_datetime.ToString()));
                insertSQL.Parameters.Add(new SQLiteParameter("@class_age_range", class_age_range));
                insertSQL.Parameters.Add(new SQLiteParameter("@kid_max", kid_max.ToString()));
                insertSQL.Parameters.Add(new SQLiteParameter("@seq", class_datetime + class_type));// 唯一值

                insertSQL.ExecuteNonQuery();
                b = true;
            }
            catch (Exception)
            {
                //throw new Exception(ex.Message);
            }
            finally
            {
                this.sqlite_connect.Close();
            }
            return b;
        }
        #endregion

        #region 對應幼兒選課
        public Boolean Set_Selected_class(String class_type, String kid_nid)
        {
            Boolean b = false;
            try
            {
                this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
                this.sqlite_connect.Open(); // Open

                SQLiteCommand insertSQL = new SQLiteCommand("INSERT INTO selectclass ( class_type ,kid_nid ,seq) VALUES ( ? ,? ,? ) ", sqlite_connect);

                insertSQL.Parameters.Add(new SQLiteParameter("@class_type", class_type.ToString()));
                insertSQL.Parameters.Add(new SQLiteParameter("@kid_nid", kid_nid.ToString()));
                insertSQL.Parameters.Add(new SQLiteParameter("@seq", class_type.ToString() + kid_nid.ToString()));

                insertSQL.ExecuteNonQuery();
                b = true;
            }
            catch (Exception)
            {
                // throw new Exception(ex.Message);
            }
            finally
            {
                this.sqlite_connect.Close();
            }
            return b;
        }
        #endregion

        #region 取得課程資訊
        /**
         * 取得課程資訊 **/
        public Boolean GetClassInfo(out List<String[]> list)
        {
            Boolean b = false;
            list = new List<String[]>();
            try
            {
                this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
                this.sqlite_connect.Open(); // Open

                String sql = "SELECT * FROM classinfo;";
                SQLiteCommand insertSQL = new SQLiteCommand(sqlite_connect);
                insertSQL.CommandType = CommandType.Text;
                insertSQL.CommandText = sql.ToString();

                SQLiteDataReader sqlite_datareader = insertSQL.ExecuteReader();

                String class_name = "";
                String class_type = "";
                int count = 0;
                String[] data;

                // 一筆一筆列出查詢的資料
                while (sqlite_datareader.Read())
                {
                    data = new String[2];
                    count += 1;
                    class_name = sqlite_datareader["class_name"].ToString();
                    class_type = sqlite_datareader["class_type"].ToString();
                    data[0] = class_name;
                    data[1] = class_type;
                    list.Add(data);
                }
                if (count >= 1) b = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                this.sqlite_connect.Close();
            }
            return b;
        }
        #endregion

        #region 列出課程列表
        /**
         * 列出課程列表
         **/
        public Boolean GetClassList(String dateTime, out List<String[]> list)
        {
            Boolean b = false;
            list = new List<String[]>();
            try
            {
                this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
                this.sqlite_connect.Open(); // Open

                //String sql = "SELECT * FROM classtable a ,classinfo b ";
                // sql += "WHERE a.class_type = b.class_type AND class_datetime LIKE '" + dateTime + "%' ORDER BY class_datetime ASC;";
                String sql = "SELECT a.* , b.* ,c.class_type,c.count FROM classtable a";
                sql += " left join classinfo b on a.class_type = b.class_type";
                sql += " left join(select class_type, count(*) count from selectclass group by class_type) c on a.nid = c.class_type";
                sql += " WHERE a.class_datetime like '" + dateTime + "%' ORDER BY class_datetime ASC;";

                SQLiteCommand insertSQL = new SQLiteCommand(sqlite_connect);
                insertSQL.CommandType = CommandType.Text;
                insertSQL.CommandText = sql.ToString();

                SQLiteDataReader sqlite_datareader = insertSQL.ExecuteReader();


                String class_name = "";
                String class_type = "";
                String class_type_number = "";
                String class_datetime = "";
                String class_age_range = "";
                String kid_max = "";
                String kid_select_class = "";

                int count = 0;
                String[] data;

                // 一筆一筆列出查詢的資料
                while (sqlite_datareader.Read())
                {
                    data = new String[7];
                    count += 1;
                    class_name = sqlite_datareader["class_name"].ToString();
                    class_type_number = sqlite_datareader.GetInt32(0).ToString();//sqlite_datareader["nid"].ToString();
                    class_type = sqlite_datareader.GetString(9);
                    class_datetime = sqlite_datareader.GetString(3);
                    class_age_range = sqlite_datareader["class_age_range"].ToString();
                    kid_max = sqlite_datareader["kid_max"].ToString();
                    kid_select_class = sqlite_datareader["count"].ToString();

                    data[0] = class_name;
                    data[1] = class_type;
                    data[2] = class_datetime;
                    data[3] = class_age_range;
                    data[4] = kid_max;
                    data[5] = class_type_number;
                    data[6] = kid_select_class;
                    list.Add(data);                              
                }
                if (count >= 1) b = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                this.sqlite_connect.Close();
            }
            return b;
        }
        #endregion     

        #region 查某課程幼兒選課
        /**
         * 查某課程有哪些幼兒選課
         **/
        public Boolean GetSelected_kid_Class_List(String class_nid, out List<String[]> list)
        {
            //MessageBox.Show(class_nid);
            Boolean b = false;
            list = new List<String[]>();
            try
            {
                this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
                this.sqlite_connect.Open(); // Open

                String sql = "SELECT * FROM selectclass a, kidinfo b ";
                sql += "WHERE a.kid_nid = b.nid AND a.class_type ='" + class_nid + "';"; //+class_nid +

                SQLiteCommand insertSQL = new SQLiteCommand(sqlite_connect);
                insertSQL.CommandType = CommandType.Text;
                insertSQL.CommandText = sql.ToString();

                SQLiteDataReader sqlite_datareader = insertSQL.ExecuteReader();

                String kid_name = "";
                String kid_cell = "";
                String kid_sex = "";
                String kid_birthday = "";
                String kid_age = "";

                int count = 0;
                String[] data;

                // 一筆一筆列出查詢的資料
                while (sqlite_datareader.Read())
                {

                    data = new String[5];
                    count += 1;
                    kid_name = sqlite_datareader.GetString(6);
                    kid_cell = sqlite_datareader.GetString(8);
                    kid_sex = sqlite_datareader.GetString(9);
                    kid_birthday = sqlite_datareader.GetString(10);
                    kid_age = sqlite_datareader.GetString(11);

                    //MessageBox.Show(kid_name);

                    data[0] = kid_name;
                    data[1] = kid_cell;
                    data[2] = kid_sex;
                    data[3] = kid_birthday;
                    data[4] = kid_age;

                    list.Add(data);
                }
                if (count >= 1) b = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                this.sqlite_connect.Close();
            }

            return b;
        }
        #endregion

        #region 查幼兒選課資訊
        /**
         * 查幼兒選課資訊
        **/
        public Boolean Getkid_Selected_Class_List(String kid_nid, out List<String[]> list)
        {
            Boolean b = false;
            list = new List<String[]>();
            // MessageBox.Show(kid_nid);
            try
            {
                this.sqlite_connect = new SQLiteConnection("Data source=kids.db");
                this.sqlite_connect.Open(); // Open

                String sql = "SELECT * FROM selectclass a, classtable b ,classinfo c ";
                sql += "WHERE a.class_type = b.nid AND b.class_type = c.class_type AND a.kid_nid = '" + kid_nid + "';";

                //String sql  = "SELECT * FROM selectclass a, classtable b ,classinfo c ";
                //sql += "WHERE a.class_type = b.class_type AND b.class_type = c.class_type AND a.kid_nid ='" + kid_nid + "';";

                SQLiteCommand insertSQL = new SQLiteCommand(sqlite_connect);
                insertSQL.CommandType = CommandType.Text;
                insertSQL.CommandText = sql.ToString();

                SQLiteDataReader sqlite_datareader = insertSQL.ExecuteReader();

                int count = 0;
                String[] data;
                String class_datetime;
                String class_name;
                String class_nid;

                // 一筆一筆列出查詢的資料
                while (sqlite_datareader.Read())
                {

                    data = new String[3];
                    count += 1;

                    class_datetime = sqlite_datareader.GetString(8);
                    class_name = sqlite_datareader.GetString(14);
                    class_nid = sqlite_datareader.GetInt32(0).ToString();
                    //MessageBox.Show(kid_name);

                    data[0] = class_datetime;
                    data[1] = class_name;
                    data[2] = class_nid;
                    list.Add(data);
                }
                if (count >= 1) b = true;
            }
            catch (Exception ex)
            {
                throw new Exception(ex.Message);
            }
            finally
            {
                this.sqlite_connect.Close();
            }
            return b;
        }
        #endregion

        #region 備份資料庫
        public void BackupDB()
        {
            try
            {
                DialogResult myResult = MessageBox.Show("是否備份資料庫?", "備份", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                string fileName = "kids.db";
                string nowdate = DateTime.Now.ToString("yyyy-MM-dd");
                string nowtime = DateTime.Now.ToString("HH：mm");
                string sourcePath = Environment.CurrentDirectory;
                string targetPath = Environment.CurrentDirectory + "\\" + nowdate + " Time-" + nowtime + " DB-Backup";
                string sourceFile = Path.Combine(sourcePath, fileName);
                string destFile = Path.Combine(targetPath, fileName);

                if (myResult == DialogResult.Yes)
                {
                    if (!Directory.Exists(targetPath))
                    {
                        Directory.CreateDirectory(targetPath);
                        File.Copy(sourceFile, destFile, true);
                    }
                    MessageBox.Show(" 備份完成!", "時間 " + nowdate + " Time-" + nowtime, MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex) { ex.Message.ToString(); }
        }
        #endregion

        #region 還原資料庫
        public void RestoreDB()
        {
            try
            {
                string strPath = "";
                string sourcePath = "";
                string fileName = "kids.db";
                string targetPath = Environment.CurrentDirectory;
                string destFile = Path.Combine(targetPath, fileName);

                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Title = "請選擇要還原的資料庫檔案";
                dialog.InitialDirectory = Environment.CurrentDirectory;
                dialog.Filter = "db (*.*)|*.db";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    strPath = dialog.FileName;
                    sourcePath = Path.GetFullPath(strPath);
                    File.Copy(sourcePath, destFile, true);
                    MessageBox.Show("資料已還原", "提示!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex) { MessageBox.Show(ex.Message.ToString()); }
        }
        #endregion
    }
}

#region 註解
/*
 * 
 *  // String 總選課筆數 = "SELECT COUNT(class_type) FROM selectclass";
            // String 某時段課程人數 = "SELECT COUNT(class_type) FROM selectclass WHERE class_type = '1'"; 
            //select class_type , count(*) from selectclass where class_type in ('1','2','3','4','5') group by class_type

            //select class_type, count(*) from selectclass where class_type > '0' and class_type <= '100' group by class_type
            //select class_type , count(*) from selectclass group by class_type

            
              select a.* , b.* ,c.class_type,c.count from classtable a
              left join classinfo b on a.class_type = b.class_type
              left join(select class_type, count(*) count from selectclass group by class_type) c on a.nid = c.class_type
              where a.class_datetime like '2017-08%'
              ORDER BY class_datetime ASC;
            
 *
    FontFactory.RegisterDirectories();
    iTextSharp.text.Font myfont = FontFactory.GetFont("微軟正黑體", BaseFont.IDENTITY_H, 8, iTextSharp.text.Font.UNDEFINED, BaseColor.BLACK);
    Document pdfDoc = new Document(PageSize.A4, 10f,-10,10f,1);

    pdfDoc.Open();
    PdfWriter wri = PdfWriter.GetInstance(pdfDoc, new FileStream(FilePath, FileMode.Create));
    pdfDoc.Open();
    PdfPTable _mytable = new PdfPTable(Dv.ColumnCount);

    for (int j = 0; j < Dv.Columns.Count; ++j)
    {
        Phrase p = new Phrase(Dv.Columns[j].HeaderText, myfont);
        PdfPCell cell = new PdfPCell(p);
        cell.HorizontalAlignment = Element.ALIGN_CENTER;
        cell.RunDirection = PdfWriter.RUN_DIRECTION_DEFAULT;
        _mytable.AddCell(cell);
    }
    //-------------------------

    for (int i = 0; i < Dv.Rows.Count - 1; ++i)
    {
        for (int j = 0; j < Dv.Columns.Count; ++j)
        {

            Phrase p = new Phrase(Dv.Rows[i].Cells[j].Value.ToString(), myfont);
            PdfPCell cell = new PdfPCell(p);
            cell.HorizontalAlignment = Element.ALIGN_RIGHT; 
            cell.RunDirection = PdfWriter.RUN_DIRECTION_RTL;
            _mytable.AddCell(cell);
        }
    }


    //------------------------           
    pdfDoc.Add(_mytable);
    pdfDoc.Close();
    System.Diagnostics.Process.Start(FilePath);
}
*/
/* public bool ExportDataGridview(DataGridView gridView, bool isShowExcle)
 {
     if (gridView.Rows.Count == 0)
         return false;
     string pathFile = @"D:\test";

     Excel.Application excelApp;
     Excel._Workbook wBook;
     Excel._Worksheet wSheet;
     Excel.Range wRange;

     // 開啟一個新的應用程式
     excelApp = new Excel.Application();

     // 讓Excel文件可見
     excelApp.Visible = true;

     // 停用警告訊息
     excelApp.DisplayAlerts = false;

     // 加入新的活頁簿
     excelApp.Workbooks.Add(Type.Missing);

     // 引用第一個活頁簿
     wBook = excelApp.Workbooks[1];

     // 設定活頁簿焦點
     wBook.Activate();

     try
     {
         // 引用第一個工作表
         wSheet = (Excel._Worksheet)wBook.Worksheets[1];

         // 命名工作表的名稱
         wSheet.Name = "簽到表";

         // 設定工作表焦點
         wSheet.Activate();

         //excelApp.Cells[1, 1] = "Excel測試";

         //建立Excel对象
         //Excel.Application excel = new Excel.Application();
         // excel.Application.Workbooks.Add(true);
         // excel.Visible = isShowExcle;
         //生成字段名称

         // 設定第1列顏色
         wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 7]];
         wRange.Select();
         wRange.Font.Color = ColorTranslator.ToOle(Color.White);
         wRange.Interior.Color = ColorTranslator.ToOle(Color.DimGray);
         //填充数据
         for (int i = 0; i < gridView.ColumnCount; i++)
         {
             excelApp.Cells[1, i + 1] = gridView.Columns[i].HeaderText;
         }
         for (int i = 0; i < gridView.RowCount - 1; i++)
         {                                   
             for (int j = 0; j < gridView.ColumnCount; j++)
             {
                 if (gridView[j, i].ValueType == typeof(string))
                 {
                     excelApp.Cells[i + 2, j + 1] = "'" + gridView[j, i].Value.ToString();                            
                 }
                 else
                 {
                     excelApp.Cells[i + 2, j + 1] = gridView[j, i].Value.ToString();
                 }                                               
             }
         }
         // 自動調整欄寬
         // wRange = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[5, 2]];
        // wRange.Columns.AutoFit();
     }
     catch { }
     return true;*/
#endregion
