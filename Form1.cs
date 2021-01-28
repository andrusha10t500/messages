using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net.Mail;
using Microsoft.VisualBasic;
using System.IO;
using System.Threading;


namespace messages
{
    public partial class Form1 : Form
    {
        public string ceh;
        public Form1()
        {            
            InitializeComponent();
            string dir = Directory.GetCurrentDirectory();
            if (File.Exists(dir + "\\recovery.txt"))
                button5.Enabled = true;
            
        }
        public void Thread_sending()
        {
            ThreadStart start = new ThreadStart(thread_start);
            Thread Thread = new Thread(start);
            Thread.Start();
        }
        public void thread_start()
        {            
            MailAddress from = new MailAddress(Environment.UserName + "@mail.ru", "Фамилия Имя Отчество");            
            MailMessage m = new MailMessage();
            int count_messages = 0;
            m.From = from;
            m.Subject = textBox1.Text;
            m.Body = textBox2.Text;
            SmtpClient send = new SmtpClient("test.mail.ru");
            string dir = Directory.GetCurrentDirectory();
            if (!Directory.Exists(dir + "\\Output Mail Directory"))
                Directory.CreateDirectory(dir + "\\Output Mail Directory");
            for (int z = 0; z <= dataGridView2.RowCount - 2; z++)
            {
                if (dataGridView2.Rows[z].Cells[0].Value.Equals(true))
                    count_messages++;
            }

            if (MessageBox.Show("будет отправлено " + count_messages + (((Int32)count_messages == 1) ? " письмо. \n" : (((Int32)count_messages < 5 && (Int32)count_messages != 0) ? " письма. \n" : " писем. \n")) +
                "Вы уверены что хотите отправить все сообщения?", "ВНИМАНИЕ!!!", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                for (int i = 0; i <= dataGridView2.RowCount - 2; i++)
                {
                    if (dataGridView2.Rows[i].Cells[0].Value.Equals(true))
                    {
                        m.Attachments.Clear();
                        at(m, i);

                        if (dataGridView1.RowCount == 0)
                        {
                            string a = get_user(((dataGridView2.Rows[i].Cells[1].Value.ToString().Length == 1) ?
                                dataGridView2.Rows[i].Cells[1].Value.ToString().Substring(0, 1) : dataGridView2.Rows[i].Cells[1].Value.ToString().Substring(0, 2))); //по цеху
                            if (a.Length != 0)
                            {
                                string[] t = a.Split(',');
                                m.To.Clear();
                                for (int j = 0; j <= t.Length - 1; j++)
                                {
                                    m.To.Add(t[j]);
                                }
                                send.PickupDirectoryLocation = dir + "\\Output Mail Directory";
                                send.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                                send.Send(m);
                                send.DeliveryMethod = SmtpDeliveryMethod.Network;
                                send.Send(m);
                                File.AppendAllText(dir + "\\log.txt", DateTime.Now.ToString() + " было отправлено цеху " + dataGridView2.Rows[i].Cells[1].Value + " на адрес " 
                                    + m.To + " с " + m.Attachments.Count + " файлами.");
                                File.AppendAllText(dir + "\\log.txt", Environment.NewLine);
                            }                            
                        }
                        else
                        {
                            string t1 = "";
                            for (int k = 0; k <= dataGridView1.RowCount - 2; k++)
                            {
                                if (dataGridView1.Rows[k].Cells[0].Value.Equals(true) &&
                                        dataGridView1.Rows[k].Cells[2].Value.ToString().Equals(dataGridView2.Rows[i].Cells[1].Value.ToString()))
                                {
                                    t1 += dataGridView1.Rows[k].Cells[4].Value.ToString() + ", "; //по участникам                                        
                                }
                            }
                            if (t1 != "")
                            {
                                string[] t = t1.Substring(0, t1.Length - 2).Split(',');
                                m.To.Clear();
                                for (int j = 0; j <= t.Length - 1; j++)
                                {
                                    m.To.Add(t[j]);
                                }
                            }
                            send.PickupDirectoryLocation = dir + "\\Output Mail Directory";
                            send.DeliveryMethod = SmtpDeliveryMethod.SpecifiedPickupDirectory;
                            send.Send(m);
                            send.DeliveryMethod = SmtpDeliveryMethod.Network;
                            send.Send(m);
                            File.AppendAllText(dir + "\\log.txt", DateTime.Now.ToString() + " было отправлено цеху " + dataGridView2.Rows[i].Cells[1].Value + " на адрес " 
                                + m.To + " с " + m.Attachments.Count + " файлами.");
                            File.AppendAllText(dir + "\\log.txt", Environment.NewLine);
                        }
                    }
                }
            }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Thread_sending();
        }
        public void at(MailMessage m,int row)
        {
            //Attachment vl = new Attachment();
            string dir = Directory.GetCurrentDirectory();            
            string[] fl = Directory.GetFiles(dir + "\\" + dataGridView2.Rows[row].Cells[1].Value.ToString());
            foreach (string str in fl)
            {
                Attachment p = new Attachment(str);
                m.Attachments.Add(p);
            }                
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DataBase db = new DataBase();
            string dir = Directory.GetCurrentDirectory();
            Array query;
            if (!db.Open()) Console.WriteLine("Error db!");
            query = db.ExecuteQuery("select distinct cast(substring(cast(CWOC as varchar),1,2) as int) as CWOC1 from other_podrazd order by CWOC1", 1).ToArray();                        
            DataGridViewCheckBoxColumn q = new DataGridViewCheckBoxColumn();
            dataGridView2.Columns.Add(q);
            dataGridView2.Columns[0].HeaderText = "Признак";
            dataGridView2.Columns.Add("col1", "Подразделение");
            //dataGridView2.Columns.Add("col2", "Название");

            dataGridView2.Columns[0].ReadOnly.Equals(false);
            dataGridView2.Columns[1].ReadOnly.Equals(true);
            //dataGridView2.Columns[2].ReadOnly.Equals(true);
            foreach (string[] str in query)
            {
                dataGridView2.Rows.Add(false,(object)str[0]);                
                if (!Directory.Exists(dir + "\\" + str[0]))
                { Directory.CreateDirectory(dir + "\\" + str[0]); }
                if (Directory.GetFiles(dir + "\\" + str[0]).Length != 0)
                {
                    for (int i = 0; i <= dataGridView2.RowCount - 2; i++)
                    {
                        if (dataGridView2.Rows[i].Cells[1].Value.ToString() == str[0])
                            dataGridView2.Rows[i].Cells[0].Value = true;
                    }
                }
            }
        }

        private void dataGridView2_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();            
            if (dataGridView2.SelectedCells[0].Value.GetType().ToString() != "System.Boolean" && dataGridView2.SelectedCells[0].Value.ToString().Length<=2)                
            {                
                DataBase db = new DataBase();                
                if (!db.Open()) Console.WriteLine("Error db!");

                string query_string = "select distinct tcpas027.FIO,cast(substring(cast(PODR  as varchar),1,2) as int) as CEHX,other_roles.Dsca,tcpas026.MAIL from advanced_users  " +
                                        "left join tcpas027 on tcpas027.TNXX=advanced_users.TANO " +
                                        "left join other_roles on advanced_users.CODE=other_roles.Rkod " +
                                        "left join tcpas026 on tcpas027.TNXX=tcpas026.TNXX " +
                                        "where PODR like " +
                                        ((dataGridView2.SelectedCells[0].Value.ToString() == "68") ? "'" + dataGridView2.SelectedCells[0].Value.ToString().Substring(0, 2)
                                        + "%'" : dataGridView2.SelectedCells[0].Value.ToString())
                                        + " and (other_roles.Rkod=3 or other_roles.Rkod=4 or other_roles.Rkod=6)" +
                                        " ORDER BY CEHX";
                Array query;
                query = db.ExecuteQuery(query_string, 1).ToArray();
                DataGridViewCheckBoxColumn q = new DataGridViewCheckBoxColumn();
                dataGridView1.Columns.Add(q);
                dataGridView1.Columns[0].HeaderText = "۷";
                dataGridView1.Columns.Add("col2", "ФИО");
                dataGridView1.Columns.Add("col3", "цех");
                dataGridView1.Columns.Add("col4", "роль");
                dataGridView1.Columns.Add("col5", "MAIL");

                dataGridView1.Columns[0].ReadOnly.Equals(false);
                dataGridView1.Columns[1].ReadOnly.Equals(true);
                dataGridView1.Columns[2].ReadOnly.Equals(true);
                dataGridView1.Columns[3].ReadOnly.Equals(true);
                dataGridView1.Columns[4].ReadOnly.Equals(true);
                foreach (string[] str in query)
                {
                    dataGridView1.Rows.Add(false, (object)str[0], (object)str[1], (object)str[2], (object)str[3]);
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                for (int i = 0; i <= dataGridView2.Rows.Count-1; i++)
                    dataGridView2.Rows[i].Cells[0].Value = true;
            }
            else
            {
                for (int i = 0; i <= dataGridView2.Rows.Count-1; i++)
                    dataGridView2.Rows[i].Cells[0].Value = false;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
        }
        public void second()
        {
            
            DataBase db = new DataBase();
            CheckBox chk = new CheckBox();
            if (!db.Open()) Console.WriteLine("Error db!");
            string query_substr = "";
            for (int i = 0; i <= dataGridView2.RowCount - 2; i++)
            {
                if (dataGridView2.Rows[i].Cells[0].Value.Equals(true))
                {
                    query_substr += "PODR like " + ((dataGridView2.Rows[i].Cells[1].Value.ToString() == "68") ? "'" + dataGridView2.Rows[i].Cells[1].Value.ToString().Substring(0, 2) + "%'" : dataGridView2.Rows[i].Cells[1].Value.ToString()) + " or ";
                }

            }
            if (query_substr.Length != 0)
            {
                string query_string = "select distinct tcpas027.FIO,cast(substring(cast(PODR  as varchar),1,2) as int) as CEHX,other_roles.Dsca,tcpas026.MAIL from advanced_users  " +
                                        "left join tcpas027 on tcpas027.TNXX=advanced_users.TANO " +
                                        "left join other_roles on advanced_users.CODE=other_roles.Rkod " +
                                        "left join tcpas026 on tcpas027.TNXX=tcpas026.TNXX " +
                                        "where (" +
                                        query_substr.Substring(0, query_substr.Length - 4)
                                        + ") and (other_roles.Rkod=3 or other_roles.Rkod=4 or other_roles.Rkod=6)" +
                                        " ORDER BY CEHX";

                Array query;
                query = db.ExecuteQuery(query_string, 1).ToArray();
                DataGridViewCheckBoxColumn q = new DataGridViewCheckBoxColumn();
                dataGridView1.Columns.Add(q);
                dataGridView1.Columns[0].HeaderText = "۷";
                dataGridView1.Columns.Add("col2", "ФИО");
                dataGridView1.Columns.Add("col3", "цех");
                dataGridView1.Columns.Add("col4", "роль");
                dataGridView1.Columns.Add("col5", "MAIL");

                dataGridView1.Columns[0].ReadOnly.Equals(false);
                dataGridView1.Columns[1].ReadOnly.Equals(true);
                dataGridView1.Columns[2].ReadOnly.Equals(true);
                dataGridView1.Columns[3].ReadOnly.Equals(true);
                dataGridView1.Columns[4].ReadOnly.Equals(true);
                foreach (string[] str in query)
                {
                    dataGridView1.Rows.Add(false, (object)str[0], (object)str[1], (object)str[2], (object)str[3]);
                }
            }
            else { MessageBox.Show("Выберете подразделение галочкой!", "Сообщение"); }
            
        }
        private void button4_Click(object sender, EventArgs e)
        {
            second();
            button1.Enabled = true;
        }
        public string get_user(string podr)
        {
            DataBase db = new DataBase();
            if (!db.Open()) Console.WriteLine("Error db!");
            string query_string = "select distinct tcpas026.MAIL from advanced_users  " +
                                    "left join tcpas027 on tcpas027.TNXX=advanced_users.TANO " +
                                    "left join other_roles on advanced_users.CODE=other_roles.Rkod " +
                                    "left join tcpas026 on tcpas027.TNXX=tcpas026.TNXX " +
                                    "where podr=" +
                                    podr
                                    + " and (other_roles.Rkod=3 or other_roles.Rkod=4 or other_roles.Rkod=6) " +
                                    "and tcpas026.MAIL is not null";

            Array query;
            string ret="";
            query = db.ExecuteQuery(query_string, 1).ToArray();
            if (query.Length != 0)
            {
                foreach (string[] str in query)
                {
                    try
                    {
                        foreach (string str1 in str)
                        {
                            ret += str1 + ", ";
                        }
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.Message);
                    }
                }                
            }
            else
            {
                ret = "  ";
            }
            return ret.Substring(0,ret.Length-2);
        }
        public void saves()
        {
            string dir = Directory.GetCurrentDirectory();
            using (StreamWriter sw = File.CreateText(dir + "\\recovery.txt"))
            {
                string str="";
                for (int k = 0; k <= dataGridView2.RowCount - 2; k++)
                {
                    if (dataGridView2.Rows[k].Cells[0].Value.Equals(true))
                    {
                        str += "podr like " + ((dataGridView2.Rows[k].Cells[1].Value.ToString() == "68") ? ("'" + 
                            dataGridView2.Rows[k].Cells[1].Value.ToString() + "%'") : (dataGridView2.Rows[k].Cells[1].Value.ToString()))  + " or ";
                            
                    }
                }
                
                DataBase db = new DataBase();
                if (!db.Open()) Console.WriteLine("Error db!");
                Array query = db.ExecuteQuery("select distinct cast(substring(cast(PODR  as varchar),1,2) as int) as CEHX,tcpas026.MAIL from advanced_users  " +
                                "left join tcpas027 on tcpas027.TNXX=advanced_users.TANO " +
                                "left join other_roles on advanced_users.CODE=other_roles.Rkod " +
                                "left join tcpas026 on tcpas027.TNXX=tcpas026.TNXX " +
                                "where (" + str.Substring(0, str.Length - 4) + ") and (other_roles.Rkod=3 or other_roles.Rkod=4 or other_roles.Rkod=6)", 1).ToArray();
                foreach (string[] stri in query)
                {                    
                    for (int i = 0; i <= dataGridView2.RowCount - 2; i++)
                    {                        
                        if (dataGridView2.Rows[i].Cells[0].Value.Equals(true))
                        {
                            sw.WriteLine(dataGridView2.Rows[i].Cells[1].Value.ToString());
                            for(int j=0; j<=dataGridView1.RowCount-2; j++)
                            {                                
                                if (dataGridView1.Rows[j].Cells[0].Value.Equals(true) && (stri[0] == dataGridView2.Rows[i].Cells[1].Value.ToString() &&
                                    stri[1] == dataGridView1.Rows[j].Cells[4].Value.ToString()))
                                {
                                    //sw.WriteLine(dataGridView2.Rows[i].Cells[1].Value.ToString());
                                    sw.WriteLine(dataGridView1.Rows[j].Cells[4].Value.ToString());
                                }
                            }
                        }
                    }                    
                }
                sw.Close();
            }
            button5.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {            
            saves();
        }
        public void recovery()
        {
            dataGridView1.Rows.Clear();
            string dir = Directory.GetCurrentDirectory();
            string[] str = new string[InFile(dir + "\\recovery.txt")];
            if (File.Exists(dir + "\\recovery.txt") && MessageBox.Show("есть соханенные данные, восстановить?", "Сообщение", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (InFile(dir + "\\recovery.txt") != 0)
                {                    
                    using (StreamReader sr = File.OpenText(dir + "\\recovery.txt"))
                    {
                        int i = 0;
                        while (!sr.EndOfStream)
                        {
                            str.SetValue(sr.ReadLine(), i);
                            i++;
                        }
                        foreach (string str1 in str)
                        {
                            for (int j = 0; j <= dataGridView2.RowCount - 2; j++)
                            {
                                if (dataGridView2.Rows[j].Cells[1].Value.Equals(str1))
                                {
                                    dataGridView2.Rows[j].Cells[0].Value = true;
                                }
                            }
                        }
                        second();
                        foreach (string str1 in str)
                        {
                            for (int k = 0; k <= dataGridView1.RowCount - 2; k++)
                            {
                                if (dataGridView1.Rows[k].Cells[4].Value.Equals(str1))
                                {
                                    dataGridView1.Rows[k].Cells[0].Value = true;
                                }
                            }
                        }
                        sr.Close();
                    }
                }
                else { MessageBox.Show("Файл пуст", "Сообщение"); }
            }
        }

        public int InFile(string FilePath)
        {
            int number = 0;
            using (StreamReader file1 = File.OpenText(FilePath))
            {
                while (!file1.EndOfStream)
                {
                    number++;
                    file1.ReadLine();
                }
                file1.Close();
            }
            return number;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            recovery();
        }

        private void dataGridView1_RowsAdded(object sender, DataGridViewRowsAddedEventArgs e)
        {
            button1.Enabled = true;
        }

        private void Form1_Resize(object sender, EventArgs e)
        {
            
            dataGridView1.Width = this.Width-409;
            //685-276
            dataGridView1.Height = this.Height-136;
            dataGridView2.Height = this.Height - 136;
            //347-211
            //dataGridView2.Width = this.Width - 410;
            //685-275
            button3.Top = this.Height - 73;
            button2.Top = this.Height - 73;
            //347-274
            button3.Left=this.Width -191;
            //685-494
            button2.Left= this.Width - 110;
            //685-575
        }
    }
}

