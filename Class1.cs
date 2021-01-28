using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Data;


namespace messages
{
    public class DataBase
    {
        private SqlConnection con;
        
	//Строка подключения к Oracle
        private const string conStr = @"Server=; DataBase=;" +
                                      @"User id=; Password=";

        

        public bool Open()
        {
            bool result = true;
            try
            {
                con = new SqlConnection(conStr);
                con.Open();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                //if (con.State != ConnectionState.Closed) con.Close();
                result = false;
            }
            return result;
        }
        //Закрыть
        public void Close()
        {
            if (con.State != ConnectionState.Closed) con.Close();
        }
        //Поля
        public List<object> ExecuteQuery(string query, int property)
        {
            //property: 1 - select
            //2 - Update, insert
            if (property == 1)
            {
                List<object> resList = new List<object>();
                try
                {
                    
                    SqlCommand sqlCmd = new SqlCommand(query, con);
                    SqlDataReader reader = sqlCmd.ExecuteReader();

                    while (reader.Read())
                    {
                        List<string> arr = new List<string>();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            arr.Add(reader.GetValue(i).ToString());
                        }
                        resList.Add(arr.ToArray());
                    }
                    //resList.Add(reader.FieldCount);
                    reader.Close();
                    return resList;
                }
                catch (Exception e)
                {
                    MessageBox.Show(e.Message);
                }
                return resList;
            }
            else
            {
                SqlCommand command = con.CreateCommand();
                List<object> res = new List<object>();
                //SqlTransaction transaction;
                try
                {
                    command.CommandText = "update arch_test.dbo.zhora set char='4', integer=4, data=getdate() where char='3'";
                    res.Add(command.ExecuteNonQuery());
                }
                catch (Exception ex)
                {
                    res.Add("-1");
                    MessageBox.Show(ex.Message);
                }
                return res;
            }
        }

        //Выполнение запроса
        public int ExecuteScalarQuery(string query)
        {
            SqlCommand sqlCmd = new SqlCommand(query, con);
            int result;
            try
            {
                result = int.Parse(sqlCmd.ExecuteScalar().ToString());
            }
            catch (Exception)
            {
                result = -1;
            }
            return result;
        }
    }
}
