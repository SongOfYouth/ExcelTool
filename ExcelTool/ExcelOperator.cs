using System;
using System.Text;
using System.Data;
using System.Data.OleDb;

namespace ExcelTool
{
    /**//// <summary>
    /// DataToExcel 的摘要说明。

    /// DataToExcel 的摘要说明。
    /// </summary>
    public class ExcelOperator
    {
        const string ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties=Excel 8.0;";

        public ExcelOperator()
        {
        }

        public string DataTableToExcel(DataSet dstables, string excelPath)
        {
            string connString = string.Format(ConnectionString, excelPath);

            using (OleDbConnection objConn = new OleDbConnection(connString))
            {
                OleDbCommand objCmd = new OleDbCommand();
                objCmd.Connection = objConn;
                objConn.Open();

                foreach (DataTable dt in dstables.Tables)
                {
                    if (dt == null)
                    {
                        return "DataTable不能为空";
                    }
                    int rows = dt.Rows.Count;
                    int cols = dt.Columns.Count;
                    StringBuilder sb;

                    if (rows == 0)
                    {
                        return "没有数据";
                    }

                    sb = new StringBuilder();

                    //生成创建表的脚本
                    sb.Append("CREATE TABLE ");
                    sb.Append(dt.TableName + " ( ");

                    for (int i = 0; i < cols; i++)
                    {
                        if (i < cols - 1)
                            sb.Append(string.Format("{0} varchar,", dt.Columns[i].ColumnName));
                        else
                            sb.Append(string.Format("{0} varchar)", dt.Columns[i].ColumnName));
                    }

                    try
                    {
                        objCmd.CommandText = sb.ToString();
                        objCmd.ExecuteNonQuery();
                    }
                    catch (Exception e)
                    {
                        return "在Excel中创建表失败，错误信息：" + e.Message;
                    }

                    #region 生成插入数据脚本
                    sb.Remove(0, sb.Length);
                    sb.Append("INSERT INTO ");
                    sb.Append(dt.TableName + "( ");

                    for (int i = 0; i < cols; i++)
                    {
                        if (i < cols - 1)
                            sb.Append(dt.Columns[i].ColumnName + ",");
                        else
                            sb.Append(dt.Columns[i].ColumnName + ") values (");
                    }

                    for (int i = 0; i < cols; i++)
                    {
                        if (i < cols - 1)
                            sb.Append("@" + dt.Columns[i].ColumnName + ",");
                        else
                            sb.Append("@" + dt.Columns[i].ColumnName + ")");
                    }
                    #endregion


                    //建立插入动作的Command
                    objCmd.CommandText = sb.ToString();
                    OleDbParameterCollection param = objCmd.Parameters;

                    //遍历DataTable将数据插入新建的Excel文件中
                    foreach (DataRow row in dt.Rows)
                    {
                        for (int i = 0; i < cols; i++)
                        {
                            param.Add(new OleDbParameter(dt.Columns[i].ColumnName, row[i].ToString()));
                        }

                        objCmd.ExecuteNonQuery();
                        param.Clear();
                    }
                }
                objConn.Close();
            }
            return "数据已成功导入Excel";
        }


        public DataSet DataFromExcel(string excelPath)
        {
            DataSet ds = new DataSet();
            string connString = string.Format(ConnectionString, excelPath);

            try
            {
                //实例化一个Oledbconnection类(实现了IDisposable,要using)
                using (OleDbConnection objConn = new OleDbConnection(connString))
                {
                    objConn.Open();
                    using (OleDbCommand ole_cmd = objConn.CreateCommand())
                    {
                        //类似SQL的查询语句这个[Sheet1$对应Excel文件中的一个工作表]
                        ole_cmd.CommandText = "select * from [Sheet1$]";
                        OleDbDataAdapter adapter = new OleDbDataAdapter(ole_cmd);
                        adapter.Fill(ds, "Sheet1");
                    }
                }
                return ds;
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("获取文件数据失败："+ ex.Message);
                return null;
            }
        }
    }
}
