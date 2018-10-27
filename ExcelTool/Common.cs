using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace ExcelTool
{
    public static class Common
    {
        #region 规范数据
        //规范表格式:去除无效数据
        public static DataTable StandardTable(DataTable Tsource, int StartIndex = 1)
        {
            if (Tsource == null) return null;
            DataTable table = new DataTable();
            //将第二行作为列头
            for (int c = 0; c < Tsource.Columns.Count; c++)
            {
                if (string.IsNullOrEmpty(Tsource.Rows[StartIndex][c].ToString())) continue;
                table.Columns.Add(Tsource.Rows[StartIndex][c].ToString(), typeof(string));
            }
            for (int r = 2; r < Tsource.Rows.Count; r++)
            {
                if (Tsource.Rows[r].HasErrors) continue;
                object[] rd = new object[table.Columns.Count];
                for (int i = 0; i < table.Columns.Count; i++)
                {
                    rd[i] = Tsource.Rows[r].ItemArray[i];
                }
                table.Rows.Add(rd);
            }
            return table;
        }
        //规范合并单元格数据
        public static DataTable StandardData(DataTable Tsource)
        {
            if (Tsource == null) return null;
            DataTable table = Tsource.Copy();
            for (int c = 0; c < Tsource.Columns.Count; c++)
            {
                string cell = null;
                for (int r = 0; r < Tsource.Rows.Count; r++)
                {
                    if (!string.IsNullOrEmpty(Tsource.Rows[r][c].ToString()))
                    {
                        cell = Tsource.Rows[r][c].ToString();
                    }
                    table.Rows[r][c] = cell;
                }
            }
            return table;
        }
        #endregion

        #region 路径函数

        //获取路径
        public static string GetPath(string path, string instr)
        {
            if (string.IsNullOrEmpty(instr) || string.IsNullOrEmpty(path)) return path;
            DirectoryInfo directory = new DirectoryInfo(path);
            DirectoryInfo[] infos = directory.GetDirectories(instr);
            if (infos.Count() > 1 || infos.Count() == 0)
            {
                var files = directory.GetFiles(instr);
                if (files.Count() > 1 || files.Count() == 0)
                {
                    MessageBox.Show("当前" + (files.Count() > 1 ? "存在多个" : "不存在") + "匹配，请核查路径【" + path + "】");
                }
                return files[0].FullName;
            }
            return infos[0].FullName;
        }

        /// <summary>逐级获取关联的最终文件路径</summary>
        /// <param name="strStartPath"></param>
        /// <param name="lsSearchstr"></param>
        /// <returns></returns>
        public static string GetFilePath(string strStartPath, List<string> lsSearchstr)
        {
            if (string.IsNullOrEmpty(strStartPath) || lsSearchstr.Count == 0) return null;
            try
            {
                string filePath = strStartPath;
                foreach (string str in lsSearchstr)
                {
                    filePath = GetPath(filePath, str);
                }
                return filePath;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }
        /// <summary> 获取文件列表</summary>
        /// <param name="ds"></param>
        /// <param name="originPath"></param>
        /// <returns></returns>
        public static DataTable GetFilePathes(string originPath)
        {
            if (string.IsNullOrEmpty(originPath)) return null;
            DataTable dt = new DataTable();
            dt.Columns.AddRange(new DataColumn[] {new DataColumn("ID",typeof(string)),
                new DataColumn("NAME", typeof(string)),
            new DataColumn("TYPE", typeof(string))});
            DirectoryInfo directory = new DirectoryInfo(originPath);
            DirectoryInfo[] infos = directory.GetDirectories();
            foreach (var dir in infos)
            {
                if (dir.Name.StartsWith(".")) continue; //排除.开头的隐藏文件
                dt.Rows.Add(new object[] { dir.FullName,dir.Name,"Path"});
            }
            foreach (var file in directory.GetFiles())
            {
                dt.Rows.Add(new object[] { file.FullName, file.Name,"File"});
            }
            return dt;
        }
        #endregion

        #region 字符操作


        /// <summary>
        /// 替换Datatable指定列中的字符串
        /// </summary>
        public static DataTable StrRplForDataTable(DataTable data, string fieldName, string sourceStr, string newStr)
        {
            if (string.IsNullOrEmpty(sourceStr) ||
                string.IsNullOrEmpty(fieldName) ||
                data == null) return data;
            DataTable dt = data.Copy();
            try
            {
                foreach (DataRow dr in dt.Rows)
                {
                    dr[fieldName] = (dr[fieldName] ?? "").ToString().Replace(sourceStr, newStr);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
            return dt;
        }
        public static DataTable StrRplForDataTable(DataTable data, string fieldName, string sourceStr1, string newStr1, string sourceStr2, string newStr2)
        {
            if (string.IsNullOrEmpty(sourceStr1) ||
                string.IsNullOrEmpty(sourceStr2) ||
                string.IsNullOrEmpty(fieldName) ||
                data == null) return data;
            DataTable dt = data.Copy();
            try
            {
                foreach (DataRow dr in dt.Rows)
                {
                    dr[fieldName] = (dr[fieldName] ?? "").ToString().Replace(sourceStr1, newStr1);
                    dr[fieldName] = (dr[fieldName] ?? "").ToString().Replace(sourceStr2, newStr2);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
            return dt;
        }

        #endregion

        #region 数据集操作
        /// <summary>
        /// Datatable 对象之间的操作（交集、并集、差集）
        /// </summary>
        /// <param name="opType">“union”，“intersect”，“except”</param>
        /// <returns></returns>
        public static DataTable TablesTo(DataTable table1, DataTable table2, string opType)
        {
            if (opType == "union")
            {
                IEnumerable<DataRow> union = table1.AsEnumerable().Union(table2.AsEnumerable(), DataRowComparer.Default);
                return union.CopyToDataTable();
            }
            else if (opType == "intersect")
            {
                IEnumerable<DataRow> union = table1.AsEnumerable().Intersect(table2.AsEnumerable(), DataRowComparer.Default);
                return union.CopyToDataTable();
            }
            else
            {
                IEnumerable<DataRow> union = table1.AsEnumerable().Except(table2.AsEnumerable(), DataRowComparer.Default);
                return union.CopyToDataTable();
            }
        }

        #endregion
    }
}
