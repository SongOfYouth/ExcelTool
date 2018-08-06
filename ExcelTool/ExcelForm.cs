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

namespace ExcelTool
{
    public partial class ExcelOS : Form
    {
        DataSet dsExcelData;
        DataTable dtCreateResult;
        ExcelOperator ExcelHelper = new ExcelOperator();
        readonly string MainPath = @".\模块库";
        readonly string OutPath = @".\Excels";

        #region 属性
        private List<string> lsPathStr
        {
            get
            {
                if (rdgAttribute1.EditValue == null || rdgAttribute2.EditValue == null || rdgAttribute3.EditValue == null) return null;
                List<string> list = new List<string>();
                list.Add(rdgAttribute1.EditValue.ToString() + "*");
                list.Add(rdgAttribute2.EditValue.ToString() + "*");
                list.Add(rdgAttribute3.EditValue.ToString() + "*");
                return list;
            }
        }

        #endregion


        public ExcelOS()
        {
            InitializeComponent();

            dtCreateResult = new DataTable();
            dtCreateResult.Columns.AddRange(new DataColumn[] {
                    new DataColumn{ ColumnName = "ID",Caption = "序号"},
                    new DataColumn{ ColumnName = "NAME",Caption = "源文件名称"},
                    new DataColumn{ ColumnName ="STATUS",Caption = "读取状态" },
                    new DataColumn{ ColumnName ="FILENAME",Caption = "文件名"} });
        }

        private void ExcelOS_Load(object sender, EventArgs e)
        {
            this.dsExcelData = new DataSet();
            if (!Directory.Exists(MainPath))
            {
                Directory.CreateDirectory(MainPath);
            }
            if (!Directory.Exists(OutPath))
            {
                Directory.CreateDirectory(OutPath);
            }
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            if (this.txtFileName.Tag == null|| string.IsNullOrEmpty(txtFileName.Text)) return;
            string filepath = this.txtFileName.Tag.ToString();
            if (string.IsNullOrEmpty(filepath)) return;

            DataSet ds =this.ExcelHelper.DataFromExcel(filepath);
            if (ds == null) return;
            DataTable dt = Common.StandardData(ds.Tables[0]);
            DataTable newdt = new DataTable();
            if (txtKey1.Text.Trim() != "" && txtNewKey1.Text.Trim() != "")
            {
                newdt = Common.StrRplForDataTable(dt, "描述", txtKey1.Text.Trim(), txtNewKey1.Text.Trim());
                if (txtKey2.Text.Trim() != "" && txtNewKey2.Text.Trim() != "")
                {
                    newdt = Common.StrRplForDataTable(newdt, "描述", txtKey2.Text.Trim(), txtNewKey2.Text.Trim());
                }
                if (txtKey3.Text.Trim() != "" && txtNewKey3.Text.Trim() != "")
                {
                    newdt = Common.StrRplForDataTable(newdt, "描述", txtKey3.Text.Trim(), txtNewKey3.Text.Trim());
                }
            }
           
            newdt.TableName = txtFileName.Text.Trim();
            while (dsExcelData.Tables.Contains(newdt.TableName))
            {
                if(newdt.TableName.Last<char>()<58&& newdt.TableName.Last<char>() >47)
                {
                    newdt.TableName = newdt.TableName.Substring(0, newdt.TableName.Length - 1) + (Char)(newdt.TableName.Last<char>() + 1);
                }
                else
                {
                    newdt.TableName += "1";
                }
            }
            dsExcelData.Tables.Add(newdt);

            dtCreateResult.Rows.Add(new object[] {dtCreateResult.Rows.Count+1,filepath.Split('\\').Last<string>(),"Success", newdt.TableName });
            this.gridcData.DataSource = dtCreateResult;
            this.gridvData.BestFitColumns();
        }

        private void bbiOutputSingle_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            foreach (DataTable t in dsExcelData.Tables)
            {
                DataSet ds = new DataSet();
                DataTable dt = t.Copy();
                ds.Tables.Add(dt);
                this.ExcelHelper.DataTableToExcel(ds, VerifyPath(dt.TableName));
            }
        }

        private void bbiOutputBySheet_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            this.ExcelHelper.DataTableToExcel(dsExcelData, VerifyPath(txtFileName.Text.Trim()));
        }

        private void btnOutput_Click(object sender, EventArgs e)
        {
            DataSet ds = new DataSet();
            DataTable dt = dsExcelData.Tables[0].Clone();
            dt.TableName = "合并输出";
            foreach (DataTable  t in dsExcelData.Tables)
            {
                foreach (DataRow dr in t.Rows)
                {
                    dt.Rows.Add(dr.ItemArray);
                }
            }
            ds.Tables.Add(dt);
            this.ExcelHelper.DataTableToExcel(ds, VerifyPath(txtFileName.Text.Trim()));
        }
        
        private string VerifyPath(string fileName)
        {
            try
            {
                string file = OutPath + "\\" + fileName + ".xls";
                if (File.Exists(file))
                {
                    File.Delete(file);
                }
                return file;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }

        private void rdgAttribute3_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(MainPath) || lsPathStr == null) return;
            string filepath = Common.GetFilePath(MainPath, lsPathStr);
            if (string.IsNullOrEmpty(filepath)) return;

            this.txtFileName.Text = filepath.Split('\\').Last<string>().Split('.')[1].Trim("（V1）".ToCharArray());
            this.txtFileName.Tag = filepath;
        }

        private void rdgAttribute3_MouseClick(object sender, MouseEventArgs e)
        {
            this.rdgAttribute3_SelectedIndexChanged(null, null);
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            dtCreateResult.Clear();
            this.dsExcelData.Tables.Clear();
            this.gridcData.DataSource = null;
        }

        private void txtFileName_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                this.btnCreate_Click(null, null);
            }
        }
    }
}
