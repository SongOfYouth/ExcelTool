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
using DevExpress.XtraEditors;

namespace ExcelTool
{
    public partial class ExcelOS : Form
    {
        DataSet dsExcelData;
        DataTable dtCreateResult;
        ExcelOperator ExcelHelper = new ExcelOperator();
        string TempletPath = null;
        string OutPath = null;


        public ExcelOS()
        {
            InitializeComponent();

            dtCreateResult = new DataTable();
            dtCreateResult.Columns.AddRange(new DataColumn[] {
                    new DataColumn{ ColumnName = "ID",Caption = "序号"},
                    new DataColumn{ ColumnName = "NAME",Caption = "源文件名称"},
                    new DataColumn{ ColumnName ="STATUS",Caption = "读取状态" },
                    new DataColumn{ ColumnName ="FILE_NAME",Caption = "文件名"} });
        }

        private void ExcelOS_Load(object sender, EventArgs e)
        {
            this.dsExcelData = new DataSet();

            //从配置文件加载路径配置
            while(string.IsNullOrEmpty(TempletPath))
                this.TempletPath = BaseConfig.LoadConfigInfo("FilePath", "TempletPath", "模板路径");
            while(string.IsNullOrEmpty(OutPath))
                this.OutPath = BaseConfig.LoadConfigInfo("FilePath", "OutPath", "输出路径");

            //初始化文件列表
            DataBindingForCombo(cmbPath1, this.TempletPath);

            if (!Directory.Exists(OutPath))
            {
                Directory.CreateDirectory(OutPath);
            }
        }

        private void MenuItemSysTemplet_Click(object sender, EventArgs e)
        {
            this.TempletPath = BaseConfig.SetConfigInfo("FilePath", "TempletPath", "模板路径");
            //初始化文件列表
            DataBindingForCombo(cmbPath1, this.TempletPath);
        }

        private void MenuItemSysOut_Click(object sender, EventArgs e)
        {
            this.OutPath = BaseConfig.SetConfigInfo("FilePath", "OutPath", "输出路径");
        }

        private void DataBindingForCombo(LookUpEdit cmb, string originPath)
        {
            cmb.Properties.DataSource = Common.GetFilePathes(originPath);
            cmb.Properties.ValueMember = "ID";
            cmb.Properties.DisplayMember = "NAME";
            cmb.Properties.BestFit();
            cmb.ItemIndex = 0;
        }

        #region 文件生成及输出
        private void btnCreate_Click(object sender, EventArgs e)
        {
            if (this.txtFileName.Tag == null|| string.IsNullOrEmpty(txtFileName.Text)) return;
            string filepath = this.txtFileName.Tag.ToString();
            if (string.IsNullOrEmpty(filepath)) return;

            DataSet ds =this.ExcelHelper.DataFromExcel(filepath);
            if (ds == null) return;
            DataTable dt = Common.StandardData(ds.Tables[0]);
            DataTable newdt = new DataTable();
            if (txtKey1.Text.Trim() != "")
            {
                newdt = Common.StrRplForDataTable(dt, "描述", txtKey1.Text.Trim(), txtNewKey1.Text.Trim());
                if (txtKey2.Text.Trim() != "")
                {
                    newdt = Common.StrRplForDataTable(newdt, "描述", txtKey2.Text.Trim(), txtNewKey2.Text.Trim());
                }
                if (txtKey3.Text.Trim() != "")
                {
                    newdt = Common.StrRplForDataTable(newdt, "描述", txtKey3.Text.Trim(), txtNewKey3.Text.Trim());
                }
            }
            else
            {
                newdt = dt;
            }
           
            newdt.TableName = txtFileName.Text.Trim();
            while (dsExcelData.Tables.Contains(newdt.TableName))
            {
                //实现末尾追加序号
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
            string outPath = VerifyPath(txtFileName.Text.Trim());
            foreach (DataTable t in dsExcelData.Tables)
            {
                DataSet ds = new DataSet();
                DataTable dt = t.Copy();
                ds.Tables.Add(dt);
                this.ExcelHelper.DataTableToExcel(ds, outPath);
            }
            this.OpenFile(outPath);

        }

        private void bbiOutputBySheet_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
        {
            string outPath = VerifyPath(txtFileName.Text.Trim());
            this.ExcelHelper.DataTableToExcel(dsExcelData, outPath);
            this.OpenFile(outPath);
        }

        private void btnOutput_Click(object sender, EventArgs e)
        {
            List<string> lsTableName = new List<string>();
            foreach (int i in gridvData.GetSelectedRows())
            {
                lsTableName.Add(gridvData.GetDataRow(i)["FILE_NAME"].ToString());
            }
            if(lsTableName.Count ==0)return;
            DataSet ds = new DataSet();
            DataTable dt = dsExcelData.Tables[0].Clone();
            foreach (DataTable  t in dsExcelData.Tables)
            {
                if (lsTableName.Contains(t.TableName))
                {
                    foreach (DataRow dr in t.Rows)
                    {
                        dt.Rows.Add(dr.ItemArray);
                    }
                }
            }
            dt.TableName = "合并输出";
            ds.Tables.Add(dt);
            string outPath = VerifyPath(txtFileName.Text.Trim());
            this.ExcelHelper.DataTableToExcel(ds, outPath);
            this.OpenFile(outPath);
        }
        private void btnClear_Click(object sender, EventArgs e)
        {
            if (gridvData.DataRowCount == 0) return;
            List<DataRow> rows = new List<DataRow>();
            gridvData.GetSelectedRows().AsParallel().ForAll<int>(r => rows.Add(gridvData.GetDataRow(r)));
            foreach (DataRow dr in rows)
            {
                dsExcelData.Tables.Remove(dr["FILE_NAME"].ToString());
                dtCreateResult.Rows.Remove(dr);
            }

        }
        private void txtFileName_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                this.btnCreate_Click(null, null);
            }
        }
        #endregion

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
        /// <summary>在资源管理器打开指定文件<param name="fileFullName"></param>
        private void OpenFile(string fileFullName)
        {
            System.Diagnostics.ProcessStartInfo psi = new System.Diagnostics.ProcessStartInfo("Explorer.exe");
            psi.Arguments = "/e,/select," + fileFullName;
            System.Diagnostics.Process.Start(psi);
        }

        #region 路径加载
        private void cmbPath1_EditValueChanged(object sender, EventArgs e)
        {
            if(cmbPath1.EditValue == null)
            {
                this.cmbPath2.Properties.DataSource = null;
                return;
            }
            if (cmbPath1.GetColumnValue("TYPE").Equals("File")) return;
            DataBindingForCombo(this.cmbPath2, cmbPath1.EditValue.ToString());
        }

        private void cmbPath2_EditValueChanged(object sender, EventArgs e)
        {
            if (cmbPath2.EditValue == null)
            {
                this.cmbPath3.Visible = false;
                this.cmbPath4.Visible = false;
                this.cmbPath3.Properties.DataSource = null;
                return;
            }
            if (cmbPath2.GetColumnValue("TYPE").Equals("File"))
            {
                SetFileInfo(cmbPath2.EditValue.ToString());
                this.cmbPath3.Visible = false;
                this.cmbPath4.Visible = false;
                return;
            }
            this.cmbPath3.Visible = true;
                DataBindingForCombo(this.cmbPath3, cmbPath2.EditValue.ToString());
        }

        private void cmbPath3_EditValueChanged(object sender, EventArgs e)
        {
            if (cmbPath3.EditValue == null)
            {
                this.cmbPath4.Properties.DataSource = null;
                this.cmbPath4.Visible = false;
                return;
            }
            if (cmbPath3.GetColumnValue("TYPE").Equals("File"))
            {
                SetFileInfo(cmbPath3.EditValue.ToString());
                this.cmbPath4.Visible = false;
                return;
            }
            this.cmbPath4.Visible = true;
            DataBindingForCombo(this.cmbPath4, cmbPath3.EditValue.ToString());
        }

        private void cmbFile_EditValueChanged(object sender, EventArgs e)
        {
            if (cmbPath4.EditValue == null) return;
            SetFileInfo(cmbPath4.EditValue.ToString());
        }
        private void SetFileInfo(string filePath)
        {
            this.txtFileName.Text = filePath.Split('\\').Last<string>().Split('.')[1].Trim("（V1）".ToCharArray());
            this.txtFileName.Tag = filePath;
        }
        #endregion

    }
}
