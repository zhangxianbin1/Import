using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Import
{
    public partial class Form1 : Form
    {
        private string filePath = "";
        private DataTable dataTable = null;
        private string sheetName = "sheet1";
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                dialog.Filter = "excel文件(*.xls,*.xlsx)|*.xls;*xlsx";
                dialog.ShowDialog();
                filePath = dialog.FileName;
                if (filePath != "")
                {
                    dataTable = ExcelToDataTable(filePath, sheetName);
                }
                else
                {
                    MessageBox.Show("请选择导入文件");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        private DataTable ExcelToDataTable(string filePath,string sheetName)
        {
            FileStream fs = null;
            IWorkbook workbook = null;
            ISheet sheet = null;
            DataTable resultDataTable = new DataTable();
            fs = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            if (filePath.IndexOf(".xlsx") > 0)//2007版本
            {
                workbook = new XSSFWorkbook(fs);
            }
            else if (filePath.IndexOf(".xls") > 0)//2003版本
            {
                workbook = new HSSFWorkbook(fs);
            }
            if (string.IsNullOrEmpty(sheetName))
            {
                sheet = workbook.GetSheetAt(0);
            }
            else
            {
                sheet = workbook.GetSheet(sheetName);
                if (sheet == null)
                {
                    sheet = workbook.GetSheetAt(0);
                }
            }
            if (sheet != null)
            {
                IRow firstRow = sheet.GetRow(0);
                int rowCount = sheet.LastRowNum;
                int cellCount = firstRow.LastCellNum;
                for (int i = firstRow.FirstCellNum; i < cellCount; ++i)
                {
                    ICell cell = firstRow.GetCell(i);
                    if (cell != null)
                    {
                        string cellValue = cell.StringCellValue.ToString().Trim();
                        if (cellValue != null)
                        {
                            DataColumn column = new DataColumn(cellValue);
                            resultDataTable.Columns.Add(column);
                        }
                    }
                }
                for (int i = 0; i <= rowCount; ++i)
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        if (row.Cells.Count != 0)
                        {
                            DataRow dataRow = resultDataTable.NewRow();
                            for (int j = row.FirstCellNum; j < cellCount; ++j)
                            {
                                if (row.GetCell(j) != null)
                                {
                                    dataRow[j] = row.GetCell(j).ToString().Trim();
                                }
                            }
                            resultDataTable.Rows.Add(dataRow);
                        }
                    }
                }

            }
            return resultDataTable;
        }
    }
}
