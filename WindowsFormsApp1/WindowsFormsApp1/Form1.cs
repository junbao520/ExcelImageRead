using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        //读取Excel 图片
        private void button1_Click(object sender, EventArgs e)
        {
            NPOI.SS.UserModel.ISheet sheet = null;//工作表
            var fs = new FileStream(txtExcelName.Text, FileMode.Open, FileAccess.Read);
            var workbook = new XSSFWorkbook(fs);
            for(int i = 0; i <Convert.ToInt32( txtNumber.Text); i++)
            {
                sheet = workbook.GetSheetAt(i);
               
                var list = sheet.GetAllPictureInfos();
                var data = ExcelToDataTable(sheet);
                int index = 0;

                var FilePath = DateTime.Now.ToString("yyyyMMdd") + sheet.SheetName;
                if (!System.IO.Directory.Exists(FilePath))
                {
                    //创建pic文件夹
                    System.IO.Directory.CreateDirectory(FilePath);
                }
                else
                {
                    System.IO.Directory.Delete(FilePath, true);
                    System.IO.Directory.CreateDirectory(FilePath);
                }

                foreach (var item in list)
                {
                   
                    try
                    {
                        index++;
                        int row = item.MinRow;
                        var info = data.Rows[row - 1];
                        var name = ""; ;
                        for(var j = 0; j < Convert.ToInt32(txtColCount.Text); j++)
                        {
                            name += info.ItemArray[j].ToString();
                        }
                        name+= "_" + index.ToString() + ".png";
                        //var name = info.ItemArray[0].ToString() + info.ItemArray[1].ToString() + info.ItemArray[2].ToString() + info.ItemArray[3].ToString() + info.ItemArray[4].ToString() + "_" + index.ToString() + ".png";
                        //单品序号+品牌+渠道+达人账号这样命名
                        name = name.Replace("/", "");
                        writePic(item.PictureData, name, FilePath);
                    }
                    catch (Exception)
                    {

                        continue;
                    }

                }
                //重名的怎么处理
            }




            MessageBox.Show("读取完成");
        

        }
        public void writePic(byte[] data,string name,string FilePath)
        {
          
            MemoryStream ms = new MemoryStream(data);
          
            FileStream fs = new FileStream(string.Format("{0}/{1}", FilePath, name),FileMode.Create);
            ms.WriteTo(fs);
            ms.Close();
            fs.Close();
        }

        /// <summary>
        /// 将Excel导入DataTable
        /// </summary>
        /// <param name="filepath">导入的文件路径（包括文件名）</param>
        /// <param name="sheetname">工作表名称</param>
        /// <param name="isFirstRowColumn">第一行是否是DataTable的列名</param>
        /// <returns>DataTable</returns>
        public  DataTable ExcelToDataTable(ISheet sheet)
        {
      
            DataTable data = new DataTable();

            var startrow = 0;
            {
                try
                {
                    
                    if (sheet != null)
                    {
                        IRow firstrow = sheet.GetRow(0);
                        int cellCount = firstrow.LastCellNum; //行最后一个cell的编号 即总的列数
                        if (true)
                        {
                            for (int i = firstrow.FirstCellNum; i < cellCount; i++)
                            {
                                ICell cell = firstrow.GetCell(i);
                                if (cell != null)
                                {
                                    string cellvalue = cell.StringCellValue;
                                    if (cellvalue != null)
                                    {
                                        DataColumn column = new DataColumn(cellvalue);
                                        data.Columns.Add(column);
                                    }
                                }
                            }
                            startrow = sheet.FirstRowNum + 1;
                        }
                        else
                        {
                            startrow = sheet.FirstRowNum;
                        }
                        //读数据行
                        int rowcount = sheet.LastRowNum;
                        for (int i = startrow; i < rowcount; i++)
                        {
                            IRow row = sheet.GetRow(i);
                            if (row == null)
                            {
                                continue; //没有数据的行默认是null
                            }
                            DataRow datarow = data.NewRow();//具有相同架构的行
                            for (int j = row.FirstCellNum; j < cellCount; j++)
                            {
                                if (row.GetCell(j) != null)
                                {
                                    datarow[j] = row.GetCell(j).ToString();
                                }
                            }
                            data.Rows.Add(datarow);
                        }
                    }
                    return data;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception: " + ex.Message);
                    return null;
                }
            }
        }
    }
}
