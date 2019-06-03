using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net;
using System.IO;
using System.Threading;

namespace DFGZ
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        int i = 2;

        public MainWindow()
        {
            InitializeComponent();
        }

        /// <summary>
        /// 获取待解析文件路径路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void getFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.Filter = "Excel文件|*.xlsx;*.xls";
            if (dialog.ShowDialog() == true)
            {
                filePathText.Text = dialog.FileName;
                TextBoxView("获取待解析文件");
                //filePath = dialog.FileName;
            }
        }

        /// <summary>
        /// 获取解析后保存文件的路径
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void save_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog fbd = new FolderBrowserDialog();
            fbd.ShowDialog();
            if (fbd.SelectedPath != "") 
            {
                savePathText.Text = fbd.SelectedPath;
                TextBoxView("获取解析后保存文件的路径");
                //saveFilePath = fbd.SelectedPath;
            }
           
        }

        private void analysisFile_Click(object sender, RoutedEventArgs e)
        {
            string excelFile = filePathText.Text;//需要解析的Excel文件名
            string saveFilePath = savePathText.Text;//解析结果保存路径
            string saveTxtName = saveFileName.Text;//Txt文件名
            if (excelFile.Equals("待解析文件路径"))
            {
                TextBoxView("请选择待解析文件");
                return;
            }
            if (saveFilePath.Equals("解析后文件保存路径"))
            {
                TextBoxView("请选择解析后文件保存路径");
                return;
            }
            if (saveTxtName.Equals(""))
            {
                TextBoxView("请输入解析后文件名");
                return;
            }
            
            //System.Windows.Forms.MessageBox.Show("开始解析文件！", "提示对话框", MessageBoxButtons.OK, MessageBoxIcon.Information);
            TextBoxView("开始解析文件,请不要关闭Excel表格");


            GetExcelDataAndWriteTxt(excelFile, saveFilePath, saveTxtName);          
        }

        public void GetExcelDataAndWriteTxt(string excelFile, string saveFilePath, string saveTxtName) 
        {
            
                
                int renshuCount = 0;//符合标准的人数

                Excel.Application xApp = new Excel.Application();
                xApp.Visible = true;

                Excel.Workbook xBook = xApp.Workbooks._Open(excelFile, Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value);

                Excel.Worksheet xSheet = (Excel.Worksheet)xBook.Sheets[1];

                int RowCount = xSheet.UsedRange.Cells.Rows.Count; //得到行数

                FileStream fs = new FileStream(saveFilePath + "/" + saveTxtName + ".txt", FileMode.OpenOrCreate, FileAccess.Write);//创建文本文件
                StreamWriter sw = new StreamWriter(fs, Encoding.Default);
                //sw.Write("0001|");
                double zongjine = 0.00; //总金额
                string xh = "";//序号
                string zh = "";//账号
                string xm = "";//姓名
                double je = 0.00;//金额

                int count = 0;
                try
                {
                    for (i = 2; i <= RowCount; i++)
                    {

                        Excel.Range rng4 = (Excel.Range)xSheet.Cells[i, 4];//金额
                        if (rng4.Value2 != null)
                        {
                            je = double.Parse(rng4.Value2.ToString());
                            string jeS = je.ToString("f2");
                            //检查金额是否为0，若为0，直接continue               
                            if (je != 0.00)
                            {
                                count++;
                                //Excel.Range rng1 = (Excel.Range)xSheet.Cells[i, 1];//序号
                                //xh = (count + "").PadLeft(5, '0');//将序号补零
                                xh = (count + "");
                                Excel.Range rng2 = (Excel.Range)xSheet.Cells[i, 2];//账号
                                //检查账号是否有空格
                                zh = rng2.Value2.ToString();
                                zh = zh.Replace(" ", "");

                                Excel.Range rng3 = (Excel.Range)xSheet.Cells[i, 3];//姓名
                                //检查姓名是否含有空格以及（和(等内容
                                xm = rng3.Value2.ToString();
                                xm = xm.Replace(" ", "");
                                xm = xm.Replace("　", "");
                                xm = System.Text.RegularExpressions.Regex.Replace(xm, @"[\(（][\s\S]*[\)）]", "");

                                zongjine = zongjine + je;//计算金额

                                //开始写入值
                                //if (i != 2)
                                //    sw.Write("\r\n" + xh + "|" + zh + "|" + xm + "|" + jeS + "|");
                                //else
                                //    sw.Write(xh + "|" + zh + "|" + xm + "|" + jeS + "|");
                                sw.Write("\r\n" + xh + "|" + zh + "|" + xm + "|" + jeS + "|");

                                //人数自加
                                renshuCount++;


                            }
                            else
                            {
                                continue;
                            }
                        }
                        else
                        {
                            continue;
                        }

                    }
                    jinE1.Text = zongjine.ToString("f2");//显示金额
                    renshu.Text = renshuCount.ToString();//显示代发人数


                    System.Windows.Forms.MessageBox.Show("文件解析结束！", "提示对话框", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    TextBoxView("文件解析结束");
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show("文件解析错误！", "提示对话框", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    TextBoxView("文件解析错误");
                    TextBoxView(ex.Message);
                }
                finally
                {
                    sw.Close();
                    fs.Close();
                }
                      
        }

        private void TextBoxView(string context)
        {
            OutTxtBox.AppendText("\r" + context + "。。。");
            OutTxtBox.ScrollToEnd();
        }
    }
}
