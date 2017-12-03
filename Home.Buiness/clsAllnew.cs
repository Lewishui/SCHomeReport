using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Home.DB;
using Order.Common;
using Word = Microsoft.Office.Interop.Word;

namespace Home.Buiness
{
    public class clsAllnew
    {
        public BackgroundWorker bgWorker1;
        public ToolStripProgressBar pbStatus { get; set; }
        public ToolStripStatusLabel tsStatusLabel1 { get; set; }
        public WorkerArgument arg;
        public BackgroundWorker backgroundWorker1;

        public log4net.ILog ProcessLogger { get; set; }
        public log4net.ILog ExceptionLogger { get; set; }
        List<clsjinchushuju_budongchaninfo> zhuzaidiya_Result;

        public List<clsjinchushuju_budongchaninfo> Buiness_Bankcharge(ref BackgroundWorker bgWorker)
        {
            try
            {
                string fin = "";

                bgWorker1 = bgWorker;
                string ZFCEPath = AppDomain.CurrentDomain.BaseDirectory + "Resources";
                List<clsjinchushuju_budongchaninfo> R2r_bankResult = new List<clsjinchushuju_budongchaninfo>();
                List<string> Alist = GetBy_CategoryReportFileName(ZFCEPath);
                for (int i = 0; i < Alist.Count; i++)
                {
                    if (Alist[i].Contains("住宅抵押模版"))
                        R2r_bankResult = Getzhuzaidiya(ZFCEPath + "\\" + Alist[i]);
                }
                ReplaceToExcel(R2r_bankResult[0]);

                return null;
            }
            catch (Exception ex)
            {

                throw;
            }

        }
        private List<string> GetBy_CategoryReportFileName(string dirPath)
        {

            List<string> FileNameList = new List<string>();
            ArrayList list = new ArrayList();

            if (Directory.Exists(dirPath))
            {
                list.AddRange(Directory.GetFiles(dirPath));
            }
            if (list.Count > 0)
            {
                foreach (object item in list)
                {
                    if (!item.ToString().Contains("~$"))
                        FileNameList.Add(item.ToString().Replace(dirPath + "\\", ""));
                }
            }

            return FileNameList;
        }
        //住宅抵押模版（房产证、土地证）读取
        public List<clsjinchushuju_budongchaninfo> Getzhuzaidiya(string Alist)
        {

            List<clsjinchushuju_budongchaninfo> MAPPINGResult = new List<clsjinchushuju_budongchaninfo>();
            try
            {
                List<clsjinchushuju_budongchaninfo> WANGYINResult = new List<clsjinchushuju_budongchaninfo>();
                System.Globalization.CultureInfo CurrentCI = System.Threading.Thread.CurrentThread.CurrentCulture;
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
                Microsoft.Office.Interop.Excel.Application excelApp;
                {
                    string path = Alist;
                    excelApp = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook analyWK = excelApp.Workbooks.Open(path, Type.Missing, Type.Missing, Type.Missing,
                        "htc", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing, Type.Missing, Type.Missing);

                    Microsoft.Office.Interop.Excel.Worksheet WS = (Microsoft.Office.Interop.Excel.Worksheet)analyWK.Worksheets["基础数据（不动产）"];
                    Microsoft.Office.Interop.Excel.Range rng;
                    rng = WS.Range[WS.Cells[1, 1], WS.Cells[WS.UsedRange.Rows.Count, 16]];
                    int rowCount = WS.UsedRange.Rows.Count - 1;
                    object[,] o = new object[1, 1];
                    o = (object[,])rng.Value2;
                    int wscount = analyWK.Worksheets.Count;
                    clsCommHelp.CloseExcel(excelApp, analyWK);


                    //for (int i = 2; i <= rowCount; i++)
                    {
                        clsjinchushuju_budongchaninfo temp = new clsjinchushuju_budongchaninfo();

                        #region 基础信息

                        temp.quanliren = "";
                        if (o[8, 2] != null)
                            temp.quanliren = o[8, 2].ToString().Trim();

                        temp.zuoluo = "";
                        if (o[5, 2] != null)
                            temp.zuoluo = o[5, 2].ToString().Trim();


                        temp.gujiaweituoren = "";
                        if (o[8, 4] != null)
                            temp.gujiaweituoren = o[8, 4].ToString().Trim();

                        //卖场代码

                        temp.gujiashi1 = "";
                        if (o[42, 2] != null)
                            temp.gujiashi1 = o[42, 2].ToString().Trim();

                        temp.zhucehao1 = "";
                        if (o[42, 3] != null)
                            temp.zhucehao1 = o[42, 3].ToString().Trim();

                        temp.gujiashi2 = "";
                        if (o[43, 2] != null)
                            temp.gujiashi2 = o[43, 2].ToString().Trim();
                        temp.zhucehao2 = "";
                        if (o[43, 3] != null)
                            temp.zhucehao2 = o[43, 3].ToString().Trim();

                        temp.chujubaogaoriqi = "";
                        if (o[44, 2] != null)
                            temp.chujubaogaoriqi = o[44, 2].ToString().Trim();

                        //卖场名称
                        temp.baoggaobianhao = "";
                        if (o[3, 2] != null)
                            temp.baoggaobianhao = o[3, 2].ToString().Trim();

                        temp.baoggaobianhao2 = "";
                        if (o[3, 3] != null)
                            temp.baoggaobianhao2 = o[3, 3].ToString().Trim();

                        temp.Danfangwuyongtu = "";
                        if (o[13, 5] != null)
                            temp.Danfangwuyongtu = o[13, 5].ToString().Trim();
                        

                        temp.Input_Date = DateTime.Now.ToString("yyyy/MM/dd");

                        #endregion
                        MAPPINGResult.Add(temp);
                    }

                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: 01032" + ex);
                return null;

                throw;
            }
            return MAPPINGResult;

        }
        #region 替换word


        /// <summary>
        /// 替换word中的文本，并导出word
        /// </summary>
        protected void ReplaceToExcel(clsjinchushuju_budongchaninfo temp)
        {
            string ZFCEPath = AppDomain.CurrentDomain.BaseDirectory + "Results";
               
            Word.Application app = null;
            Word.Document doc = null;
            //将要导出的新word文件名
            string newFile = DateTime.Now.ToString("yyyyMMddHHmmssss") + ".doc";
          //  string physicNewFile = Server.MapPath(DateTime.Now.ToString("yyyyMMddHHmmssss") + ".doc");
          string  physicNewFile = AppDomain.CurrentDomain.BaseDirectory + "Results\\" + DateTime.Now.ToString("yyyyMMddHHmmssss") + ".doc";
               
            try
            {
                app = new Word.Application();//创建word应用程序
               
                //object fileName = Server.MapPath("template.doc");//模板文件
                object fileName = AppDomain.CurrentDomain.BaseDirectory + "Resources\\1\\" + "办公抵押模版（不动产） .doc";
             
                //打开模板文件
                object oMissing = System.Reflection.Missing.Value;
                doc = app.Documents.Open(ref fileName,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing,
                ref oMissing, ref oMissing, ref oMissing, ref oMissing, ref oMissing);

                //构造数据
                Dictionary<string, string> datas = new Dictionary<string, string>();
                datas.Add("{name}", temp.quanliren);
                datas.Add("{坐落}", temp.zuoluo);
                datas.Add("{E13}", temp.Danfangwuyongtu);
                datas.Add("{D8}", temp.gujiaweituoren);
                datas.Add("{B42}", temp.gujiashi1);
                datas.Add("{C42}", temp.zhucehao1);
                datas.Add("{B43}", temp.gujiashi2);
                datas.Add("{C43}", temp.zhucehao2);
                datas.Add("{C44}", temp.chujubaogaoriqi);
                datas.Add("{B3}", temp.baoggaobianhao);
                datas.Add("{C3}", temp.baoggaobianhao2);


                object replace = Word.WdReplace.wdReplaceAll;
                foreach (var item in datas)
                {
                    app.Selection.Find.Replacement.ClearFormatting();
                    app.Selection.Find.ClearFormatting();
                    app.Selection.Find.Text = item.Key;//需要被替换的文本
                    app.Selection.Find.Replacement.Text = item.Value;//替换文本 

                    //执行替换操作
                    app.Selection.Find.Execute(
                    ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing, ref oMissing,
                    ref oMissing, ref replace,
                    ref oMissing, ref oMissing,
                    ref oMissing, ref oMissing);
                }

                //对替换好的word模板另存为一个新的word文档
                doc.SaveAs(physicNewFile,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing, oMissing,
                oMissing, oMissing, oMissing, oMissing, oMissing, oMissing);

                //准备导出word
                //Response.Clear();
                //Response.Buffer = true;
                //Response.Charset = "utf-8";
                //Response.AddHeader("Content-Disposition", "attachment;filename=" + DateTime.Now.ToString("yyyyMMddHHmmssss") + ".doc");
                //Response.ContentEncoding = System.Text.Encoding.GetEncoding("utf-8");
                //Response.ContentType = "application/ms-word";
                //Response.End();
            }
            catch (System.Threading.ThreadAbortException ex)
            {
                //这边为了捕获Response.End引起的异常
            }
            catch (Exception ex)
            {

            }
            finally
            {
                if (doc != null)
                {
                    doc.Close();//关闭word文档
                }
                if (app != null)
                {
                    app.Quit();//退出word应用程序
                }
                //如果文件存在则输出到客户端
                if (File.Exists(physicNewFile))
                {
                  //  Response.WriteFile(physicNewFile);
                }
            }
        }

        #endregion

    }
}
