using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Home.Buiness;
using Home.DB;
using Order.Common;
using WeifenLuo.WinFormsUI.Docking;

namespace SCHomeReport
{
    public partial class frmMain : DockContent
    {
        // 后台执行控件
        private BackgroundWorker bgWorker;
        // 消息显示窗体
        private frmMessageShow frmMessageShow;
        // 后台操作是否正常完成
        private bool blnBackGroundWorkIsOK = false;
        //后加的后台属性显
        private bool backGroundRunResult;
        List<clsjinchushuju_budongchaninfo> zhuzaidiya_Result;

        public frmMain(string A ,string B)
        {
            InitializeComponent();
        }

        private void page1ToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }
        private void InitialBackGroundWorker()
        {
            bgWorker = new BackgroundWorker();
            bgWorker.WorkerReportsProgress = true;
            bgWorker.WorkerSupportsCancellation = true;
            bgWorker.RunWorkerCompleted +=
                new RunWorkerCompletedEventHandler(bgWorker_RunWorkerCompleted);
            bgWorker.ProgressChanged +=
                new ProgressChangedEventHandler(bgWorker_ProgressChanged);
        }

        private void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error != null)
            {
                blnBackGroundWorkIsOK = false;
            }
            else if (e.Cancelled)
            {
                blnBackGroundWorkIsOK = true;
            }
            else
            {
                blnBackGroundWorkIsOK = true;
            }
        }

        private void bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (frmMessageShow != null && frmMessageShow.Visible == true)
            {
                //设置显示的消息
                frmMessageShow.setMessage(e.UserState.ToString());
                //设置显示的按钮文字
                if (e.ProgressPercentage == clsConstant.Thread_Progress_OK)
                {
                    frmMessageShow.setStatus(clsConstant.Dialog_Status_Enable);
                }
            }
        }

        private void coverClaimReportToolStripMenuItem_Click(object sender, EventArgs e)
        {
            try
            {
                InitialBackGroundWorker();
                bgWorker.DoWork += new DoWorkEventHandler(KEYFile);

                bgWorker.RunWorkerAsync();
                // 启动消息显示画面
                frmMessageShow = new frmMessageShow(clsShowMessage.MSG_001,
                                                    clsShowMessage.MSG_007,
                                                    clsConstant.Dialog_Status_Disable);
                frmMessageShow.ShowDialog();
                // 数据读取成功后在画面显示
                if (blnBackGroundWorkIsOK)
                {
                    string ZFCEPath = Path.Combine(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Results"), "");
                    System.Diagnostics.Process.Start("explorer.exe", ZFCEPath);

                    // this.dataGridView.DataSource = null;
                    //this.dataGridView.AutoGenerateColumns = false;
                    //this.dataGridView.DataSource = bankResults;
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }

        }
        private void KEYFile(object sender, DoWorkEventArgs e)
        {
            zhuzaidiya_Result = new List<clsjinchushuju_budongchaninfo>();

            //初始化信息
            clsAllnew BusinessHelp = new clsAllnew();
            //BusinessHelp.pbStatus = pbStatus;
            //BusinessHelp.tsStatusLabel1 = toolStripLabel1;
            DateTime oldDate = DateTime.Now;
            zhuzaidiya_Result = BusinessHelp.Buiness_Bankcharge(ref this.bgWorker);
            DateTime FinishTime = DateTime.Now;  //   
            TimeSpan s = DateTime.Now - oldDate;
            string timei = s.Minutes.ToString() + ":" + s.Seconds.ToString();
            string Showtime = clsShowMessage.MSG_029 + timei.ToString();
            bgWorker.ReportProgress(clsConstant.Thread_Progress_OK, clsShowMessage.MSG_009 + "\r\n" + Showtime);
        }

    }
}
