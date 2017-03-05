using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

//Ado.net
using System.Data.SqlClient;  //引用SQL Server資料來源物件
using System.Data.OleDb; //引用System.Data.OleDb命名空間

namespace Gmail_Test
{
    public partial class Form1 : Form
    {
        
        /*Def for Windows Form*/
        string strBody, strEmail, strEngName, strMdate, strBankNo, strTable, strNo, strReason, strRemark;
        string[] sid;
        int money, totalmoney, sReduce = 0, iReduce = 0, reduce = 0;

        int MailCount;

        OleDbConnection Cnn;

        MailMessage message;
        SmtpClient smtp;

        public Form1()
        {
            InitializeComponent();
        }

        private void btnSelect_Click(object sender, EventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

            dialog.Multiselect = false;
            dialog.Title = "開啟舊檔";
            dialog.InitialDirectory = ".\\"; //視窗開啟時的初始目錄位置
            dialog.Filter = "Excel檔案(*.xls; *.xlsx)|*.xls; *.xlsx"; //要過濾，可以選擇的副檔名
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK && dialog.FileName != null)
            {
                //MessageBox.Show(dialog.FileName);       
                labPath.Text = dialog.FileName; //取得完整路徑
                labShow.Text = Path.GetFileName(dialog.FileName); //只取得檔名及副檔名
            }
            else
            {
                return;
            }
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            //1.檔案位置    注意絕對路徑 -> 非 \  是 \\
            string FileName = labPath.Text;
            //2.提供者名稱  Microsoft.Jet.OLEDB.4.0適用於2003以前版本，Microsoft.ACE.OLEDB.12.0 適用於2007以後的版本處理 xlsx 檔案
            string ProviderName = "Microsoft.ACE.OLEDB.12.0;";
            //3.Excel版本，Excel 8.0 針對Excel2000及以上版本，Excel5.0 針對Excel97。
            string ExtendedString = "'Excel 8.0;";
            //4.第一行是否為標題
            string Hdr = "Yes;";
            //5.IMEX=1 通知驅動程序始終將「互混」數據列作為文本讀取
            string IMEX = "0';";

            //連線字串
            string cs =
                    "Data Source=" + FileName + ";" +
                    "Provider=" + ProviderName +
                    "Extended Properties=" + ExtendedString +
                    "HDR=" + Hdr +
                    "IMEX=" + IMEX;
            //Excel 的工作表名稱 (Excel左下角有的分頁名稱)
            string SheetName = "Sheet1";

            using (OleDbConnection cn = new OleDbConnection(cs))
            {
                cn.Open();
                string qs = "select * from[" + SheetName + "$]";

                try
                {
                    using (OleDbDataAdapter dr = new OleDbDataAdapter(qs, cn))
                    {
                        DataTable dt = new DataTable();
                        dr.Fill(dt);
                        OleDbCommand cmd = cn.CreateCommand();
                        cmd.CommandText = qs;
                        OleDbDataReader DR = cmd.ExecuteReader();
                        strTable = "<table border=1 cellSpacing=0 cellPadding=0>" +
                                          "<tr><td align=center>帳款單號</td><td align=center>類別</td><td align=center>備註說明</td><td align=center>金額</td><td align=center>出差借支</td></tr>";
                        int addmon = 0;
                        while (DR.Read())
                        {

                            strEmail = DR["EMAIL"].ToString();
                            strEngName = DR["ENGNAME"].ToString();
                            strMdate = DR["MDATE"].ToString();
                            strBankNo = DR["BANKNO"].ToString();
                            strNo = DR["NO"].ToString();
                            strReason = DR["REASON"].ToString();
                            strRemark = DR["REMARK"].ToString();

                            if (!"".Equals(DR["REDUCE"].ToString().Trim()))
                            {
                                reduce += Convert.ToInt32(DR["REDUCE"].ToString().Trim());
                                //  MessageBox.Show(reduce+"");
                            }
                            else
                            {
                                reduce += 0;
                                // MessageBox.Show(reduce+"");
                            }
                            //  reduce = Convert.ToInt16(DR["REDUCE"].ToString().Trim());
                            money = Convert.ToInt32(DR["MONEY"].ToString().Trim());

                            addmon += money;
                            strTable = strTable + "<tr><td>&nbsp;" + strNo + "&nbsp;</td><td>&nbsp;" + strReason + "&nbsp;</td><td>&nbsp;" + strRemark + "&nbsp;</td><td align=right>&nbsp;" + money + "</td><td align=right>&nbsp;" + reduce + "</td><td></tr>";

                            totalmoney = addmon + reduce;

                            /*信件內容設定*/
                            message = new MailMessage();
                            message.From = new MailAddress("abc@abc.com");
                            message.To.Add(strEmail + "@gmail.com");
                            message.Subject = "總統府電子匯款通知";
                            message.SubjectEncoding = System.Text.Encoding.UTF8;
                            message.BodyEncoding = System.Text.Encoding.UTF8;
                            message.IsBodyHtml = true;

                            sid = strEngName.Split(' ');

                            strBody = "Dear " + sid[0] + ":<br><br>" +
                                                "將於 " + strMdate + " 存入帳號 " + strBankNo + ",<br>" +
                                                 "金額共計: " + totalmoney + " 明細如下<br>" +
                                                 strTable + "</table><br>" +
                            "PS: 出差借支為負數者，為RMB報銷大於出差借支，差額轉換NTD後，併入差旅費給付.<br>" +
                            "以上帳款若有問題，請與AA聯絡！<br>" +
                            "以上若有任何問題，請與BB聯絡！<br><br><br><br>" +
                            "**********************************************************************<br>" +
                            "John Wang(王約翰)<br>" +
                            "ACCOUNTING (會計部)<br>" +
                            "LINE Co., Ltd.<br>" +
                            "中華民國總統府網站<br>" +
                            "Tel:02-8825252 #123<br>" +
                            "Fax:037-321122<br>" +
                            "Website：http://www.president.gov.tw<br>" +
                            "Add: 10040 臺北市中正區重慶南路1段122號";

                            message.Body = strBody;
                            MailCount++;
                        }
                        for (int i = 1; i <= MailCount; i++)
                        {
                            // set smtp details
                            smtp = new SmtpClient("smtp.gmail.com");
                            smtp.Port = 587;
                            smtp.EnableSsl = true;
                            smtp.Credentials = new NetworkCredential("imgeniuss17@gmail.com", "vkwjppcassxxpgox");

                            smtp.SendAsync(message, message.Subject);
                            smtp.SendCompleted += new SendCompletedEventHandler(smtp_SendCompleted);
                        }
                        //  strTable = strTable + "</table>";
                        smtp.SendCompleted += new SendCompletedEventHandler(smtp_SendCompleted);

                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                }
            }

            // btnSubmit.Enabled = true;
            pnlShow.Visible = false;
        
            llabPre.Visible = true;       
        }

        void smtp_SendCompleted(object sender, AsyncCompletedEventArgs e)
        {
            if (e.Cancelled == true)
            {
                MessageBox.Show("信件傳送失敗！", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else if (e.Error != null)
            {
                MessageBox.Show(e.Error.Message);
            }
            else
            {
                MessageBox.Show("匯款通知作業成功，共發送了" + MailCount + "封Mail！", "提示訊息", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
            }

        }




        private void labPre_Click(object sender, EventArgs e)
        {
            pnlShow.Visible = true;
        }

    }
}
