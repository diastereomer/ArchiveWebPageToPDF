using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Pechkin;
using System.IO;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Data;
using System.Threading;
using System.Configuration;
using System.Collections.Concurrent;
using System.Net;
using System.Drawing.Printing;
//using Pechkin.Synchronized;

namespace ArchiveWebPageToPDF
{
    class Program
    {
        static bool isProcessDone = true;
        static int threadNumber = 0;
        static int pdfThreadNumber = 0;
        static List<string> ls = new List<string>();
        static ConcurrentQueue<DataRow> htmlQueue = new ConcurrentQueue<DataRow>();
        static ConcurrentQueue<DataRow> dataRowQueue = null;
        static string environment = "1";
        static int maxTreadNumber = 1;
        static int rounds = 1;

        [STAThread]
        static void Main(string[] args)
        {
            Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
            maxTreadNumber = int.Parse(ConfigurationManager.AppSettings["MaxThreadNumber"].ToString());
            string[] paras = args[0].Split(',');
            environment = paras[0];
            switch (environment)
            {
                case "1":
                    environment = "DEV";
                    break;
                case "2":
                    environment = "QA";
                    break;
                case "3":
                    environment = "UAT";
                    break;
                case "4":
                    environment = "PROD";
                    break;
                default:
                    break;
            }

            // get data from databases
            string connetionStr = ConfigurationManager.ConnectionStrings["SQL" + environment].ConnectionString;

            string cmd = "SELECT [ID],[Control_ID],[Claim_Number],[Claimant_Number],[Form_ID],[sequence],[FormName] ,[URL],[URLDate],[URLFlag],[AttachmentName],[PDFName],[PDFDate] ,[Processed] FROM test where id> (1591748 +" + (1000 * int.Parse(paras[1])).ToString() + ") and id<=(1591748 +"
                + (1000 * int.Parse(paras[1]) + 1000).ToString() + ")";// [TRANSFORMDB].[dbo].[CMT_ASPX_to_PDF_URL]";
            using (SqlConnection sc = new SqlConnection(connetionStr))
            {
                SqlCommand scmd = new SqlCommand(cmd, sc);
                try
                {
                    sc.Open();
                    SqlDataAdapter sda = new SqlDataAdapter(scmd);
                    DataSet sd = new DataSet();
                    sda.Fill(sd);
                    dataRowQueue = new ConcurrentQueue<DataRow>(sd.Tables[0].AsEnumerable().ToList<DataRow>());
                    sda.Dispose();
                }
                catch (Exception e)
                {
                    throw e;
                }
            }

            //get the html contents
            while (threadNumber < maxTreadNumber)
            {
                //control thread amount
                threadNumber++;
                var thread = getWebBrowerThread();
                //thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
            }

            Console.WriteLine("started");

            while (dataRowQueue != null && dataRowQueue.Count > 0 && htmlQueue.Count == 0)
            {
                Thread.Sleep(1000);
            }

            while (pdfThreadNumber < maxTreadNumber)
            {
                pdfThreadNumber++;
                Thread pdfThread = new Thread(() =>
                {
                    GlobalConfig gc = new GlobalConfig();
                    gc.SetMargins(new Margins(100, 100, 100, 100));
                    gc.SetPaperSize(PaperKind.Letter);
                    IPechkin pechkin = new SimplePechkin(gc);
                    ObjectConfig configuration = new ObjectConfig();
                    configuration.SetCreateExternalLinks(true).SetFallbackEncoding(Encoding.UTF8).SetLoadImages(true).SetCreateForms(false);
                    /*test error*/
                    int m = 0;
                    /*end*/

                    while (htmlQueue.Count != 0)
                    {
                        //while(threadNumber>=maxTreadNumber)
                        //    {
                        //        Thread.Sleep(1000);
                        //    }

                        /*test error*/
                        m++;
                        if (m == 50)
                        {
                            throw (new Exception("forced"));
                        }

                        /*end*/

                        while (dataRowQueue != null && dataRowQueue.Count > 0 && htmlQueue.Count == 0)
                        {
                            Thread.Sleep(1000);
                        }

                        /*isProcessDone = false;
                        var subThread = new Thread(() =>
                        {*/
                        //threadNumber++;

                        DataRow dr;
                        if (htmlQueue.TryDequeue(out dr))
                        {
                            rounds++;
                            try
                            {
                                byte[] pdfBuf = pechkin.Convert(configuration, dr["URL"].ToString());
                                var testFile = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                                    "testpdf\\" + dr["Claim_Number"].ToString() + "_" + dr["Claimant_Number"].ToString() + "_" + dr["Form_ID"].ToString() + "_" + dr["sequence"].ToString() + ".pdf");
                                System.IO.File.WriteAllBytes(testFile, pdfBuf);
                                pdfBuf = null;
                                dr.Delete();
                                dr = null;
                                testFile = null;
                            }
                            catch (Exception e)
                            {
                                MessageBox.Show(e.Message);
                            }
                            if (rounds % 1000 == 0)
                            {
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                            }
                        }

                        Application.ApplicationExit += Application_PDFApplicationExit;
                        Application.ExitThread();
                        Application.Exit();
                    }
                });
                pdfThread.Start();
            }
        }



        //get the web brower thread
        protected static Thread getWebBrowerThread()
        {
            return new Thread(() =>
            {
                CookieContainer cookies = new CookieContainer();
                HttpWebRequest wrq = (HttpWebRequest)WebRequest.Create(ConfigurationManager.AppSettings["CMT" + environment]);
                wrq.CookieContainer = cookies;
                wrq.UseDefaultCredentials = true;
                wrq.PreAuthenticate = true;
                wrq.Credentials = CredentialCache.DefaultCredentials;
                wrq.AllowAutoRedirect = true;

                HttpWebResponse wrp = (HttpWebResponse)wrq.GetResponse();
                cookies.Add(wrp.Cookies);

                /*loop the urls to save the htmls*/
                if (dataRowQueue != null)
                {
                    while (dataRowQueue.Count != 0)
                    {
                        while (htmlQueue.Count > 500 && dataRowQueue.Count != 0)
                        {
                            Thread.Sleep(60000);
                        }

                        DataRow dataRow;
                        if (dataRowQueue.TryDequeue(out dataRow))
                        {
                            wrq = (HttpWebRequest)WebRequest.Create(dataRow["URL"].ToString());
                            wrq.CookieContainer = cookies;
                            wrq.UseDefaultCredentials = true;
                            wrq.Credentials = CredentialCache.DefaultCredentials;
                            wrq.AllowAutoRedirect = true;
                            wrp = (HttpWebResponse)wrq.GetResponse();

                            using (StreamReader sr = new StreamReader(wrp.GetResponseStream(), System.Text.Encoding.GetEncoding("UTF-8")))
                            {
                                string content = sr.ReadToEnd();
                                content = content.Replace("TEXTAREA", "LABEL");
                                content = content.Replace("textarea", "LABEL");
                                content = content.Replace("Textarea", "LABEL");
                                dataRow["URL"] = content;
                                content = null;
                                htmlQueue.Enqueue(dataRow);
                                dataRow = null;
                                //Console.WriteLine("enqueue");
                            }
                        }
                    }
                }
                wrq = null;
                wrp = null;
                Application.ApplicationExit += Application_ApplicationExit;
                Application.ExitThread();
                Application.Exit();
            });
        }

        //control the thread amount
        private static void Application_ApplicationExit(object sender, EventArgs e)
        {
            threadNumber--;
            //isProcessDone = true;
        }

        private static void Application_PDFApplicationExit(object sender, EventArgs e)
        {
            pdfThreadNumber--;
        }
    }
}