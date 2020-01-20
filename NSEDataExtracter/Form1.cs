using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.IO;
using System.IO.Compression;
using System.Data.OleDb;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace NSEDataExtracter
{
    public partial class Form1 : Form
    {
        public DialogResult dialog;
        public Boolean isprocessstarted;
        private readonly string SourceLink = "https://www1.nseindia.com/live_market/dynaContent/live_watch/option_chain/optionKeys.jsp?symbolCode=-10006&symbol=NIFTY&symbol=NIFTY&instrument=-&date=-&segmentLink=17&symbolCount=2&segmentLink=17";
        private string DownloadedString;
        private HtmlAgilityPack.HtmlDocument h1,h2, h3,h4,h5;
        private string InputFile = Application.StartupPath + "\\Input\\InputTemplate.xlsx";
        private string outputfile = Application.StartupPath+ "\\Output\\";
        private bool IsExcel, IsDownloading;
        private IList<NSEData> NSEDataList1 = new List<NSEData>(), NSEDataList2 = new List<NSEData>();
        private NSEData CNSEExtractData, PNSEExtractData;
        private string Nifty = string.Empty, NiftyTime = string.Empty,cs1,ps1,cs2,ps2;
        private int cellcount = 4,pval1=0, pval2=0,Tcount=0;
        public Status valuestatus;
        private void Form1_Load(object sender, EventArgs e)
        {
            //tm.Interval = 240000;
            //tm.Interval = 1000;
            
        }

        public Form1()
        {
            InitializeComponent();
        }
        private void RunEvent(object sender, System.EventArgs e)
        {
            ProcessDataFetchProcess();
        }
        private void Button1_Click(object sender, EventArgs e)
        {
            button1.Enabled = false;
            button2.Enabled = true;
            isprocessstarted = true;
            progressBar1.Visible = true;
            backgroundWorker1.RunWorkerAsync(2000);
            MessageBox.Show("NSE Extract Process started Sucessfully", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Button2_Click(object sender, EventArgs e)
        {
            if (isprocessstarted)
            {
                dialog = MessageBox.Show("Are you sure want to exit NSE Process", "Alert", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (dialog == DialogResult.Yes)
                {
                    button1.Enabled = true;
                    button2.Enabled = false;
                    if (!backgroundWorker1.IsBusy)
                        backgroundWorker1.CancelAsync();
                    isprocessstarted = false;
                    timer.Stop();
                    progressBar1.Visible = false;
                    MessageBox.Show("NSE Extract Process Stopped Sucessfully", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            else
            {
                MessageBox.Show("NSE Process Not Yet Started", "Alert", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void BackgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            NSEScrapProcess();
        }
        private void NSEScrapProcess()
        {
            try
            {
                ExcelDataToList();
                cs1 = NSEDataList1[0].StockPrice1;
                ps1 = NSEDataList1[0].StockPrice2;
                cs2 = NSEDataList2[0].StockPrice1;
                ps2 = NSEDataList2[0].StockPrice2;
                timer.Interval = 240000;
                //timer.Interval = 1000;
                timer.Enabled = true;
                timer.Start();
                
           
               // timer.Tick += new EventHandler(RunEvent);
                //CreateOutput();
            }
            catch
            { }

        }
        public void ProcessDataFetchProcess() {
            try
            {
                LabelLog("Fetch process started " + DateTime.Now.ToString());
                #region DataDownloadRegion
                DownloadNSEData();
                #endregion
                CNSEExtractData = new NSEData();
                PNSEExtractData = new NSEData();
                valuestatus = new Status();
                h1 = null;
                h1 = new HtmlAgilityPack.HtmlDocument();
                h1.LoadHtml(DownloadedString);
                Nifty = string.Empty;
                NiftyTime = string.Empty;
                var Maindiv = h1.DocumentNode.SelectSingleNode("//div[@id='wrapper_btm']");
                h2 = null;
                h2 = new HtmlAgilityPack.HtmlDocument();
                h2.LoadHtml(Maindiv.InnerHtml.ToString());
                var niftydata = h2.DocumentNode.SelectNodes("//table");
                h3 = null;
                h3 = new HtmlAgilityPack.HtmlDocument();
                h3.LoadHtml(niftydata[0].InnerHtml.ToString());

                CNSEExtractData.NIFTY = h3.DocumentNode.SelectNodes("//span")[0].InnerText.ToString().Replace("Underlying Index:", "").Replace("NIFTY", "").Trim().ToString().Replace("&nbsp;","");
                CNSEExtractData.Time = h3.DocumentNode.SelectNodes("//span")[1].InnerText.ToString().Replace("As on ", "").Trim().ToString();

                PNSEExtractData.NIFTY = CNSEExtractData.NIFTY;
                PNSEExtractData.Time = CNSEExtractData.Time;

                h2 = null;
                h2 =new HtmlAgilityPack.HtmlDocument();
                h2.LoadHtml(h1.DocumentNode.SelectSingleNode("//table[@id='octable']").InnerHtml.ToString());

                var trdata = h2.DocumentNode.SelectNodes("//tr");
                    for (int i = 2; i < trdata.Count(); i++)
                    {
                        h4 = null;
                        h4 = new HtmlAgilityPack.HtmlDocument();
                        h4.LoadHtml(trdata[i].InnerHtml.ToString());
                        var td = h4.DocumentNode.SelectNodes("//td");
                        if (cs1+".00" == td[11].InnerText.ToString().Trim())
                        {
                            valuestatus.cs1 = true;
                            CNSEExtractData.OI1 = td[1].InnerText.ToString().Trim();
                            CNSEExtractData.COI1 = td[2].InnerText.ToString().Trim();
                            CNSEExtractData.VOL1 = td[3].InnerText.ToString().Trim();
                            CNSEExtractData.IV1 = td[4].InnerText.ToString().Trim();
                            CNSEExtractData.LPT1 = td[5].InnerText.ToString().Trim();
                            CNSEExtractData.NetC1 = td[6].InnerText.ToString().Trim();
                            CNSEExtractData.BIDQty1 = td[7].InnerText.ToString().Trim();
                            CNSEExtractData.BIDPrice1 = td[8].InnerText.ToString().Trim();
                            CNSEExtractData.Askqty1 = td[9].InnerText.ToString().Trim();
                            CNSEExtractData.ASKPrice1 = td[10].InnerText.ToString().Trim();
                            CNSEExtractData.StockPrice1 = cs1;
                            //CNSEExtractData
                        }
                        else if (ps1 + ".00" == td[11].InnerText.ToString().Trim())
                        {
                            valuestatus.ps1 = true;
                            CNSEExtractData.BIDQty2 = td[12].InnerText.ToString().Trim();
                            CNSEExtractData.BIDPrice2 = td[13].InnerText.ToString().Trim();
                            CNSEExtractData.Askqty2 = td[14].InnerText.ToString().Trim();
                            CNSEExtractData.ASKPrice2 = td[15].InnerText.ToString().Trim();
                            CNSEExtractData.NetC2 = td[16].InnerText.ToString().Trim();
                            CNSEExtractData.LPT2 = td[17].InnerText.ToString().Trim();
                            CNSEExtractData.IV2 = td[18].InnerText.ToString().Trim();
                            CNSEExtractData.VOL2 = td[19].InnerText.ToString().Trim();
                            CNSEExtractData.COI2 = td[20].InnerText.ToString().Trim();
                            CNSEExtractData.OI2 = td[21].InnerText.ToString().Trim();
                            CNSEExtractData.StockPrice2 = ps1;
                            //CNSEExtractData
                        }
                        else if (cs2 + ".00" == td[11].InnerText.ToString().Trim())
                        {
                            valuestatus.cs2 = true;
                            //PNSEExtractData
                            PNSEExtractData.OI1 = td[1].InnerText.ToString().Trim();
                            PNSEExtractData.COI1 = td[2].InnerText.ToString().Trim();
                            PNSEExtractData.VOL1 = td[3].InnerText.ToString().Trim();
                            PNSEExtractData.IV1 = td[4].InnerText.ToString().Trim();
                            PNSEExtractData.LPT1 = td[5].InnerText.ToString().Trim();
                            PNSEExtractData.NetC1 = td[6].InnerText.ToString().Trim();
                            PNSEExtractData.BIDQty1 = td[7].InnerText.ToString().Trim();
                            PNSEExtractData.BIDPrice1 = td[8].InnerText.ToString().Trim();
                            PNSEExtractData.Askqty1 = td[9].InnerText.ToString().Trim();
                            PNSEExtractData.ASKPrice1 = td[10].InnerText.ToString().Trim();
                            PNSEExtractData.StockPrice1 = cs2;

                        }
                        else if (ps2 + ".00" == td[11].InnerText.ToString().Trim())
                        {
                            valuestatus.ps2 = true;
                            //PNSEExtractData
                            PNSEExtractData.BIDQty2 = td[12].InnerText.ToString().Trim();
                            PNSEExtractData.BIDPrice2 = td[13].InnerText.ToString().Trim();
                            PNSEExtractData.Askqty2 = td[14].InnerText.ToString().Trim();
                            PNSEExtractData.ASKPrice2 = td[15].InnerText.ToString().Trim();
                            PNSEExtractData.NetC2 = td[16].InnerText.ToString().Trim();
                            PNSEExtractData.LPT2 = td[17].InnerText.ToString().Trim();
                            PNSEExtractData.IV2 = td[18].InnerText.ToString().Trim();
                            PNSEExtractData.VOL2 = td[19].InnerText.ToString().Trim();
                            PNSEExtractData.COI2 = td[20].InnerText.ToString().Trim();
                            PNSEExtractData.OI2 = td[21].InnerText.ToString().Trim();
                            PNSEExtractData.StockPrice2 = ps2;
                        }
                        if(valuestatus.cs1 && valuestatus.cs2 && valuestatus.ps1 && valuestatus.ps2)
                        {
                            break;
                        }

                    }
                    Tcount = NSEDataList1.Count();
                    if (NSEDataList1.Count()>2)
                    {
                    //CNSEExtractData
                    try
                    {
                        int.TryParse(NSEDataList1[Tcount - 1].LPT1, out pval1);
                        int.TryParse(CNSEExtractData.LPT1, out pval2);
                        CNSEExtractData.CLPTDecay = (pval2 - pval1).ToString();
                    }
                    catch { }
                    try
                    {
                        int.TryParse(NSEDataList1[Tcount - 1].IV1, out pval1);
                        int.TryParse(CNSEExtractData.IV1, out pval2);
                        CNSEExtractData.CIVCHANG = ((pval2 - pval1)/ pval2).ToString()+"%";
                }
                    catch { }
                    try
                    {
                        int.TryParse(NSEDataList1[Tcount - 1].LPT2, out pval1);
                        int.TryParse(CNSEExtractData.LPT2, out pval2);
                        CNSEExtractData.PLPTDecay = (pval2 - pval1).ToString();
        }
                    catch { }
                    try
                    {
                        int.TryParse(NSEDataList1[Tcount - 1].IV2, out pval1);
                        int.TryParse(CNSEExtractData.IV2, out pval2);
                        CNSEExtractData.CIVCHANG = ((pval2 - pval1) / pval2).ToString() + "%";
                    }
                    catch { }
                    }
                     NSEDataList1.Add(CNSEExtractData);
                     Tcount = NSEDataList2.Count();
                    if (NSEDataList2.Count() > 2)
                    {
                    try
                    {
                        int.TryParse(NSEDataList2[Tcount - 1].LPT1, out pval1);
                        int.TryParse(PNSEExtractData.LPT1, out pval2);
                        PNSEExtractData.CLPTDecay = (pval2 - pval1).ToString();

                    }
                    catch { }
                    try
                    {
                        int.TryParse(NSEDataList2[Tcount - 1].IV1, out pval1);
                        int.TryParse(PNSEExtractData.IV1, out pval2);
                        PNSEExtractData.CIVCHANG = ((pval2 - pval1) / pval2).ToString() + "%";

                    }
                    catch { }
                    try
                    {
                        int.TryParse(NSEDataList2[Tcount - 1].LPT2, out pval1);
                        int.TryParse(PNSEExtractData.LPT2, out pval2);
                        PNSEExtractData.PLPTDecay = (pval2 - pval1).ToString();
                    }
                    catch { }
                    try
                    {
                        int.TryParse(NSEDataList2[Tcount - 1].IV2, out pval1);
                        int.TryParse(PNSEExtractData.IV2, out pval2);
                        PNSEExtractData.CIVCHANG = ((pval2 - pval1) / pval2).ToString() + "%";
                    }
                    catch { }
                    //PNSEExtractData
                }
                NSEDataList2.Add(PNSEExtractData);
                CreateOutputAsync();
                LabelLog("Fetch process stopped " + DateTime.Now.ToString());
            }
            catch(Exception ex)
            {
            }
            finally
                {
                DownloadedString = string.Empty;
                h1 = null;
                h2 = null;
                h3 = null;
                h4 = null;
                PNSEExtractData = null;
                CNSEExtractData = null;
            }
        }
        public void DownloadNSEData()
        {
            HttpWebResponse response = null;
            HttpWebRequest request = null;
            try
            {
                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                request = (HttpWebRequest)WebRequest.Create(SourceLink);
                System.Net.HttpWebRequest.DefaultWebProxy = null;
                request.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8";
                request.Headers.Set(HttpRequestHeader.AcceptLanguage, "en-US,en;q=0.5");
                request.Headers.Add("Upgrade-Insecure-Requests", @"1");
                request.UserAgent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/70.0.3538.102 Safari/537.36 Edge/18.18363";
                request.Headers.Set(HttpRequestHeader.AcceptEncoding, "gzip, deflate, br");
               // request.Headers.Set(HttpRequestHeader.Cookie, @"ak_bmsc=0E3A9349B3DF534E65673CBD947B869417CB3F0E17760000D1AC225E90E49B47~plQxo13Xj9LqJQuR4W0ZqCkCN9Zy13IIz9lFdzmhDgYBfRhOuvkeEglmIwyY3sRGsmPKin7yNMRLMQH01a2foYO4QbCteemx6txZYq+uSCyRKkh4RWaoDVnOB8pI9W4j3oQHD2N4zqC0xjKJO3lRStpP7RdfRBc6mS9Vm2AT4uFKhvbmxMxuneJZm2RxEKmSNPtoeE+PO5KaJhEQGIAq1o+iaE0yv7Q/z6H7Y+SZ3cXuo=");
                using (response = (HttpWebResponse)request.GetResponse())
                {
                    DownloadedString = ReadResponse(response);
                }
            }
            catch (Exception ex) {

            }
        }
        private static string ReadResponse(HttpWebResponse response)
        {
            using (Stream responseStream = response.GetResponseStream())
            {
                Stream streamToRead = responseStream;
                if (response.ContentEncoding.ToLower().Contains("gzip"))
                {
                    streamToRead = new GZipStream(streamToRead, CompressionMode.Decompress);
                }
                else if (response.ContentEncoding.ToLower().Contains("deflate"))
                {
                    streamToRead = new DeflateStream(streamToRead, CompressionMode.Decompress);
                }

                using (StreamReader streamReader = new StreamReader(streamToRead, Encoding.UTF8))
                {
                    return streamReader.ReadToEnd();
                }
            }
        }

        private void ExcelDataToList()
        {
            try
            {
                var dsdata = ReadExcel(InputFile);
                #region FirstData
                CNSEExtractData = new NSEData();
                CNSEExtractData.Time = dsdata.Rows[3][0].ToString();
                CNSEExtractData.OI1 = dsdata.Rows[3][1].ToString();
                CNSEExtractData.COI1 = dsdata.Rows[3][2].ToString();
                CNSEExtractData.VOL1 = dsdata.Rows[3][3].ToString();
                CNSEExtractData.IV1 = dsdata.Rows[3][4].ToString();
                CNSEExtractData.LPT1 = dsdata.Rows[3][5].ToString();
                CNSEExtractData.NetC1 = dsdata.Rows[3][6].ToString();
                CNSEExtractData.BIDQty1 = dsdata.Rows[3][7].ToString();
                CNSEExtractData.BIDPrice1 = dsdata.Rows[3][8].ToString();
                CNSEExtractData.Askqty1 = dsdata.Rows[3][9].ToString();
                CNSEExtractData.ASKPrice1 = dsdata.Rows[3][10].ToString();
                CNSEExtractData.StockPrice1 = dsdata.Rows[3][11].ToString();
                CNSEExtractData.StockPrice2 = dsdata.Rows[3][12].ToString();
                CNSEExtractData.BIDQty2 = dsdata.Rows[3][13].ToString();
                CNSEExtractData.BIDPrice2 = dsdata.Rows[3][14].ToString();
                CNSEExtractData.Askqty2 = dsdata.Rows[3][15].ToString();
                CNSEExtractData.ASKPrice2 = dsdata.Rows[3][16].ToString();
                CNSEExtractData.NetC2 = dsdata.Rows[3][17].ToString();
                CNSEExtractData.LPT2 = dsdata.Rows[3][18].ToString();
                CNSEExtractData.IV2 = dsdata.Rows[3][19].ToString();
                CNSEExtractData.VOL2 = dsdata.Rows[3][20].ToString();
                CNSEExtractData.COI2 = dsdata.Rows[3][21].ToString();
                CNSEExtractData.OI2 = dsdata.Rows[3][22].ToString();
                CNSEExtractData.NIFTY = dsdata.Rows[3][23].ToString();
                CNSEExtractData.CLPTDecay = dsdata.Rows[3][24].ToString();
                CNSEExtractData.CIVCHANG = dsdata.Rows[3][25].ToString();
                CNSEExtractData.PLPTDecay = dsdata.Rows[3][26].ToString();
                CNSEExtractData.PIVCHANG = dsdata.Rows[3][27].ToString();
                CNSEExtractData.ACALL = dsdata.Rows[3][28].ToString();
                CNSEExtractData.APUT = dsdata.Rows[3][29].ToString();
                NSEDataList1.Add(CNSEExtractData);
                #endregion

                #region SecoundData
                PNSEExtractData = new NSEData();
                PNSEExtractData.Time = dsdata.Rows[4][0].ToString();
                PNSEExtractData.OI1 = dsdata.Rows[4][1].ToString();
                PNSEExtractData.COI1 = dsdata.Rows[4][2].ToString();
                PNSEExtractData.VOL1 = dsdata.Rows[4][3].ToString();
                PNSEExtractData.IV1 = dsdata.Rows[4][4].ToString();
                PNSEExtractData.LPT1 = dsdata.Rows[4][5].ToString();
                PNSEExtractData.NetC1 = dsdata.Rows[4][6].ToString();
                PNSEExtractData.BIDQty1 = dsdata.Rows[4][7].ToString();
                PNSEExtractData.BIDPrice1 = dsdata.Rows[4][8].ToString();
                PNSEExtractData.Askqty1 = dsdata.Rows[4][9].ToString();
                PNSEExtractData.ASKPrice1 = dsdata.Rows[4][10].ToString();
                PNSEExtractData.StockPrice1 = dsdata.Rows[4][11].ToString();
                PNSEExtractData.StockPrice2 = dsdata.Rows[4][12].ToString();
                PNSEExtractData.BIDQty2 = dsdata.Rows[4][13].ToString();
                PNSEExtractData.BIDPrice2 = dsdata.Rows[4][14].ToString();
                PNSEExtractData.Askqty2 = dsdata.Rows[4][15].ToString();
                PNSEExtractData.ASKPrice2 = dsdata.Rows[4][16].ToString();
                PNSEExtractData.NetC2 = dsdata.Rows[4][17].ToString();
                PNSEExtractData.LPT2 = dsdata.Rows[4][18].ToString();
                PNSEExtractData.IV2 = dsdata.Rows[4][19].ToString();
                PNSEExtractData.VOL2 = dsdata.Rows[4][20].ToString();
                PNSEExtractData.COI2 = dsdata.Rows[4][21].ToString();
                PNSEExtractData.OI2 = dsdata.Rows[4][22].ToString();
                PNSEExtractData.NIFTY = dsdata.Rows[4][23].ToString();
                PNSEExtractData.CLPTDecay = dsdata.Rows[4][24].ToString();
                PNSEExtractData.CIVCHANG = dsdata.Rows[4][25].ToString();
                PNSEExtractData.PLPTDecay = dsdata.Rows[4][26].ToString();
                PNSEExtractData.PIVCHANG = dsdata.Rows[4][27].ToString();
                PNSEExtractData.ACALL = dsdata.Rows[4][28].ToString();
                PNSEExtractData.APUT = dsdata.Rows[4][29].ToString();
                NSEDataList2.Add(PNSEExtractData);
                #endregion
            }
            catch(Exception ex)
            { }
            finally
            {
                CNSEExtractData = null;
                PNSEExtractData = null;
            }
        }
        public DataTable ReadExcel(string fileName, string fileExt= ".xlsx")
        {
            string conn = string.Empty;
            DataTable dtexcel = new DataTable();
            if (fileExt.CompareTo(".xls") == 0)
                conn = string.Format(@"provider=Microsoft.Jet.OLEDB.4.0;Data Source={0};Extended Properties='Excel 8.0;HRD=Yes;IMEX=1';",fileName); //for below excel 2007  
            else
                conn = string.Format(@"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties='Excel 12.0;HDR=NO';", fileName); //for above excel 2007  
            using (OleDbConnection con = new OleDbConnection(conn))
            {
                try
                {
                    OleDbDataAdapter oleAdpt = new OleDbDataAdapter("select * from [Sheet1$]", con); //here we read data from sheet1  
                    oleAdpt.Fill(dtexcel); //fill excel data into dataTable  
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Input Excel Reading process Failed","Alert");
                }
            }
            return dtexcel;
        }

        private async Task CreateOutputAsync()
        {
            try
            {
                ExcelPackage excel = new ExcelPackage();
                var workSheet = excel.Workbook.Worksheets.Add("OutputSheet");
                workSheet.Row(1).Height = 20;
                workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Row(1).Style.Font.Bold = true;
                workSheet.Cells[1, 5].Value = "CALL SIDE";
                workSheet.Cells[1, 18].Value = "PUT SIDE";
                workSheet.Cells[1, 24].Value = "NIFTY 50";
                workSheet.Cells[1, 25].Value = "CALL DECAY";
                workSheet.Cells[1, 27].Value = "PUT DECAY";
                workSheet.Cells[1, 29].Value = "AVERAGE DECAY";
                workSheet.Cells["E1:G1"].Merge = true;
                workSheet.Cells["R1:T1"].Merge = true;
                workSheet.Cells["Y1:Z1"].Merge = true;
                workSheet.Cells["AA1:AB1"].Merge = true;
                workSheet.Cells["AC1:AD1"].Merge = true;
                workSheet.Row(2).Height = 15;
                workSheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //workSheet.Row(2).Style.Fill = ExcelColor;
                workSheet.Row(2).Style.Font.Bold = true;


                workSheet.Cells[2, 1].Value = "TIME";
                workSheet.Cells["A2:A3"].Merge = true;
                workSheet.Cells[2, 2].Value = "OI";
                workSheet.Cells["B2:B3"].Merge = true;
                workSheet.Cells[2, 3].Value = "Chng in OI";
                workSheet.Cells["C2:C3"].Merge = true;
                workSheet.Cells[2, 4].Value = "Volume";
                workSheet.Cells["D2:D3"].Merge = true;
                workSheet.Cells[2, 5].Value = "IV";
                workSheet.Cells["E2:E3"].Merge = true;
                workSheet.Cells[2, 6].Value = "LTP";
                workSheet.Cells["F2:F3"].Merge = true;
                workSheet.Cells[2, 7].Value = "Net Chng";
                workSheet.Cells["G2:G3"].Merge = true;
                workSheet.Cells[2, 8].Value = "Bid";
                workSheet.Cells[3, 8].Value = "Qty";
                workSheet.Cells[2, 9].Value = "Bid";
                workSheet.Cells[3, 9].Value = "Price";
                workSheet.Cells[2, 10].Value = "Ask";
                workSheet.Cells[3, 10].Value = "Price";
                workSheet.Cells[2, 11].Value = "Ask";
                workSheet.Cells[3, 11].Value = "Qty";
                workSheet.Cells[2, 12].Value = "Strike Price";
                workSheet.Cells["L2:L3"].Merge = true;
                workSheet.Cells[2, 13].Value = "Strike Price";
                workSheet.Cells["M2:M3"].Merge = true;
                workSheet.Cells[2, 14].Value = "Bid";
                workSheet.Cells[3, 14].Value = "Qty";
                workSheet.Cells[2, 15].Value = "Bid";
                workSheet.Cells[3, 15].Value = "Price";
                workSheet.Cells[2, 16].Value = "Ask";
                workSheet.Cells[3, 16].Value = "Price";
                workSheet.Cells[2, 17].Value = "Ask";
                workSheet.Cells[3, 17].Value = "Qty";
                workSheet.Cells[2, 18].Value = "Net Chng";
                workSheet.Cells["R2:R3"].Merge = true;
                workSheet.Cells[2, 19].Value = "LTP";
                workSheet.Cells["S2:S3"].Merge = true;
                workSheet.Cells[2, 20].Value = "IV";
                workSheet.Cells["T2:T3"].Merge = true;
                workSheet.Cells[2, 21].Value = "Volume";
                workSheet.Cells["U2:U3"].Merge = true;
                workSheet.Cells[2, 22].Value = "Chng in OI";
                workSheet.Cells["V2:V3"].Merge = true;
                workSheet.Cells[2, 23].Value = "OI";
                workSheet.Cells["W2:W3"].Merge = true;
                workSheet.Row(3).Height = 15;
                workSheet.Row(3).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                //workSheet.Row(2).Style.Fill = ExcelColor;
                workSheet.Row(3).Style.Font.Bold = true;

                workSheet.Cells[3, 25].Value = "LTP DECAY";
                workSheet.Cells[3, 26].Value = "IV CHNG";
                workSheet.Cells[3, 27].Value = "LTP DECAY";
                workSheet.Cells[3, 28].Value = "IV CHNG";
                workSheet.Cells[3, 29].Value = "CALL";
                workSheet.Cells[3, 30].Value = "PUT";

                cellcount = 4;
                AssignExcelSheetValue(workSheet, excel);
            }
            catch(Exception ex)
            { }
        }
        private void AssignExcelSheetValue(ExcelWorksheet workSheet, ExcelPackage excel)
        {
            try
            {
                #region CellValue Assign
                foreach (var listdata in NSEDataList1)
                {
                    workSheet.Cells[cellcount, 1].Value = listdata.Time;
                    workSheet.Cells[cellcount, 2].Value = listdata.OI1;
                    workSheet.Cells[cellcount, 3].Value = listdata.COI1;
                    workSheet.Cells[cellcount, 4].Value = listdata.VOL1;
                    workSheet.Cells[cellcount, 5].Value = listdata.IV1;
                    workSheet.Cells[cellcount, 6].Value = listdata.LPT1;
                    workSheet.Cells[cellcount, 7].Value = listdata.NetC1;
                    workSheet.Cells[cellcount, 8].Value = listdata.BIDQty1;
                    workSheet.Cells[cellcount, 9].Value = listdata.BIDPrice1;
                    workSheet.Cells[cellcount, 10].Value = listdata.Askqty1;
                    workSheet.Cells[cellcount, 11].Value = listdata.ASKPrice1;
                    workSheet.Cells[cellcount, 12].Value = listdata.StockPrice1;
                    workSheet.Cells[cellcount, 13].Value = listdata.StockPrice2;
                    workSheet.Cells[cellcount, 14].Value = listdata.BIDQty2;
                    workSheet.Cells[cellcount, 15].Value = listdata.BIDPrice2;
                    workSheet.Cells[cellcount, 16].Value = listdata.Askqty2;
                    workSheet.Cells[cellcount, 17].Value = listdata.ASKPrice2;
                    workSheet.Cells[cellcount, 18].Value = listdata.NetC2;
                    workSheet.Cells[cellcount, 19].Value = listdata.LPT2;
                    workSheet.Cells[cellcount, 20].Value = listdata.IV2;
                    workSheet.Cells[cellcount, 21].Value = listdata.VOL2;
                    workSheet.Cells[cellcount, 22].Value = listdata.COI2;
                    workSheet.Cells[cellcount, 23].Value = listdata.OI2;
                    workSheet.Cells[cellcount, 24].Value = listdata.NIFTY;
                    workSheet.Cells[cellcount, 25].Value = listdata.CLPTDecay;
                    workSheet.Cells[cellcount, 26].Value = listdata.CIVCHANG;
                    workSheet.Cells[cellcount, 27].Value = listdata.PLPTDecay;
                    workSheet.Cells[cellcount, 28].Value = listdata.PIVCHANG;
                    workSheet.Cells[cellcount, 29].Value = listdata.ACALL;
                    workSheet.Cells[cellcount, 30].Value = listdata.APUT;


                    cellcount++;
                }
                cellcount += 3;
                foreach (var listdata in NSEDataList2)
                {
                    workSheet.Cells[cellcount, 1].Value = listdata.Time;
                    workSheet.Cells[cellcount, 2].Value = listdata.OI1;
                    workSheet.Cells[cellcount, 3].Value = listdata.COI1;
                    workSheet.Cells[cellcount, 4].Value = listdata.VOL1;
                    workSheet.Cells[cellcount, 5].Value = listdata.IV1;
                    workSheet.Cells[cellcount, 6].Value = listdata.LPT1;
                    workSheet.Cells[cellcount, 7].Value = listdata.NetC1;
                    workSheet.Cells[cellcount, 8].Value = listdata.BIDQty1;
                    workSheet.Cells[cellcount, 9].Value = listdata.BIDPrice1;
                    workSheet.Cells[cellcount, 10].Value = listdata.Askqty1;
                    workSheet.Cells[cellcount, 11].Value = listdata.ASKPrice1;
                    workSheet.Cells[cellcount, 12].Value = listdata.StockPrice1;
                    workSheet.Cells[cellcount, 13].Value = listdata.StockPrice2;
                    workSheet.Cells[cellcount, 14].Value = listdata.BIDQty2;
                    workSheet.Cells[cellcount, 15].Value = listdata.BIDPrice2;
                    workSheet.Cells[cellcount, 16].Value = listdata.Askqty2;
                    workSheet.Cells[cellcount, 17].Value = listdata.ASKPrice2;
                    workSheet.Cells[cellcount, 18].Value = listdata.NetC2;
                    workSheet.Cells[cellcount, 19].Value = listdata.LPT2;
                    workSheet.Cells[cellcount, 20].Value = listdata.IV2;
                    workSheet.Cells[cellcount, 21].Value = listdata.VOL2;
                    workSheet.Cells[cellcount, 22].Value = listdata.COI2;
                    workSheet.Cells[cellcount, 23].Value = listdata.OI2;
                    workSheet.Cells[cellcount, 24].Value = listdata.NIFTY;
                    workSheet.Cells[cellcount, 25].Value = listdata.CLPTDecay;
                    workSheet.Cells[cellcount, 26].Value = listdata.CIVCHANG;
                    workSheet.Cells[cellcount, 27].Value = listdata.PLPTDecay;
                    workSheet.Cells[cellcount, 28].Value = listdata.PIVCHANG;
                    workSheet.Cells[cellcount, 29].Value = listdata.ACALL;
                    workSheet.Cells[cellcount, 30].Value = listdata.APUT;

                    cellcount++;
                }
                #endregion
                string p_strPath = outputfile+DateTime.Now.ToString("yyyyddMM") +".xlsx";
                if (File.Exists(p_strPath))
                    File.Delete(p_strPath);

                string modelRange = "A1:AD" + cellcount;
                var modelTable = workSheet.Cells[modelRange];

                // Assign borders
                modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
                modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
                modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
                modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                // Fill worksheet with data to export
                modelTable.AutoFitColumns();
                excel.SaveAs(new FileInfo(p_strPath));

            }
            catch(Exception ex)
            { }
        }
        private void LabelLog(string Logmsg)
        {
            label1.Invoke((MethodInvoker)(() => label1.Text = Logmsg.ToString()));

        }
    }
    public class NSEData
    {
        public string Time { get; set; }
        public string OI1 { get; set; }
        public string COI1 { get; set; }
        public string VOL1 { get; set; }
        public string IV1 { get; set; }
        public string LPT1 { get; set; }
        public string NetC1 { get; set; }
        public string BIDQty1 { get; set; }
        public string BIDPrice1 { get; set; }
        public string Askqty1 { get; set; }
        public string ASKPrice1 { get; set; }
        public string StockPrice1 { get; set; }
        public string StockPrice2 { get; set; }
        public string BIDQty2 { get; set; }
        public string BIDPrice2 { get; set; }
        public string Askqty2 { get; set; }
        public string ASKPrice2 { get; set; }
        public string NetC2 { get; set; }
        public string LPT2 { get; set; }
        public string IV2 { get; set; }
        public string VOL2 { get; set; }
        public string COI2 { get; set; }
        public string OI2 { get; set; }
        public string NIFTY { get; set; }
        public string CLPTDecay { get; set; }
        public string CIVCHANG { get; set; }
        public string PLPTDecay { get; set; }
        public string PIVCHANG { get; set; }
        public string ACALL { get; set; }
        public string APUT { get; set; }
    }
    public class Status
    {
        public bool cs1 { get; set; }
        public bool cs2 { get; set; }
        public bool ps1 { get; set; }
        public bool ps2 { get; set; }
    }
}
