using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.tool.xml;
using iTextSharp.tool.xml.css.parser;
using iTextSharp.tool.xml.net;
using ZXing;

namespace DocumentFile
{
    class Program
    {
        static string filePath = System.Configuration.ConfigurationManager.AppSettings.Get("FilePath");
        static string PicPath = System.Configuration.ConfigurationManager.AppSettings.Get("PicPath");
        static string RQCodePath = System.Configuration.ConfigurationManager.AppSettings.Get("RQCodePath");
        static string MarkPath = System.Configuration.ConfigurationManager.AppSettings.Get("MarkPath");
        static string LogPath = System.Configuration.ConfigurationManager.AppSettings.Get("LogPath");
        static string GifPath = System.Configuration.ConfigurationManager.AppSettings.Get("GifPath");
        static string SetCaseNo = System.Configuration.ConfigurationManager.AppSettings.Get("SetCaseNo");
        static string ProcessDate = System.Configuration.ConfigurationManager.AppSettings.Get("ProcessDate");
        static IOConn ioCon = new IOConn();
        static string cssStr = "";

        static void Main(string[] args)
        {
            recLog("PDFforUpload 程式開始執行:" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));

            string pdate = addDate(-Convert.ToInt32(ProcessDate));
            cssStr = File.ReadAllText("main.css");
            //批次產生PDF檔: 現場圖, 現場照片, 分析研判表, + 表一,表二
            //條件: 日期( N天前) + '已上傳'
            //      如有指定案號(SetCaseNo)則無視其他條件.
            string sql = "select a.* ";
            sql += "from CaseRec a,CaseStatuRec c where a.CaseNo = c.CaseNo ";
            if (SetCaseNo.Trim().Equals(""))
            {
                sql += "and a.OccurTime >= '" + pdate + "' ";
                sql += "and a.CaseFlag in ('1','2') ";  //A3不上傳
               // sql += "and isnull(c.SendCheck,'') = '3' ";  //'3': 已上傳
                sql += "and isnull(c.SendCheck,'') = '3' ";  //'3': 已上傳
                sql += "and isnull(c.SendCheck,'') not in ('3','4') and c.CaseFlow in ('05','07','08','12') and c.CaseStatu='1' ";  //中端可上傳
            }
            else
            {
                sql += "and a.CaseNo = '" + SetCaseNo + "' ";
            }
            
            List<Dictionary<string, string>> lsdt = ioCon.SelectDct(sql);
            int err = 0;
            for (int i = 0; i < lsdt.Count; i++)
            {
                Dictionary<string, string> row = lsdt[i];
                try
                {
                    GenDocument1(row);  //現場圖
                    GenDocument2(row);  //現場照片
                    GenDocument3(row);  //分析研判表
                    GenDocument4(row);  //表一
                    GenDocument5(row);  //表二
                }
                catch (Exception e)
                {
                    recLog("CaseNo :" + row["CaseNo"] + " Error:" + e.Message);
                    err++;
                }
            }

            //deleteTmp();
            recLog("讀入筆數:" + lsdt.Count);
            recLog("錯誤筆數:" + err);
            recLog("PDFforUpload 程式執行結束:" + DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss"));
        }

        static void GenDocument1(Dictionary<string, string> row)
        {
            string caseNo = row["CaseNo"];
            string html = File.ReadAllText("doc1.html");

            /***********************************************************************************************/
            string imgPath = PicPath + caseNo.Substring(0, 5) + "\\" + caseNo + "\\Media\\";
            html = html.Replace("@caseNo", "*" + caseNo + "-S01*");
            html = html.Replace("@IeokNo", row["IeokNo"]);
            html = html.Replace("@Unit", getUnit(row["OccurPoliceUnit"]));
            html = html.Replace("@ProcessNo", caseNo);
            string[] sts = new string[5];
            for (int i = 0; i < 5; i++) sts[i] = "";
            switch (row["CaseFlag"])
            {
                case "1": sts[1] = "V"; break;
                case "2": sts[2] = "V"; break;
                default: sts[3] = "V"; break;
            }
            html = html.Replace("@V1", sts[1]);
            html = html.Replace("@V2", sts[2]);
            html = html.Replace("@V3", sts[3]);

            string sql = "select a.*,b.FSPath,b.FName from CaseRepData a,CaseMediaFile b ";
            sql += "where a.CaseNo = b.CaseNo and a.DocFld13 = b.MediaSno ";
            sql += "and a.DocNm='現場圖' and  b.MediaNm='現場圖' ";
            sql += "and a.CaseNo='" + caseNo + "' order by a.DocSno";
            List<Dictionary<string, string>> lsdy = ioCon.SelectDct(sql);
            if (lsdy.Count == 0) return;

            /***********************************************************************************************/
            string tmpfil = filePath + caseNo + "_04.pdf";
            Document doc = new Document(PageSize.A4.Rotate(), 15, 15, 15, 15);
            FileStream fs = new FileStream(tmpfil, FileMode.Create);
            PdfWriter pdf = PdfWriter.GetInstance(doc, fs);
            doc.Open();
            /***********************************************************************************************/
            Dictionary<string, string> col = lsdy[0];
            html = html.Replace("@OccurTime", getDateFrm(row["OccurTime"], 3));
            html = html.Replace("@DocFld5", getLigthScp(col["DocFld5"]));
            string adr = "地點:" + getAddr(row);
            if (row["CaseType"] == "3") adr += "<br/>[事後報案]";

            html = html.Replace("@Pich", imgPath + col["FName"]);
            html = html.Replace("@Addr", getSkip(adr, 35));
            html = html.Replace("@Point", getPoint(col["DocFld14"]));
            html = html.Replace("@Per", getPer(col["DocFld4"], col["DocFld15"]));
            html = html.Replace("@DocFld6", col["DocFld7"]);
            html = html.Replace("@DocFld7", col["DocFld6"]);
            html = html.Replace("@Scrip", row["Course"].Replace("<", "&lt;").Replace(">", "&gt;"));
            html = html.Replace("@Remake", col["DocTxt2"]);
            html = html.Replace("@CreatDate", getDateFrm(col["DocDate1"], 4));
            string mark = getPsnMark(row["ProcessId"]);
            html = html.Replace("@Mark1", MarkPath + mark);
            mark = getChfMark(row["CaseNo"]);
            html = html.Replace("@Mark2", MarkPath + mark);
            mark = getUntMarkh(row["OccurPoliceStation"], row["OccurPoliceUnit"]);
            html = html.Replace("@Mark3", MarkPath + mark);
            /***********************************************************************************************/
            byte[] bytehtml = System.Text.Encoding.UTF8.GetBytes(html);
            MemoryStream mshtml = new MemoryStream(bytehtml);
            byte[] bytecss = System.Text.Encoding.UTF8.GetBytes(cssStr);
            MemoryStream mscss = new MemoryStream(bytecss);

            XMLWorkerHelper.GetInstance().ParseXHtml(pdf, doc, mshtml, mscss);

            PdfDestination pdfDest = new PdfDestination(PdfDestination.XYZ, 0, doc.PageSize.Height, 1f);
            PdfAction action = PdfAction.GotoLocalPage(1, pdfDest, pdf);
            pdf.SetOpenAction(action);

            doc.Close();
            //EncryptPDF(tmpfil, filePath + row["ApplyNo"] + "_1.pdf", row["IdNo"]);
            //if (row["MailFlag1"] != "1") UpdateMailFlag(row["ApplyNo"], "1");
        }

        static void GenDocument2(Dictionary<string, string> row)
        {
            string caseNo = row["CaseNo"];
            string head = File.ReadAllText("doc20.html");
            string pict = File.ReadAllText("doc21.html");
            string html = "";

            /***********************************************************************************************/
            string mark = getPsnMark(row["ProcessId"]);
            head = head.Replace("@Mark", MarkPath + mark);

            string imgPath = PicPath + caseNo.Substring(0, 5) + "\\" + caseNo + "\\Media\\";
            head = head.Replace("@caseNo", "*" + caseNo + "-S15*");
            string sql = "select a.*,b.FSPath,b.FName from CaseRepData a,CaseMediaFile b ";
            sql += "where a.CaseNo = b.CaseNo and a.DocFld13 = b.MediaSno ";
            sql += "and a.DocNm='照片黏貼紀錄表' and  b.MediaNm='事故照片黏貼圖檔' ";
            sql += "and a.CaseNo='" + caseNo + "' order by a.DocSno";
            List<Dictionary<string, string>> lsdy = ioCon.SelectDct(sql);
            if (lsdy.Count == 0) return;

            /***********************************************************************************************/
            string tmpfil = filePath + caseNo + "_03.pdf";
            Document doc = new Document(PageSize.A4, 15, 10, 10, 15);
            FileStream fs = new FileStream(tmpfil, FileMode.Create);
            PdfWriter pdf = PdfWriter.GetInstance(doc, fs);
            doc.Open();
            /***********************************************************************************************/
            for (int i = 0; i < lsdy.Count; i++)
            {
                Dictionary<string, string> col = lsdy[i];
                if (i % 2 == 0)
                {
                    html += head;
                    html += pict;
                }
                else
                {
                    html += "<div style=\"height:10px\"></div>";
                    html += pict;
                }
                html = html.Replace("@Pic", imgPath + col["FName"]);
                html = html.Replace("@OccurTime", getDateFrm(col["DocTime1"], 2));
                html = html.Replace("@Seq", col["DocSno"]);
                html = html.Replace("@Remake", getScp(col));
            }

            /***********************************************************************************************/
            byte[] bytehtml = System.Text.Encoding.UTF8.GetBytes(html);
            MemoryStream mshtml = new MemoryStream(bytehtml);
            byte[] bytecss = System.Text.Encoding.UTF8.GetBytes(cssStr);
            MemoryStream mscss = new MemoryStream(bytecss);

            XMLWorkerHelper.GetInstance().ParseXHtml(pdf, doc, mshtml, mscss);

            PdfDestination pdfDest = new PdfDestination(PdfDestination.XYZ, 0, doc.PageSize.Height, 1f);
            PdfAction action = PdfAction.GotoLocalPage(1, pdfDest, pdf);
            pdf.SetOpenAction(action);

            doc.Close();
            //EncryptPDF(tmpfil, filePath + row["ApplyNo"] + "_2.pdf", row["IdNo"]);
            //if (row["MailFlag2"] != "1") UpdateMailFlag(row["ApplyNo"], "2");
        }

        static void GenDocument3(Dictionary<string, string> row)
        {
            string caseNo = row["CaseNo"];
            string html = File.ReadAllText("doc3.html");
            string codeMsg = caseNo + ";" + row["OccurPoliceUnit"] + ";" + row["OccurPoliceStation"] + ";";
            string qrCode = RQCodePath + caseNo + ".png";
            codeMsg += row["ProcessId"] + ";10.116.1.219;" + addDate(0);
            makeQRcode(codeMsg, qrCode);

            html = html.Replace("@caseNo", "*" + caseNo + "-S23*");
            //html = html.Replace("@appName", row["Name"]);
            html = html.Replace("@appName", "");
            html = html.Replace("@OccurTime", getDateFrm(row["OccurTime"], 1));
            html = html.Replace("@OccurAddr", getAddr(row));
            html = html.Replace("@QRCode", qrCode);
            string remake = "reference_only2.jpg";
            html = html.Replace("@reMake", remake);

            string sql = "select * from CasePerson where CaseNo = '" + caseNo + "'";
            List<Dictionary<string, string>> lsdy = ioCon.SelectDct(sql);
            if (lsdy.Count == 0) return;

            /***********************************************************************************************/
            string tmpfil = filePath + caseNo + "_13.pdf";
            Document doc = new Document(PageSize.A4, 10, 10, 10, 10);
            FileStream fs = new FileStream(tmpfil, FileMode.Create);
            PdfWriter pdf = PdfWriter.GetInstance(doc, fs);
            doc.Open();
            /***********************************************************************************************/
            for (int i = 1; i < 4; i++)
            {
                if (i > lsdy.Count)
                {
                    html = html.Replace("@PersonCname" + i, "");
                    html = html.Replace("@CarKind" + i, "");
                    html = html.Replace("@CarNo" + i, "");
                    html = html.Replace("@Name" + i, "");
                    html = html.Replace("@PerCause" + i, "");
                }
                else
                {
                    Dictionary<string, string> col = lsdy[i - 1];
                    html = html.Replace("@PersonCname" + i, col["PersonCname"]);
                    string carKind = getMItemVal("車種類別", col["OccurReport26"], 1);
                    if (carKind == "人")
                        carKind = getMItemVal("車種類別", col["OccurReport26"], 0);
                    else
                        carKind = getMItemVal("車種類別", col["OccurReport26"], 2);
                    html = html.Replace("@CarKind" + i, carKind);
                    string carNo = col["OccurReport27"];
                    html = html.Replace("@CarNo" + i, carNo);
                    string drName = col["PersonCname"];
                    html = html.Replace("@Name" + i, drName);
                    string perCause = col["PerCause"];
                    if (perCause == "")
                    {
                        perCause = getMItemVal("肇事因素", col["OccurReport34_1"], 0);
                        string otherA = getAAScp(col["OccurReport30"]);
                        string otherB = getBBScp(col["OccurReport32"]);
                        if (perCause != "")
                        {
                            if (otherA != "") perCause += "," + otherA;
                            if (otherB != "") perCause += "," + otherB;
                        }
                    }
                    html = html.Replace("@PerCause" + i, perCause);
                }
                string sysdat = DateTime.Now.ToString("yyyyMMdd");
                html = html.Replace("@sysDate", getDateFrm(sysdat, 0));
            }
            //string mark = getPsnMark(row["ProcessId"]);  //2018.11.16 處理人員改為結案人員
            string mark = getCh12IdMark(row["CaseNo"]);
            html = html.Replace("@Mark1", MarkPath + mark);
            mark = getUntMark(row["OccurPoliceStation"], row["OccurPoliceUnit"]);
            html = html.Replace("@Mark2", MarkPath + mark);

            /***********************************************************************************************/
            byte[] bytehtml = System.Text.Encoding.UTF8.GetBytes(html);
            MemoryStream mshtml = new MemoryStream(bytehtml);
            byte[] bytecss = System.Text.Encoding.UTF8.GetBytes(cssStr);
            MemoryStream mscss = new MemoryStream(bytecss);

            XMLWorkerHelper.GetInstance().ParseXHtml(pdf, doc, mshtml, mscss);

            PdfDestination pdfDest = new PdfDestination(PdfDestination.XYZ, 0, doc.PageSize.Height, 1f);
            PdfAction action = PdfAction.GotoLocalPage(1, pdfDest, pdf);
            pdf.SetOpenAction(action);

            doc.Close();
            //EncryptPDF(tmpfil, filePath + row["ApplyNo"] + "_3.pdf", row["IdNo"]);
            //if (row["MailFlag3"] != "1") UpdateMailFlag(row["ApplyNo"], "3");
        }

        static void GenDocument4(Dictionary<string, string> row)
        {
            string caseNo = row["CaseNo"];
            string html = File.ReadAllText("doc10.html");

            /***********************************************************************************************/
            html = html.Replace("@caseNo", "*" + caseNo + "-S02*");
            html = html.Replace("@IeokNo", row["IeokNo"]);
            html = html.Replace("@Unit", getUnit(row["OccurPoliceUnit"]));
            html = html.Replace("@ProcessNo", caseNo);
            string[] sts = new string[5];
            for (int i = 0; i < 5; i++) sts[i] = "";
            switch (row["IeokNo"])
            {
                case "1": sts[1] = "V"; break;
                case "2": sts[2] = "V"; break;
                default: sts[3] = "V"; break;
            }
            html = html.Replace("@V1", sts[1]);
            html = html.Replace("@V2", sts[2]);
            html = html.Replace("@V3", sts[3]);

            /***********************************************************************************************/
            string tmpfil = filePath + caseNo + "_10.pdf";
            Document doc = new Document(PageSize.A4.Rotate(), 15, 15, 15, 15);
            FileStream fs = new FileStream(tmpfil, FileMode.Create);
            PdfWriter pdf = PdfWriter.GetInstance(doc, fs);
            doc.Open();
            /***********************************************************************************************/

            string OccurTime = row["OccurTime"];
            html = html.Replace("@YY", OccurTime.Substring(0, 3));
            html = html.Replace("@MM", OccurTime.Substring(3, 2));
            html = html.Replace("@DD", OccurTime.Substring(5, 2));
            html = html.Replace("@HH", OccurTime.Substring(7, 2));
            html = html.Replace("@MI", OccurTime.Substring(9, 2));
            html = html.Replace("@WW", getCWeek(OccurTime.Substring(0, 7)));

            html = html.Replace("@OccurAddr1_1", row["OccurAddr1_1"]);
            //string adr = "地點:" + getAddr(row);
            html = html.Replace("@Addr1", getSkip(getAddr(row), 35));
            html = html.Replace("@Addr2", "");
            html = html.Replace("@Addr3", "");
            html = html.Replace("@DieNum", row["DieNum"]);
            html = html.Replace("@HurtNum", row["HurtNum"]);
            html = html.Replace("@DieO24Num", row["DieO24Num"]);

            html = html.Replace("@OccurReport4", row["OccurReport4"]);
            html = html.Replace("@OccurReport5", row["OccurReport5"]);
            html = html.Replace("@OccurReport6", row["OccurReport6"]);
            html = html.Replace("@OccurReport7", row["OccurReport7"]);
            html = html.Replace("@OccurReport8", row["OccurReport8"]);
            html = html.Replace("@OccurReport9", row["OccurReport9"]);
            html = html.Replace("@OccurReport10_1", row["OccurReport10_1"]);
            html = html.Replace("@OccurReport10_2", row["OccurReport10_2"]);
            html = html.Replace("@OccurReport10_3", row["OccurReport10_3"]);
            html = html.Replace("@OccurReport11_1", row["OccurReport11_1"]);
            html = html.Replace("@OccurReport11_2", row["OccurReport11_2"]);
            html = html.Replace("@OccurReport12_1", row["OccurReport12_1"]);
            html = html.Replace("@OccurReport12_2", row["OccurReport12_2"]);
            html = html.Replace("@OccurReport13", row["OccurReport13"]);
            html = html.Replace("@OccurReport14_1", row["OccurReport14_1"]);
            html = html.Replace("@OccurReport14_2", row["OccurReport14_2"]);
            html = html.Replace("@OccurReport14_3", row["OccurReport14_3"]);
            html = html.Replace("@OccurReport15", row["OccurReport15"]);

            string gifx = "";
            for (int i = 1; i < 16; i++)
            {
                gifx = GifPath + "no"+ i +".gif";
                html = html.Replace("@no" + i.ToString("00"), gifx);
            }

            /***********************************************************************************************/
            byte[] bytehtml = System.Text.Encoding.UTF8.GetBytes(html);
            MemoryStream mshtml = new MemoryStream(bytehtml);
            byte[] bytecss = System.Text.Encoding.UTF8.GetBytes(cssStr);
            MemoryStream mscss = new MemoryStream(bytecss);

            XMLWorkerHelper.GetInstance().ParseXHtml(pdf, doc, mshtml, mscss);

            PdfDestination pdfDest = new PdfDestination(PdfDestination.XYZ, 0, doc.PageSize.Height, 1f);
            PdfAction action = PdfAction.GotoLocalPage(1, pdfDest, pdf);
            pdf.SetOpenAction(action);

            doc.Close();
            //EncryptPDF(tmpfil, filePath + row["ApplyNo"] + "_1.pdf", row["IdNo"]);
            //if (row["MailFlag1"] != "1") UpdateMailFlag(row["ApplyNo"], "1");
        }

        static void GenDocument5(Dictionary<string, string> row)
        {
            string caseNo = row["CaseNo"];
            string html = File.ReadAllText("doc11.html");
            string NullStr = "&nbsp;";

            /***********************************************************************************************/
            string imgPath = PicPath + caseNo.Substring(0, 5) + "\\" + caseNo + "\\Media\\";
            html = html.Replace("@caseNo", "*" + caseNo + "-S03*");
            html = html.Replace("@IeokNo", row["IeokNo"]);
            html = html.Replace("@Unit", getUnit(row["OccurPoliceUnit"]));
            html = html.Replace("@ProcessNo", caseNo);
            //html = html.Replace("@OccurReport34_2", row["OccurReport34_2"]);
            html = html.Replace("@OccurReport34_2", row["OccurReport34_2"].Equals("") ? NullStr : row["OccurReport34_2"]);

            string sql = "select * from CasePerson where CaseNo = '" + caseNo + "' order by PersonSno ";
            List<Dictionary<string, string>> lsdy = ioCon.SelectDct(sql);
            if (lsdy.Count == 0) return;

            /***********************************************************************************************/
            string tmpfil = filePath + caseNo + "_11.pdf";
            Document doc = new Document(PageSize.A4.Rotate(), 15, 15, 15, 15);
            FileStream fs = new FileStream(tmpfil, FileMode.Create);
            PdfWriter pdf = PdfWriter.GetInstance(doc, fs);
            doc.Open();
            /***********************************************************************************************/
            
            for (int i = 1; i < 5; i++)
            {
                if (i > lsdy.Count)
                {
                    html = html.Replace("@PersonSno" + i, NullStr);
                    html = html.Replace("@PersonName" + i, NullStr);
                    html = html.Replace("@PersonFlag" + i, NullStr);
                    html = html.Replace("@IdNo" + i, NullStr);
                    html = html.Replace("@Birthday" + i, NullStr);
                    html = html.Replace("@Addr" + i, NullStr);
                    html = html.Replace("@Tel" + i, NullStr);
                    html = html.Replace("@Remark" + i, NullStr);
                }
                else
                {
                    Dictionary<string, string> col = lsdy[i - 1];
                    html = html.Replace("@PersonSno" + i, col["PersonSno"]);
                    html = html.Replace("@PersonName" + i, col["PersonCname"].Equals("") ? col["PersonEname"] : col["PersonCname"]);
                    html = html.Replace("@PersonFlag" + i, col["PersonFlag"]);
                    html = html.Replace("@IdNo" + i, col["IdNo"]);
                    //html = html.Replace("@Birthday" + i, getDateFrm(col["Birthday"], 4));
                    string dat = col["Birthday"];
                    if (!dat.Trim().Equals("")) dat = dat.Substring(0,3) + "年" + dat.Substring(3,2) + "月" + dat.Substring(5,2) + "日";
                    html = html.Replace("@Birthday" + i, dat);
                    html = html.Replace("@Addr" + i, col["Addr"]);
                    html = html.Replace("@Tel" + i, col["Tel2"] + "  " + col["Tel1"]);
                    html = html.Replace("@Remark" + i, col["Remark"]);
                }
            }

            string[] ss = new string[4] { "a", "b", "c", "d" };
            for (int i = 0; i < 4; i++)
            {
                if (i >= lsdy.Count)
                {
                    html = html.Replace("@OccurReport22" + ss[i], NullStr);
                    html = html.Replace("@OccurReport23" + ss[i], NullStr);
                    html = html.Replace("@OccurReport24" + ss[i], NullStr);
                    html = html.Replace("@OccurReport25" + ss[i], NullStr);
                    html = html.Replace("@OccurReport26" + ss[i], NullStr);
                    html = html.Replace("@OccurReport27" + ss[i], NullStr);
                    html = html.Replace("@OccurReport28" + ss[i], NullStr);
                    html = html.Replace("@OccurReport29" + ss[i], NullStr);
                    html = html.Replace("@OccurReport30" + ss[i], NullStr);
                    html = html.Replace("@OccurReport31" + ss[i], NullStr);
                    html = html.Replace("@OccurReport32" + ss[i], NullStr);
                    html = html.Replace("@OccurReport33_1" + ss[i], NullStr);
                    html = html.Replace("@OccurReport33_2" + ss[i], NullStr);
                    html = html.Replace("@OccurReport34_1" + ss[i], NullStr);
                    html = html.Replace("@OccurReport35" + ss[i], NullStr);
                    html = html.Replace("@OccurReport36" + ss[i], NullStr);
                    html = html.Replace("@OccurReport37" + ss[i], NullStr);
                }
                else
                {
                    Dictionary<string, string> col = lsdy[i];
                    html = html.Replace("@OccurReport22" + ss[i], col["OccurReport22"].Equals("") ? NullStr : col["OccurReport22"]);
                    html = html.Replace("@OccurReport23" + ss[i], col["OccurReport23"].Equals("") ? NullStr : col["OccurReport23"]);
                    html = html.Replace("@OccurReport24" + ss[i], col["OccurReport24"].Equals("") ? NullStr : col["OccurReport24"]);
                    html = html.Replace("@OccurReport25" + ss[i], col["OccurReport25"].Equals("") ? NullStr : col["OccurReport25"]);
                    html = html.Replace("@OccurReport26" + ss[i], col["OccurReport26"].Equals("") ? NullStr : col["OccurReport26"]);
                    html = html.Replace("@OccurReport27" + ss[i], col["OccurReport27"].Equals("") ? NullStr : col["OccurReport27"]);
                    html = html.Replace("@OccurReport28" + ss[i], col["OccurReport28"].Equals("") ? NullStr : col["OccurReport28"]);
                    html = html.Replace("@OccurReport29" + ss[i], col["OccurReport29"].Equals("") ? NullStr : col["OccurReport29"]);
                    html = html.Replace("@OccurReport30" + ss[i], col["OccurReport30"].Equals("") ? NullStr : col["OccurReport30"]);
                    html = html.Replace("@OccurReport31" + ss[i], col["OccurReport31"].Equals("") ? NullStr : col["OccurReport31"]);
                    html = html.Replace("@OccurReport32" + ss[i], col["OccurReport32"].Equals("") ? NullStr : col["OccurReport32"]);
                    html = html.Replace("@OccurReport33_1" + ss[i], col["OccurReport33_1"].Equals("") ? NullStr : col["OccurReport33_1"]);
                    html = html.Replace("@OccurReport33_2" + ss[i], col["OccurReport33_2"].Equals("") ? NullStr : col["OccurReport33_2"]);
                    html = html.Replace("@OccurReport34_1" + ss[i], col["OccurReport34_1"].Equals("") ? NullStr : col["OccurReport34_1"]);
                    html = html.Replace("@OccurReport35" + ss[i], col["OccurReport35"].Equals("") ? NullStr : col["OccurReport35"]);
                    html = html.Replace("@OccurReport36" + ss[i], col["OccurReport36"].Equals("") ? NullStr : col["OccurReport36"]);
                    html = html.Replace("@OccurReport37" + ss[i], col["OccurReport37"].Equals("") ? NullStr : col["OccurReport37"]);
                    //html = html.Replace("@OccurReport23" + ss[i], col["OccurReport22"]);
                    //html = html.Replace("@OccurReport23" + ss[i], col["OccurReport23"]);
                    //html = html.Replace("@OccurReport24" + ss[i], col["OccurReport24"]);
                    //html = html.Replace("@OccurReport25" + ss[i], col["OccurReport25"]);
                    //html = html.Replace("@OccurReport26" + ss[i], col["OccurReport26"]);
                    //html = html.Replace("@OccurReport27" + ss[i], col["OccurReport27"]);
                    //html = html.Replace("@OccurReport28" + ss[i], col["OccurReport28"]);
                    //html = html.Replace("@OccurReport29" + ss[i], col["OccurReport29"]);
                    //html = html.Replace("@OccurReport30" + ss[i], col["OccurReport30"]);
                    //html = html.Replace("@OccurReport31" + ss[i], col["OccurReport31"]);
                    //html = html.Replace("@OccurReport32" + ss[i], col["OccurReport32"]);
                    //html = html.Replace("@OccurReport33_1" + ss[i], col["OccurReport33_1"]);
                    //html = html.Replace("@OccurReport33_2" + ss[i], col["OccurReport33_2"]);
                    //html = html.Replace("@OccurReport34_1" + ss[i], col["OccurReport34_1"]);
                    //html = html.Replace("@OccurReport35" + ss[i], col["OccurReport35"]);
                    //html = html.Replace("@OccurReport36" + ss[i], col["OccurReport36"]);
                    //html = html.Replace("@OccurReport37" + ss[i], col["OccurReport37"]);
                }
            }

            string gifx = "";
            for (int i = 16; i < 38; i++)
            {
                gifx = GifPath + "no" + i + ".gif";
                html = html.Replace("@no" + i.ToString("00"), gifx);
            }
            string car = GifPath + "car.gif";
            html = html.Replace("@car", car);

            string mark = getPsnMark(row["ProcessId"]);
            html = html.Replace("@Mark1", MarkPath + mark);
            mark = getChfMark(row["CaseNo"]);
            html = html.Replace("@Mark2", MarkPath + mark);
            mark = getUntMark(row["OccurPoliceStation"], row["OccurPoliceUnit"]);
            html = html.Replace("@Mark3", MarkPath + mark);

            string codeMsg = caseNo + ";" + row["OccurPoliceUnit"] + ";" + row["OccurPoliceStation"] + ";";
            string qrCode = RQCodePath + caseNo + ".png";
            codeMsg += row["ProcessId"] + ";10.116.1.219;" + addDate(0);
            makeQRcode(codeMsg, qrCode);

            html = html.Replace("@QRCode", qrCode);

            /***********************************************************************************************/
            byte[] bytehtml = System.Text.Encoding.UTF8.GetBytes(html);
            MemoryStream mshtml = new MemoryStream(bytehtml);
            byte[] bytecss = System.Text.Encoding.UTF8.GetBytes(cssStr);
            MemoryStream mscss = new MemoryStream(bytecss);

            XMLWorkerHelper.GetInstance().ParseXHtml(pdf, doc, mshtml, mscss);

            PdfDestination pdfDest = new PdfDestination(PdfDestination.XYZ, 0, doc.PageSize.Height, 1f);
            PdfAction action = PdfAction.GotoLocalPage(1, pdfDest, pdf);
            pdf.SetOpenAction(action);

            doc.Close();
            //EncryptPDF(tmpfil, filePath + row["ApplyNo"] + "_1.pdf", row["IdNo"]);
            //if (row["MailFlag1"] != "1") UpdateMailFlag(row["ApplyNo"], "1");
        }

        static string getPsnMark(string cod)
        {
            string rtn = "";
            string sql = "select MarkPath from Police where IdNo = '" + cod + "'";
            List<Dictionary<string, string>> lsdy = ioCon.SelectDct(sql);
            if (lsdy.Count > 0) rtn = lsdy[0]["MarkPath"];

            return rtn;
        }

        static string getChfMark(string cod)
        {
            string rtn = "";
            string sql = "select MarkPath,AgentMarkPath,AuthMarkPath from Police where IdNo = ";
            sql += "(select casestatu3id from CaseStatuRec where CaseNo = '" + cod + "')";
            List<Dictionary<string, string>> lsdy = ioCon.SelectDct(sql);
            if (lsdy.Count > 0) rtn = lsdy[0]["MarkPath"];

            return rtn;
        }

        static string getCh12IdMark(string cod)
        {
            string rtn = "";
            string sql = "select MarkPath,AgentMarkPath,AuthMarkPath from Police where IdNo = ";
            sql += "(select CaseStatu12Id from CaseStatuRec where CaseNo = '" + cod + "')";
            List<Dictionary<string, string>> lsdy = ioCon.SelectDct(sql);
            if (lsdy.Count > 0) rtn = lsdy[0]["MarkPath"];

            return rtn;
        }

        static string getUntMark(string cod1, string cod2)
        {
            string rtn = "";
            string sql = "select MarkPath from PoliceStation where PStationId = '" + cod1 + "'";
            List<Dictionary<string, string>> lsdy = ioCon.SelectDct(sql);
            if (lsdy.Count > 0) return lsdy[0]["MarkPath"];

            sql = "select MarkPath from PoliceUnit where PUnitId = '" + cod2 + "'";
            if (lsdy.Count > 0) rtn = lsdy[0]["MarkPath"];

            return rtn;
        }

        static string getUntMarkh(string cod1, string cod2)
        {
            string rtn = "";
            string sql = "select MarkPath from PoliceStation where PStationId = '" + cod1 + "'";
            List<Dictionary<string, string>> lsdy = ioCon.SelectDct(sql);
            if (lsdy.Count > 0)
            {
                rtn = lsdy[0]["MarkPath"];
                rtn = rtn.Substring(0, rtn.Length - 4) + "-h" + rtn.Substring(rtn.Length - 4);
                return rtn;
            }

            sql = "select MarkPath from PoliceUnit where PUnitId = '" + cod2 + "'";
            if (lsdy.Count > 0) rtn = lsdy[0]["MarkPath"];
            rtn = rtn.Substring(0, rtn.Length - 4) + "-h" + rtn.Substring(rtn.Length - 4);

            return rtn;
        }

        static string getUnit(string cod)
        {
            string rtn = "";
            string sql = "select * from PoliceUnit ";
            sql += "where PUnitId='" + cod + "' ";
            List<Dictionary<string, string>> lsdy = ioCon.SelectDct(sql);
            if (lsdy.Count > 0) rtn = lsdy[0]["PUnitName"];

            return rtn;
        }

        static string getMItemVal(string oType, string oVal, int rkind)
        {
            string rtn = "";
            string sql = "select mItem,mKind from MenuItem ";
            sql += "where mType='" + oType + "' and mVal='" + oVal + "' ";
            List<Dictionary<string, string>> lsdy = ioCon.SelectDct(sql);
            if (lsdy.Count > 0)
            {
                rtn = rkind == 0 ? lsdy[0]["mItem"] : lsdy[0]["mKind"];
                switch (rkind)
                {
                    case 0:
                        rtn = lsdy[0]["mItem"];
                        break;
                    case 1:
                        rtn = lsdy[0]["mKind"];
                        break;
                    case 2:
                        rtn = lsdy[0]["mItem"] + lsdy[0]["mKind"];
                        break;
                }
            }

            return rtn;
        }

        static string addDate(int dy)
        {
            DateTime sysdate = DateTime.Now;
            string aftdate = sysdate.AddDays(dy).ToString("yyyyMMdd");
            int yy = Convert.ToInt32(aftdate.Substring(0, 4)) - 1911;
            string rtn = yy < 100 ? "0" + yy : "" + yy;
            rtn += aftdate.Substring(4);
            return rtn;
        }

        static string getDateFrm(string dat, int ty)
        {
            if (dat.Length < 7) return "";
            string rtn = "";
            string yy, mm, dd, hh, ms;
            switch (ty)
            {
                case 0:
                    yy = dat.Substring(0, 4);
                    mm = dat.Substring(4, 2);
                    dd = dat.Substring(6, 2);
                    rtn = (Convert.ToInt32(yy) - 1911) + " 年 ";
                    rtn += mm + " 月 " + dd + " 日";
                    break;
                case 1:
                    yy = dat.Substring(0, 3);
                    mm = dat.Substring(3, 2);
                    dd = dat.Substring(5, 2);
                    hh = dat.Substring(7, 2);
                    ms = dat.Substring(9, 2);
                    rtn = yy + "年" + mm + "月" + dd + "日<br/>";
                    rtn += hh + "時" + ms + "分";
                    break;
                case 2:
                    yy = dat.Substring(0, 3);
                    mm = dat.Substring(3, 2);
                    dd = dat.Substring(5, 2);
                    hh = dat.Substring(7, 2);
                    ms = dat.Substring(9, 2);
                    rtn = yy + " 年 " + mm + " 月 " + dd + " 日 ";
                    rtn += hh + " 時 " + ms + " 分 ";
                    break;
                case 3:
                    yy = dat.Substring(0, 3);
                    mm = dat.Substring(3, 2);
                    dd = dat.Substring(5, 2);
                    hh = dat.Substring(7, 2);
                    ms = dat.Substring(9, 2);
                    rtn = yy + "/" + mm + "/" + dd + " " + hh + ":" + ms;
                    break;
                case 4:
                    yy = dat.Substring(0, 3);
                    mm = dat.Substring(3, 2);
                    dd = dat.Substring(5, 2);
                    rtn = yy + " 年 " + mm + " 月 " + dd + " 日 ";
                    break;
            }
            return rtn;
        }

        static string getPoint(string cod)
        {
            string rtn = "";
            switch (cod)
            {
                case "東":
                    rtn = "→";
                    break;
                case "南":
                    rtn = "↓";
                    break;
                case "西":
                    rtn = "←";
                    break;
                case "北":
                    rtn = "↑";
                    break;
                case "東南":
                    rtn = "↘";
                    break;
                case "東北":
                    rtn = "↗";
                    break;
                case "西南":
                    rtn = "↙";
                    break;
                case "西北":
                    rtn = "↖";
                    break;
            }

            return rtn;
        }

        static string getAAScp(string cod)
        {
            string rtn = "";
            switch (cod)
            {
                case "2":
                    rtn = "無照(未達考照年齡)";
                    break;
                case "3":
                    rtn = "無照(已達考照年齡)";
                    break;
                case "4":
                    rtn = "越級駕駛";
                    break;
                case "5":
                    rtn = "駕照被吊扣";
                    break;
                case "6":
                    rtn = "駕照被吊(註)銷";
                    break;
            }

            return rtn;
        }

        static string getBBScp(string cod)
        {
            string rtn = "";
            switch (cod)
            {
                case "03":
                    rtn = "經呼氣檢測未超過 0.15 mg/L或血液檢測未超過 0.03%";
                    break;
                case "04":
                    rtn = "經呼氣檢測 0.16~0.25 mg/L或血液檢測 0.031%~0.05%";
                    break;
                case "05":
                    rtn = "經呼氣檢測 0.26~0.40 mg/L或血液檢測 0.051%~0.08%";
                    break;
                case "06":
                    rtn = "經呼氣檢測 0.41~0.55 mg/L或血液檢測 0.081%~0.11%";
                    break;
                case "07":
                    rtn = "經呼氣檢測 0.56~0.80 mg/L或血液檢測 0.111%~0.16%";
                    break;
                case "08":
                    rtn = "經呼氣檢測超過 0.80~ mg/L或血液檢測超過 0.16%";
                    break;
            }

            return rtn;
        }

        static string getLigthScp(string cod)
        {
            string rtn = "";
            switch (cod)
            {
                case "00": rtn = "無號誌"; break;
                case "01": rtn = "普通二時相"; break;
                case "02": rtn = "早開二時相"; break;
                case "03": rtn = "遲閉二時相"; break;
                case "04": rtn = "輪放式三時相"; break;
                case "05": rtn = "左轉保護三時相"; break;
                case "06": rtn = "輪放式四時相"; break;
                case "07": rtn = "左轉保護四時相"; break;
                case "08": rtn = "輪放左轉保護時相"; break;
                case "09": rtn = "行人保護三時相"; break;
                case "10": rtn = "匝道儀控"; break;
                case "11": rtn = "閃光"; break;
                case "12": rtn = "其他"; break;
                case "13": rtn = "不正常運轉或無動作"; break;
            }

            return rtn;
        }

        static string getAddr(Dictionary<string, string> row)
        {
            string rtn = "";
            if (row["OccurAddr1_1"] != "") rtn += row["OccurAddr1_1"];
            if (row["OccurAddr1_2"] != "") rtn += row["OccurAddr1_2"];
            if (row["OccurAddr1_3"] != "") rtn += row["OccurAddr1_3"];
            if (row["OccurAddr1_4"] != "") rtn += row["OccurAddr1_4"] + "鄰";
            if (row["OccurAddr1_5"] != "") rtn += row["OccurAddr1_5"];
            if (row["OccurAddr1_6"] != "") rtn += row["OccurAddr1_6"] + "段";
            if (row["OccurAddr1_7"] != "") rtn += row["OccurAddr1_7"] + "巷";
            if (row["OccurAddr1_8"] != "") rtn += row["OccurAddr1_8"] + "弄";
            if (row["OccurAddr1_9"] != "") rtn += row["OccurAddr1_9"] + "號" + "前";
            if (row["OccurAddr1_10"] != "") rtn += row["OccurAddr1_10"] + "公尺處";
            if (row["OccurAddr1_11"] != "")
            {
                rtn += "與" + row["OccurAddr1_11"];
                if (row["OccurAddr1_11_1"] != "")
                {
                    rtn += row["OccurAddr1_11_1"] + "段" + "(口)";
                }
                else
                {
                    rtn += "(口)";
                }
            }
            else
            {
                if (row["OccurAddr1_11_1"] != "") rtn += row["OccurAddr1_11_1"] + "段" + "(口)";
            }

            if (row["OccurAddr1_12"] != "") rtn += row["OccurAddr1_12"] + "側";
            if (row["OccurAddr1_13"] != "") rtn += row["OccurAddr1_13"];
            if (row["OccurAddr2_1"] != "")
            {
                rtn += "<br/>";
                if (row["OccurAddr2_1"] != "") rtn += row["OccurAddr2_1"];
                if (row["OccurAddr2_2"] != "") rtn += row["OccurAddr2_2"] + "公里";
                if (row["OccurAddr2_3"] != "") rtn += row["OccurAddr2_3"] + "公尺處";
                if (row["OccurAddr2_4"] != "") rtn += row["OccurAddr2_4"] + "向";
                if (row["OccurAddr2_5"] != "") rtn += row["OccurAddr2_5"] + "車道";
            }
            if (row["OccurAddr3_1"] != "")
            {
                rtn += "<br/>";
                if (row["OccurAddr3_1"] != "") rtn += row["OccurAddr3_1"] + "線";
                if (row["OccurAddr3_2"] != "") rtn += row["OccurAddr3_2"] + "公里";
                if (row["OccurAddr3_3"] != "") rtn += row["OccurAddr3_3"] + "公尺處";
                if (row["OccurAddr3_4"] != "") rtn += row["OccurAddr3_4"] + "平交道";
            }

            return rtn;
        }

        static string getScp(Dictionary<string, string> row)
        {
            string rtn = "";

            rtn += row["DocFld2"] == "1" ? "■" : "□";
            rtn += "道路全景";
            rtn += row["DocFld3"] == "1" ? " ■" : " □";
            rtn += "車損";
            rtn += row["DocFld4"] == "1" ? " ■" : " □";
            rtn += "車體擦痕";
            rtn += row["DocFld5"] == "1" ? " ■" : " □";
            rtn += "機車倒地";
            rtn += row["DocFld6"] == "1" ? " ■" : " □";
            rtn += "煞車痕";
            rtn += row["DocFld7"] == "1" ? " ■" : " □";
            rtn += "刮地痕";
            rtn += row["DocFld8"] == "1" ? " ■" : " □";
            rtn += "拖痕";
            rtn += "<br/>";
            rtn += row["DocFld9"] == "1" ? " ■" : " □";
            rtn += "道路設施";
            rtn += row["DocFld10"] == "1" ? " ■" : " □";
            rtn += "人倒地";
            rtn += row["DocFld11"] == "1" ? " ■" : " □";
            rtn += "人受傷部位";
            rtn += row["DocFld12"] == "1" ? " ■" : " □";
            rtn += "落土";
            rtn += row["DocFld14"] == "1" ? " ■" : " □";
            rtn += "碎片";
            rtn += row["DocFld15"] == "1" ? " ■" : " □";
            rtn += "其他 ";
            rtn += row["DocFld16"];

            return rtn;
        }

        static string getPer(string cod1, string cod2)
        {
            string rtn = "";

            rtn += cod1 == "1" ? "■" : "□";
            rtn += "草圖&nbsp;&nbsp;&nbsp;&nbsp;";
            rtn += cod1 == "2" ? "&nbsp;■" : "&nbsp;□";
            rtn += "1公尺(M)<br/>";
            rtn += cod1 == "3" ? "■" : "□";
            rtn += "2公尺(M)";
            rtn += cod1 == "4" ? "&nbsp;■" : "&nbsp;□";
            rtn += "其他&nbsp;";
            rtn += cod2 == "" ? "__" : cod2 + "公尺";

            return rtn;
        }

        static string getSkip(string str, int ln)
        {
            string rtn = "";
            string tmp = str;

            while (true)
            {
                if (tmp.Length > ln)
                {
                    rtn += tmp.Substring(0, ln) + "<br/>";
                    tmp = tmp.Substring(ln);
                }
                else
                {
                    rtn += tmp;
                    break;
                }
            }

            return rtn;
        }

        static void EncryptPDF(string SrcPath, string OutPath, string passwd)
        {
            using (PdfReader reader = new PdfReader(SrcPath))
            {
                using (var os = new FileStream(OutPath, FileMode.Create))
                {
                    PdfEncryptor.Encrypt(reader, os, true, passwd, passwd, PdfWriter.ALLOW_SCREENREADERS);
                }
            }
        }
        
        static void UpdateMailFlag(string AppNo,string ii)
        {
            string sql = "update PersonApply set MailFlag" + ii + " = '2' where ApplyNo = '" + AppNo + "'";
            ioCon.ExecuteSql(sql);
        }

        static void makeQRcode(string msg,string fpath)
        {
            BarcodeWriter bw = new BarcodeWriter();
            bw.Format = BarcodeFormat.QR_CODE;
            bw.Options.Width = 200;
            bw.Options.Height = 200;
            Bitmap bitmap = bw.Write(msg);
            bitmap.Save(fpath, ImageFormat.Png);
        }

        static void recLog(string msg)
        {
            string sysdatim = DateTime.Now.ToString("yyyyMMddHHmm");
            if (!Directory.Exists(LogPath)) Directory.CreateDirectory(LogPath);
            string LogFile = LogPath + sysdatim.Substring(0, 6) + "\\";
            if (!Directory.Exists(LogFile)) Directory.CreateDirectory(LogFile);
            StreamWriter logwt = File.AppendText(LogFile + sysdatim.Substring(0, 8) + ".log");
            logwt.WriteLine(msg);
            logwt.Close();

        }

        static void deleteTmp()
        {
            FileInfo[] fils;
            DirectoryInfo dir = new DirectoryInfo(filePath);
            fils = dir.GetFiles("tmp*.pdf");

            foreach (FileInfo fi in fils)
            {
                string filn = filePath + fi.Name;
                File.Delete(filn);
            }

        }

        static string getCWeek(string iCDate)
        {
            string Ret = "";

            String yy = iCDate.Substring(0, 3);
            String mm = iCDate.Substring(3, 2);
            String dd = iCDate.Substring(5, 2);
            int iyy = int.Parse(yy);
            iyy += 1911;
            int imm = int.Parse(mm);
            int idd = int.Parse(dd);

            CultureInfo m_ciTaiwan = new CultureInfo("zh-TW");
            DateTime dateValue = new DateTime(iyy,imm,idd);

            Ret = dateValue.ToString("dddd", m_ciTaiwan);

            return Ret;
        }



    }
}
