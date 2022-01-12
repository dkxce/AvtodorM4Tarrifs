using System;
using System.Collections.Generic;
using System.Web;
using System.Net;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Serialization;
using System.Windows;
using System.Windows.Forms;

namespace avtodor_tr_m4
{
    public static class Program
    {
        public static string version = "12.01.2022";

        public static void Main()
        {            
            Console.WriteLine("Avtodor-Tr M4 Grabber by milokz@gmail.com");
            Console.WriteLine("** version " + version + " **");
            Console.WriteLine("...");
            Console.WriteLine("Enter URL in Dialog Window...");


            string url = "https://avtodor-tr.ru/ru/platnye-uchastki/fares/m4/";
            if (InputBox.Show("Grab avtodor-tr.ru prices", "URL:", ref url) != DialogResult.OK) return;

            Console.WriteLine("Grabbing data with wGet ...");
            Console.WriteLine("  " + url);
            HttpWGetRequest wGet = new HttpWGetRequest(url);
            wGet.RemoteEncoding = Encoding.UTF8;
            wGet.LocalEncoding = Encoding.UTF8;
            string body = wGet.GetResponseBody();
            if (String.IsNullOrEmpty(body))
            {
                Console.WriteLine("No data ...");
                MessageBox.Show("No DATA");
                return;
            };
            Console.WriteLine("Grabbed " + body.Length + " symbols");
            Console.WriteLine("Parsing data ... ");

            Regex tableMatch = new Regex(@"(<table[\S\s]*?\/table>)", RegexOptions.IgnoreCase);
            Regex rowMatch = new Regex(@"(\<tr[\S\s]*?\<\/tr\>)", RegexOptions.IgnoreCase);
            Regex cellMatch = new Regex(@"(<td[\S\s]*?</td>)", RegexOptions.IgnoreCase);
            Regex rowsSpanMatch = new Regex(@"rowspan=""(?<rowspan>\d)""", RegexOptions.IgnoreCase);
            Regex colspanMatch = new Regex(@"colspan=""(?<colspan>\d)""", RegexOptions.IgnoreCase);
            
            // GET TABLE ROWS
            List<string> tlines = new List<string>();
            foreach (Match tm in tableMatch.Matches(body))
            {
                if (tm.Value.IndexOf("—ÚÓËÏÓÒÚ¸ ÔÓÂÁ‰‡") < 0) continue;
                foreach (Match rm in rowMatch.Matches(tm.Value))
                    tlines.Add(rm.Value);
            };

            AvtodorTRWeb.Segments segments = new AvtodorTRWeb.Segments();
            AvtodorTRWeb.PayWays payways = new AvtodorTRWeb.PayWays();
            payways.url = url;
            payways.grabbed = DateTime.Now;

            if (true) // NEW VERSION 22.01.2021
            {                
                string FIND_CAT1 = "1 copy.";
                string FIND_CAT2 = "2 copy.";
                string FIND_CAT3 = "3 copy.";
                string FIND_CAT4 = "4 copy.";
                string FIND_CASH = "cash.";
                string FIND_PASS = "tpass.";
                string FIND_ZONE = "ÍÏ";
                string FIND_DAYT = "sun.";
                string FIND_NIGT = "night.";
                string ZONE_NAME = "œ¬œ {0}";

                // FILL GRID WITH NO ROWSPAN & NO COLSPAN
                uint currRow = 0; uint currCol = 0;
                AvtodorTRWeb.TabledData GRID = new AvtodorTRWeb.TabledData();                
                for (int l = 0; l < tlines.Count; l++) // each table row
                {
                    currRow++; currCol = 1;
                    MatchCollection cellcoll = cellMatch.Matches(tlines[l]);
                    foreach (Match ma in cellcoll) // each col
                    {
                        while (!String.IsNullOrEmpty(GRID.Get(currRow, currCol).html)) currCol++; // if rowspan

                        string html = ma.Value;                        
                        int rowspan = rowsSpanMatch.Match(html).Success ? int.Parse(rowsSpanMatch.Match(html).Groups["rowspan"].Value) : 1;
                        int colspan = colspanMatch.Match(html).Success ? int.Parse(colspanMatch.Match(html).Groups["colspan"].Value) : 1;
                        string txt = StripHTML(html).Replace("\r", " ").Replace("\n", " ").Trim('\t').Trim('\n').Trim('\r').Replace("  ", " ").Trim();

                        for (int ir = (int)currRow; ir < currRow + rowspan; ir++)
                            for (int ic = (int)currCol; ic < currCol + colspan; ic++)
                                GRID.Set((uint)ir, (uint)ic, txt, html);

                        currCol += (uint)colspan;
                    };
                };              

                // ANALYSE GRID
                bool headers = true;
                Dictionary<int, int> CAT_LIST = new Dictionary<int, int>(); // category by col
                Dictionary<int, string> DAY_LIST = new Dictionary<int, string>(); // days by col
                Dictionary<int, int> CASH_TPASS_LIST = new Dictionary<int, int>(); // cash or tpass by col  
                {
                    for (int ir = GRID.MinX; ir <= GRID.MaxX; ir++)
                    {
                        // IF EMPTY CELL
                        if (String.IsNullOrEmpty(GRID.Get(ir, 1).html)) continue;

                        // GET GRID HEADERS
                        if(headers)
                        {
                            for (int ic = GRID.MinY; ic <= GRID.MaxY; ic++)
                            {
                                string hHTML = GRID.Get(ir, ic).html;
                                if (String.IsNullOrEmpty(hHTML)) continue;

                                if (hHTML.Contains(FIND_CAT1)) CAT_LIST.Add(ic, 1);
                                if (hHTML.Contains(FIND_CAT2)) CAT_LIST.Add(ic, 2);
                                if (hHTML.Contains(FIND_CAT3)) CAT_LIST.Add(ic, 3);
                                if (hHTML.Contains(FIND_CAT4)) CAT_LIST.Add(ic, 4);
                                if (hHTML.Contains(FIND_CASH)) CASH_TPASS_LIST.Add(ic, 0);
                                if (hHTML.Contains(FIND_PASS)) CASH_TPASS_LIST.Add(ic, 1);

                                string hTEXT = GRID.Get(ir, ic).text;
                                if (String.IsNullOrEmpty(hTEXT)) continue;

                                bool exd = false;
                                foreach(string dof in AvtodorTRWeb.Tarrif.DaysOfWeek) 
                                    if (hTEXT.ToUpper().Contains(dof)) 
                                        { exd = true; break; };
                                if (exd && (!DAY_LIST.ContainsKey(ic))) DAY_LIST.Add(ic, hTEXT);
                            };
                        };

                        // GET GRID COSTS
                        string cellText = GRID.Get(ir, 1).text;
                        if (Char.IsDigit(cellText[0]) && cellText.Contains(FIND_ZONE))
                        {
                            headers = false; // no more headers

                            // GET NAME AND ID
                            AvtodorTRWeb.PayWay pw = new AvtodorTRWeb.PayWay();
                            pw.SetName(String.Format(ZONE_NAME, cellText), segments);

                            if (String.IsNullOrEmpty(pw.ID))
                            {
                                string nextCellText = GRID.Get(ir, 2).text;
                                if (!String.IsNullOrEmpty(nextCellText))
                                {
                                    Match msub = (new Regex(@"[\d][^\e]*", RegexOptions.IgnoreCase)).Match(nextCellText);
                                    if (msub.Success)
                                    {
                                        nextCellText = msub.Groups[0].Value;
                                        pw.CheckID(nextCellText, segments);
                                    };
                                };
                            };

                            // PAYWAY EXISTS
                            bool ex = false;
                            foreach (AvtodorTRWeb.PayWay pws in payways.Payways)
                                if (pws.ID == pw.ID)
                                {
                                    ex = true;
                                    pw = pws;
                                    break;
                                };
                            
                            //GET TARRIFS
                            {
                                string time = "";
                                for (int ic = GRID.MinY; ic <= GRID.MaxY; ic++)
                                {
                                    // GET TIME
                                    string cHTML = GRID.Get(ir, ic).html;
                                    if (!String.IsNullOrEmpty(cHTML))
                                    {
                                        if (cHTML.Contains(FIND_DAYT)) time = "DAY";
                                        if (cHTML.Contains(FIND_NIGT)) time = "NIGHT";
                                    };

                                    string cTEXT = GRID.Get(ir, ic).text;
                                    if (String.IsNullOrEmpty(cTEXT)) continue;
                                    if (cTEXT.Contains(FIND_ZONE)) continue;
                                    if (!char.IsDigit(cTEXT[0])) continue;
                                    
                                    int cat = CAT_LIST[ic];
                                    string day = DAY_LIST[ic];
                                    int cashtpass = CASH_TPASS_LIST[ic];
                                    
                                    float cost = 0;
                                    Regex rx = new Regex(@"[\d\.\,]+", RegexOptions.IgnoreCase);
                                    Match mx = rx.Match(cTEXT);
                                    if (mx.Success)
                                    {
                                        string val = mx.Groups[0].Value.Replace(",", ".");
                                        float.TryParse(val, System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out cost);
                                    };

                                    pw.AddCost(day, time, cat, cashtpass, cost);
                                };
                            };

                            // SORT TARRIFS
                            pw.Tarrifs.Sort(new AvtodorTRWeb.TarifSorter());

                            if (!ex) payways.Add(pw);
                        };
                    };
                };
            };

            #region OLD
            if (false) // OLD VERSION 26.04.2021
            {
                for (int l = 0; l < tlines.Count; l++)
                {
                    if ((!tlines[l].Contains("”˜‡ÒÚÓÍ")) && (!tlines[l].Contains("œ¬œ"))) continue;

                    AvtodorTRWeb.PayWay pw = new AvtodorTRWeb.PayWay();
                    AvtodorTRWeb.Tarrif tar = new AvtodorTRWeb.Tarrif();
                    int col;
                    int[] rowspans = new int[] { 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1 };
                    {
                        col = 0;
                        int ub = 0;
                        MatchCollection cellcoll = cellMatch.Matches(tlines[l]);
                        foreach (Match ma in cellcoll)
                        {
                            string cellVal = ma.Value;
                            int rowspan = rowsSpanMatch.Match(cellVal).Success ? int.Parse(rowsSpanMatch.Match(cellVal).Groups["rowspan"].Value) : 1;
                            int colspan = colspanMatch.Match(cellVal).Success ? int.Parse(colspanMatch.Match(cellVal).Groups["colspan"].Value) : 1;
                            if (rowspan > 1) rowspans[col] = rowspan;

                            if (cellVal.Contains("sun.png"))
                                tar.Time = "DAY";
                            if (cellVal.Contains("night.png"))
                                tar.Time = "NIGHT";

                            string txt = StripHTML(cellVal).Replace("\r", " ").Replace("\n", " ").Trim('\t').Trim('\n').Trim('\r').Replace("  ", " ").Trim();

                            if (txt.Contains("”˜‡ÒÚÓÍ"))
                                pw.SetName(txt, segments);
                            else if (txt.Contains(("œ¬œ")))
                                pw.SetName(pw.Name + " (" + txt + ")", segments);
                            else if (txt.Length > 0)
                            {
                                if ((char.IsDigit(txt[0]) || (char.IsDigit(txt[txt.Length - 1]))))
                                    tar.AddCost(txt, ref ub);
                            };
                            col++;
                        };
                        AvtodorTRWeb.Tarrif tar_0 = new AvtodorTRWeb.Tarrif();
                        tar_0.Time = tar.Time;
                        tar_0.Days = "2,3,4,5";
                        tar_0.Costs = new float[] { tar.Costs[0], tar.Costs[1], tar.Costs[4], tar.Costs[5], tar.Costs[6], tar.Costs[7], tar.Costs[8], tar.Costs[9] };
                        AvtodorTRWeb.Tarrif tar_1 = new AvtodorTRWeb.Tarrif();
                        tar_1.Time = tar.Time;
                        tar_1.Days = "6,7,1";
                        tar_1.Costs = new float[] { tar.Costs[2], tar.Costs[3], tar.Costs[4], tar.Costs[5], tar.Costs[6], tar.Costs[7], tar.Costs[8], tar.Costs[9] };
                        pw.Tarrifs.Add(tar_0);
                        pw.Tarrifs.Add(tar_1);
                    };
                    while (rowspans[0] > 1)
                    {
                        rowspans[0]--;
                        l++;
                        tar = tar.Clone();
                        col = 1;
                        int ub = 0;
                        foreach (Match ma in cellMatch.Matches(tlines[l]))
                        {
                            string cellVal = ma.Value;
                            int rowspan = rowsSpanMatch.Match(cellVal).Success ? int.Parse(rowsSpanMatch.Match(cellVal).Groups["rowspan"].Value) : 1;
                            int colspan = colspanMatch.Match(cellVal).Success ? int.Parse(colspanMatch.Match(cellVal).Groups["colspan"].Value) : 1;
                            if (rowspan > 1) rowspans[col] = rowspan;

                            if (cellVal.Contains("sun.png"))
                                tar.Time = "DAY";
                            if (cellVal.Contains("night.png"))
                                tar.Time = "NIGHT";

                            string txt = StripHTML(cellVal).Replace("\r", " ").Replace("\n", " ").Trim('\t').Trim('\n').Trim('\r').Replace("  ", " ").Trim();

                            if (txt.Contains("”˜‡ÒÚÓÍ"))
                                pw.SetName(txt, segments);
                            else if (txt.Contains(("œ¬œ")))
                                pw.Name += " (" + txt + ")";
                            else if (txt.Length > 0)
                            {
                                if ((char.IsDigit(txt[0]) || (char.IsDigit(txt[txt.Length - 1]))))
                                    tar.AddCost(txt, ref ub);
                            };
                            col++;
                        };
                        AvtodorTRWeb.Tarrif tar_0 = new AvtodorTRWeb.Tarrif();
                        tar_0.Time = tar.Time;
                        tar_0.Days = "2,3,4,5";
                        tar_0.Costs = new float[] { tar.Costs[0], tar.Costs[1], tar.Costs[4], tar.Costs[5], tar.Costs[6], tar.Costs[7], tar.Costs[8], tar.Costs[9] };
                        AvtodorTRWeb.Tarrif tar_1 = new AvtodorTRWeb.Tarrif();
                        tar_1.Time = tar.Time;
                        tar_1.Days = "6,7,1";
                        tar_1.Costs = new float[] { tar.Costs[2], tar.Costs[3], tar.Costs[4], tar.Costs[5], tar.Costs[6], tar.Costs[7], tar.Costs[8], tar.Costs[9] };
                        pw.Tarrifs.Add(tar_0);
                        pw.Tarrifs.Add(tar_1);
                    };
                    payways.Add(pw);
                    pw.Tarrifs.Sort(new AvtodorTRWeb.TarifSorter());
                };
            };
            #endregion OLD

            payways.Sort();

            string fname = System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase.ToString().Replace("file:///", "").Replace("/", @"\");
            fname = fname.Substring(0, fname.LastIndexOf(@"\") + 1);
            // SAVE TO DXML
            Console.WriteLine("Saving data to dxml ... ");
            AvtodorTRWeb.PayWays.Save(payways, fname + @"avtodor_tr_m4_DATA.dxml");
            // SAVE TO TXT
            Console.WriteLine("Saving data to txt ... ");
            AvtodorTRWeb.PayWays.Export2Text(payways, fname + @"avtodor_tr_m4_DATA.txt");
            // SAVE TO XML
            Console.WriteLine("Saving data to xml ... ");
            AvtodorTRWeb.PayWays.Export2ExcelXML(payways, fname += @"avtodor_tr_m4_DATA.xml");
            // LAUNCH EXCEL
            if(String.IsNullOrEmpty(System.Environment.GetEnvironmentVariable("MAPTOP")))
            {
                Console.WriteLine("Starting Excel ... ");
                try { System.Diagnostics.Process.Start("excel", "\"" + fname + "\""); }catch { };
            };
            Console.WriteLine("Done! ");            
            System.Threading.Thread.Sleep(5000);
            Console.WriteLine("avtodor_tr_m4_DATA.dxml");
        }

        private static string StripHTML(string source)
        {
            try
            {
                string result;

                // Remove HTML Development formatting
                // Replace line breaks with space
                // because browsers inserts space
                result = source.Replace("\r", " ");
                // Replace line breaks with space
                // because browsers inserts space
                result = result.Replace("\n", " ");
                // Remove step-formatting
                result = result.Replace("\t", string.Empty);
                // Remove repeating spaces because browsers ignore them
                result = System.Text.RegularExpressions.Regex.Replace(result,
                                                                      @"( )+", " ");

                // Remove the header (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*head([^>])*>", "<head>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<( )*(/)( )*head( )*>)", "</head>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(<head>).*(</head>)", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // remove all scripts (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*script([^>])*>", "<script>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<( )*(/)( )*script( )*>)", "</script>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                //result = System.Text.RegularExpressions.Regex.Replace(result,
                //         @"(<script>)([^(<script>\.</script>)])*(</script>)",
                //         string.Empty,
                //         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<script>).*(</script>)", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // remove all styles (prepare first by clearing attributes)
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*style([^>])*>", "<style>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"(<( )*(/)( )*style( )*>)", "</style>",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(<style>).*(</style>)", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert tabs in spaces of <td> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*td([^>])*>", "\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert line breaks in places of <BR> and <LI> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*br( )*>", "\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*li( )*>", "\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // insert line paragraphs (double line breaks) in place
                // if <P>, <DIV> and <TR> tags
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*div([^>])*>", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*tr([^>])*>", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<( )*p([^>])*>", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // Remove remaining tags like <a>, links, images,
                // comments etc - anything that's enclosed inside < >
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"<[^>]*>", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // replace special characters:
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @" ", " ",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&bull;", " * ",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&lsaquo;", "<",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&rsaquo;", ">",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&trade;", "(tm)",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&frasl;", "/",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&lt;", "<",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&gt;", ">",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&copy;", "(c)",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&reg;", "(r)",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove all others. More can be added, see
                // http://hotwired.lycos.com/webmonkey/reference/special_characters/
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         @"&(.{2,6});", string.Empty,
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // for testing
                //System.Text.RegularExpressions.Regex.Replace(result,
                //       this.txtRegex.Text,string.Empty,
                //       System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                // make line breaking consistent
                result = result.Replace("\n", "\r");

                // Remove extra line breaks and tabs:
                // replace over 2 breaks with 2 and over 4 tabs with 4.
                // Prepare first to remove any whitespaces in between
                // the escaped characters and remove redundant tabs in between line breaks
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)( )+(\r)", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\t)( )+(\t)", "\t\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\t)( )+(\r)", "\t\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)( )+(\t)", "\r\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove redundant tabs
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)(\t)+(\r)", "\r\r",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Remove multiple tabs following a line break with just one tab
                result = System.Text.RegularExpressions.Regex.Replace(result,
                         "(\r)(\t)+", "\r\t",
                         System.Text.RegularExpressions.RegexOptions.IgnoreCase);
                // Initial replacement target string for line breaks
                string breaks = "\r\r\r";
                // Initial replacement target string for tabs
                string tabs = "\t\t\t\t\t";
                for (int index = 0; index < result.Length; index++)
                {
                    result = result.Replace(breaks, "\r\r");
                    result = result.Replace(tabs, "\t\t\t\t");
                    breaks = breaks + "\r";
                    tabs = tabs + "\t";
                }

                // That's it.
                return result;
            }
            catch
            {
                MessageBox.Show("Error");
                return source;
            }
        }        
    }

}
namespace AvtodorTRWeb
{
    public class Segments
    {
        public Segments() { this.Load(); }

        public List<Segment> segments = new List<Segment>();
        private void Load()
        {
            this.segments.Clear();
            this.segments.AddRange(XMLSaved<Segment[]>.Load());
        }
        private void Save()
        {
            XMLSaved<Segment[]>.Save(this.segments.ToArray());
        }
    }

    public class Segment
    {
        [XmlAttribute("id")]
        public string ID = String.Empty;
        
        [XmlElement("has")]
        public List<string> Has = new List<string>();

        public Segment() { }

        public Segment(string id)
        {
            this.ID = id;
        }

        public Segment(string id, string has)
        {
            this.ID = id;
            this.Has.Add(has);
        }

        public Segment(string id, string[] has)
        {
            this.ID = id;
            this.Has.AddRange(has);
        }
    }

    public class PayWays
    {
        [XmlElement("url")]
        public string url;
        [XmlElement("grabbed")]
        public DateTime grabbed;
        [XmlElement("createdBy")]
        public string createdBy = "Avtodor-Tr M4 Grabber by milokz@gmail.com"; 

        [XmlElement("payway")]
        public List<PayWay> Payways = new List<PayWay>();

        [XmlIgnore]
        public List<AvtodorTRWeb.Tarrif> MaxTarrif
        {
            get
            {
                if (this.Payways.Count == 0) return null;
                List<AvtodorTRWeb.Tarrif> res = new List<Tarrif>();
                int max = int.MinValue;
                for (int i = 0; i < this.Payways.Count; i++)
                {
                    if (this.Payways[i].Tarrifs.Count <= max) continue;
                    max = this.Payways[i].Tarrifs.Count;
                    res = this.Payways[i].Tarrifs;
                };
                return res;
            }
        }        

        public void Add(PayWay payway) { this.Payways.Add(payway); }
        public void Sort() { this.Payways.Sort(new PayWaySorter()); }
        public List<PayWay>.Enumerator GetEnumerator() { return this.Payways.GetEnumerator(); }
        public static PayWays Load(string fname)
        {
            System.Xml.Serialization.XmlSerializer xs = new System.Xml.Serialization.XmlSerializer(typeof(PayWays));
            System.IO.StreamReader reader = System.IO.File.OpenText(fname);
            PayWays c = (PayWays)xs.Deserialize(reader);
            reader.Close();
            return c;
        }

        public static void Save(PayWays payways, string fname)
        {
            System.Xml.Serialization.XmlSerializer xs = new System.Xml.Serialization.XmlSerializer(typeof(PayWays));
            System.IO.StreamWriter writer = System.IO.File.CreateText(fname);
            xs.Serialize(writer, payways);
            writer.Flush();
            writer.Close();
        }

        public static void Export2Text(PayWays payways, string fname)
        {
            FileStream fs = new FileStream(fname, FileMode.Create, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            sw.WriteLine("ID;SEGMENT;TIME;DAY;CAT1CASH;CAT1TRANSP;CAT2CASH;CAT2TRANSP;CAT3CASH;CAT3TRANSP;CAT4CASH;CAT4TRANSP;DOW;;CASH_LO;CASH_HI;TRANSP_LO;TRANSP_HI");            
            int line = 1;
            string hO = "0", wO = "0";
            string hP = "0", wP = "0";
            string hQ = "0", wQ = "0";
            string hR = "0", wR = "0";
            List<string> payed = new List<string>();
            foreach (AvtodorTRWeb.PayWay pay in payways)
                foreach (AvtodorTRWeb.Tarrif tar in pay.Tarrifs)
                {
                    ++line;
                    sw.Write(String.Format("{0};{1};{2};{3};", pay.ID, pay.Name, tar.TimeRus, tar.DaysRus));
                    sw.Write(String.Format("{0};{1};", tar.GetCost(1, false), tar.GetCost(1, true)));
                    sw.Write(String.Format("{0};{1};", tar.GetCost(2, false), tar.GetCost(2, true)));
                    sw.Write(String.Format("{0};{1};", tar.GetCost(3, false), tar.GetCost(3, true)));
                    sw.Write(String.Format("{0};{1};", tar.GetCost(4, false), tar.GetCost(4, true)));
                    sw.Write(String.Format("{0};;", tar.Days));
                    sw.Write(String.Format("{0};{1};{2};{3};", (line % 2) == 1 ? "=E" + line.ToString() : "", (line % 2) == 0 ? "=E" + line.ToString() : "", (line % 2) == 1 ? "=F" + line.ToString() : "", (line % 2) == 0 ? "=F" + line.ToString() : ""));
                    sw.WriteLine();
                    if (!(String.IsNullOrEmpty(pay.ID)) && (char.IsDigit(pay.ID[pay.ID.Length - 1])) && (!payed.Contains(pay.ID + "_" + (line % 2).ToString())))
                    {
                        if ((line % 2) == 1)
                        {
                            hO += "+O" + line.ToString();
                            hQ += "+Q" + line.ToString();
                            wO += "+O" + (line + 2).ToString();
                            wQ += "+Q" + (line + 2).ToString();
                        }
                        else
                        {
                            hP += "+P" + line.ToString();
                            hR += "+R" + line.ToString();
                            wP += "+P" + (line + 2).ToString();
                            wR += "+R" + (line + 2).ToString();
                        };
                        payed.Add(pay.ID + "_" + (line % 2).ToString());
                    };
                };
            sw.WriteLine(); ++line;
            sw.WriteLine("ALL;;" + payways.MaxTarrif[0].TimeRus + ";" + payways.MaxTarrif[0].DaysRus + ";=O{2};=Q{2};;=E{2}-F{2};;;;;;;=—”ÃÃ({0});;=—”ÃÃ({1});;", hO, hQ, ++line);
            sw.WriteLine("ALL;;" + payways.MaxTarrif[1].TimeRus + ";" + payways.MaxTarrif[1].DaysRus + ";=P{2};=R{2};;=E{2}-F{2};;;;;;;;=—”ÃÃ({0});;=—”ÃÃ({1});", hP, hR, ++line);
            sw.WriteLine("ALL;;" + payways.MaxTarrif[2].TimeRus + ";" + payways.MaxTarrif[2].DaysRus + ";=O{2};=Q{2};;=E{2}-F{2};;;;;;;=—”ÃÃ({0});;=—”ÃÃ({1});;", wO, wQ, ++line);
            sw.WriteLine("ALL;;" + payways.MaxTarrif[3].TimeRus + ";" + payways.MaxTarrif[3].DaysRus + ";=P{2};=R{2};;=E{2}-F{2};;;;;;;;=—”ÃÃ({0});;=—”ÃÃ({1});", wP, wR, ++line);
            sw.WriteLine();
            sw.WriteLine(String.Format("Grabbed at {0} from {1}", payways.grabbed.ToString("HH:mm:ss dd.MM.yyyy"), payways.url));
            sw.WriteLine(String.Format("ƒ‡ÌÌ˚Â ÔÓÎÛ˜ÂÌ˚: {0}", payways.grabbed.ToString("dd.MM.yyyy HH:mm")));
            sw.WriteLine(String.Format("Created by: {0} (version {1})", "Avtodor-Tr M4 Grabber by milokz@gmail.com", avtodor_tr_m4.Program.version));
            sw.Close();
            fs.Close();
        }

        public static void Export2ExcelXML(PayWays payways, StreamWriter sw)
        {
            sw.WriteLine("  <Style ss:ID=\"x0\"><Alignment ss:Vertical=\"Center\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>");
            sw.WriteLine("  <Style ss:ID=\"x1\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/><Font ss:Bold=\"1\"/><Borders><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>");
            sw.WriteLine("  <Style ss:ID=\"x2\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/><Font ss:Bold=\"1\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"2\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>");
            sw.WriteLine("  <Style ss:ID=\"col0\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>");
            sw.WriteLine("  <Style ss:ID=\"col1r\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/><Interior ss:Color=\"#FFB0FF\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>");
            sw.WriteLine("  <Style ss:ID=\"col2r\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/><Interior ss:Color=\"#FFB066\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>");
            sw.WriteLine("  <Style ss:ID=\"col1b\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/><Interior ss:Color=\"#FFCCFF\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>");
            sw.WriteLine("  <Style ss:ID=\"col2b\"><Alignment ss:Horizontal=\"Center\" ss:Vertical=\"Center\"/><Interior ss:Color=\"#FFCC66\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>");
            sw.WriteLine("  <Style ss:ID=\"nor1r\"><Alignment ss:Vertical=\"Center\" ss:WrapText=\"1\"/><Interior ss:Color=\"#FFCCFF\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>");
            sw.WriteLine("  <Style ss:ID=\"nor2r\"><Alignment ss:Vertical=\"Center\" ss:WrapText=\"1\"/><Interior ss:Color=\"#FFCC66\" ss:Pattern=\"Solid\"/><Borders><Border ss:Position=\"Bottom\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Left\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Right\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/><Border ss:Position=\"Top\" ss:LineStyle=\"Continuous\" ss:Weight=\"1\"/></Borders></Style>");
            sw.WriteLine("</Styles>");
            sw.WriteLine("<Worksheet ss:Name=\"Costs\">\r\n<Table>");
            sw.WriteLine("  <Column ss:Width=\"30\"/>");
            sw.WriteLine("  <Column ss:Width=\"180\"/>");
            sw.WriteLine("  <Column ss:Width=\"64\"/>");
            sw.WriteLine("  <Column ss:Width=\"42\"/>");
            sw.WriteLine("  <Column ss:Width=\"50\"/>");
            sw.WriteLine("  <Column ss:Width=\"50\"/>");
            sw.WriteLine("  <Column ss:Width=\"50\"/>");
            sw.WriteLine("  <Column ss:Width=\"50\"/>");
            sw.WriteLine("  <Column ss:Width=\"50\"/>");
            sw.WriteLine("  <Column ss:Width=\"50\"/>");
            sw.WriteLine("  <Column ss:Width=\"50\"/>");
            sw.WriteLine("  <Column ss:Width=\"50\"/>");
            sw.WriteLine("  <Column ss:Width=\"70\"/>");
            sw.WriteLine("  <Column ss:Width=\"20\"/>");
            sw.WriteLine("  <Column ss:Width=\"50\"/>");
            sw.WriteLine("  <Column ss:Width=\"50\"/>");
            sw.WriteLine("  <Column ss:Width=\"50\"/>");
            sw.WriteLine("  <Column ss:Width=\"50\"/>");
            sw.WriteLine("  <Row>");
            sw.WriteLine("    <Cell ss:StyleID=\"x1\" ss:MergeAcross=\"1\"><Data ss:Type=\"String\">œÎ‡ÚÌ˚È Û˜‡ÒÚÓÍ</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x1\" ss:MergeAcross=\"1\"><Data ss:Type=\"String\">œ≈–»Œƒ</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x1\" ss:MergeAcross=\"1\"><Data ss:Type=\"String\"> ¿“≈√Œ–»ﬂ 1</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x1\" ss:MergeAcross=\"1\"><Data ss:Type=\"String\"> ¿“≈√Œ–»ﬂ 2</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x1\" ss:MergeAcross=\"1\"><Data ss:Type=\"String\"> ¿“≈√Œ–»ﬂ 3</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x1\" ss:MergeAcross=\"1\"><Data ss:Type=\"String\"> ¿“≈√Œ–»ﬂ 4</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x1\"><Data ss:Type=\"String\">ƒÕ»</Data></Cell>");
            sw.WriteLine("    <Cell><Data ss:Type=\"String\"></Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x1\" ss:MergeAcross=\"1\"><Data ss:Type=\"String\">Õ¿À»◊Õ€≈</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x1\" ss:MergeAcross=\"1\"><Data ss:Type=\"String\">“–¿Õ—œŒÕƒ≈–</Data></Cell>");
            sw.WriteLine("  </Row>");
            sw.WriteLine("  <Row>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">##</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">Õ¿»Ã≈ÕŒ¬¿Õ»≈</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">¬–≈Ãﬂ</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">ƒ≈Õ‹</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">Õ¿À</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">“–¿Õ—œ</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">Õ¿À</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">“–¿Õ—œ</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">Õ¿À</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">“–¿Õ—œ</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">Õ¿À</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">“–¿Õ—œ</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">Õ≈ƒ≈À»</Data></Cell>");
            sw.WriteLine("    <Cell><Data ss:Type=\"String\"></Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">ÕŒ◊‹</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">ƒ≈Õ‹</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">ÕŒ◊‹</Data></Cell>");
            sw.WriteLine("    <Cell ss:StyleID=\"x2\"><Data ss:Type=\"String\">ƒ≈Õ‹</Data></Cell>");
            sw.WriteLine("  </Row>");
            int line = 2;
            List<string> payed = new List<string>();
            List<int> pOQ = new List<int>();
            List<int> pPR = new List<int>();
            foreach (AvtodorTRWeb.PayWay pay in payways)
            {
                if (String.IsNullOrEmpty(pay.ID)) continue;
                bool isFirst = true;
                if (!(String.IsNullOrEmpty(pay.ID)) && (char.IsDigit(pay.ID[pay.ID.Length - 1])))
                    if (pay.Tarrifs.Count == 4)
                        pOQ.Add(line + (pay.Tarrifs[1].Costs[1] > pay.Tarrifs[3].Costs[1] ? 2 : 4));
                    else
                        pOQ.Add(line + 2);
                foreach (AvtodorTRWeb.Tarrif tar in pay.Tarrifs)
                {
                    ++line;
                    sw.WriteLine("  <Row>");
                    string col = "col2";
                    string nor = "nor2r";
                    try
                    {
                        col = ((char.IsDigit(pay.ID[pay.ID.Length - 1])) ? "col1" : "col2") + (((line % 2) == 0) ? "b" : "r");
                        nor = ((char.IsDigit(pay.ID[pay.ID.Length - 1])) ? "nor1r" : "nor2r");
                    }
                    catch { };
                    if (isFirst)
                    {
                        sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:MergeDown=\"" + (pay.Tarrifs.Count - 1).ToString() + "\" ss:StyleID=\"" + col + "\"><Data ss:Type=\"String\">{0}</Data></Cell>", pay.ID));
                        sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:MergeDown=\"" + (pay.Tarrifs.Count - 1).ToString() + "\" ss:StyleID=\"" + nor + "\"><Data ss:Type=\"String\"><B>{0}</B></Data></Cell>", pay.Name));
                        isFirst = false;
                    };

                    sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:Index=\"3\" ss:StyleID=\"" + col + "\"><Data ss:Type=\"String\">{0}</Data></Cell>", tar.TimeRus));
                    sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"" + col + "\"><Data ss:Type=\"String\">{0}</Data></Cell>", tar.DaysRus));
                    sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"" + col + "\"><Data ss:Type=\"Number\">{0}</Data></Cell>", tar.GetCost(1, false)));
                    sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"" + col + "\"><Data ss:Type=\"Number\">{0}</Data></Cell>", tar.GetCost(1, true)));
                    sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"" + col + "\"><Data ss:Type=\"Number\">{0}</Data></Cell>", tar.GetCost(2, false)));
                    sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"" + col + "\"><Data ss:Type=\"Number\">{0}</Data></Cell>", tar.GetCost(2, true)));
                    sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"" + col + "\"><Data ss:Type=\"Number\">{0}</Data></Cell>", tar.GetCost(3, false)));
                    sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"" + col + "\"><Data ss:Type=\"Number\">{0}</Data></Cell>", tar.GetCost(3, true)));
                    sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"" + col + "\"><Data ss:Type=\"Number\">{0}</Data></Cell>", tar.GetCost(4, false)));
                    sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"" + col + "\"><Data ss:Type=\"Number\">{0}</Data></Cell>", tar.GetCost(4, true)));
                    sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"" + col + "\"><Data ss:Type=\"String\">{0}</Data></Cell>", tar.Days));
                    sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
                    if ((line % 2) == 0) sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
                    else sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\" ss:Formula=\"=R[0]C[-10]\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
                    if ((line % 2) == 1) sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
                    else sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\" ss:Formula=\"=R[0]C[-11]\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
                    if ((line % 2) == 0) sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
                    else sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\" ss:Formula=\"=R[0]C[-11]\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
                    if ((line % 2) == 1) sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
                    else sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\" ss:Formula=\"=R[0]C[-12]\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
                    sw.WriteLine("  </Row>");
                };
            };
            sw.WriteLine("  <Row>");
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine("  </Row>");

            line += 3;
            string sO = "0";
            foreach (int iv in pOQ) sO += "+R[" + (iv - line).ToString() + "]C";

            sw.WriteLine("  <Row>");
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:MergeDown=\"3\" ss:StyleID=\"col1r\"><Data ss:Type=\"String\">{0}</Data></Cell>", "ALL"));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:MergeDown=\"3\" ss:StyleID=\"col1r\"><Data ss:Type=\"String\">{0}</Data></Cell>", "¬ÒÂ Û˜‡ÒÚÍË Ú‡ÒÒ˚ M4"));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\"><Data ss:Type=\"String\">{0}</Data></Cell>", payways.MaxTarrif[0].TimeRus));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\"><Data ss:Type=\"String\">{0}</Data></Cell>", payways.MaxTarrif[0].DaysRus));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\" ss:Formula=\"=R[0]C[10]\"><Data ss:Type=\"Number\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\" ss:Formula=\"=R[0]C[11]\"><Data ss:Type=\"Number\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\" ss:Formula=\"=RC[-3]-RC[-2]\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:MergeAcross=\"4\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\" ss:Formula=\"={0}\"><Data ss:Type=\"String\"></Data></Cell>", sO));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\" ss:Formula=\"={0}\"><Data ss:Type=\"String\"></Data></Cell>", sO));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine("  </Row>");
            sw.WriteLine("  <Row>");
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:Index=\"3\" ss:StyleID=\"col1r\"><Data ss:Type=\"String\">{0}</Data></Cell>", payways.MaxTarrif[1].TimeRus));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\"><Data ss:Type=\"String\">{0}</Data></Cell>", payways.MaxTarrif[1].DaysRus));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\" ss:Formula=\"=R[0]C[11]\"><Data ss:Type=\"Number\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\" ss:Formula=\"=R[0]C[12]\"><Data ss:Type=\"Number\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\" ss:Formula=\"=RC[-3]-RC[-2]\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:MergeAcross=\"4\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\" ss:Formula=\"={0}\"><Data ss:Type=\"String\"></Data></Cell>", sO));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\" ss:Formula=\"={0}\"><Data ss:Type=\"String\"></Data></Cell>", sO));
            sw.WriteLine("  </Row>");
            sw.WriteLine("  <Row>");
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:Index=\"3\" ss:StyleID=\"col1r\"><Data ss:Type=\"String\">{0}</Data></Cell>", payways.MaxTarrif[2].TimeRus));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\"><Data ss:Type=\"String\">{0}</Data></Cell>", payways.MaxTarrif[2].DaysRus));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\" ss:Formula=\"=R[0]C[10]\"><Data ss:Type=\"Number\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\" ss:Formula=\"=R[0]C[11]\"><Data ss:Type=\"Number\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\" ss:Formula=\"=RC[-3]-RC[-2]\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:MergeAcross=\"4\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\" ss:Formula=\"={0}\"><Data ss:Type=\"String\"></Data></Cell>", sO));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\" ss:Formula=\"={0}\"><Data ss:Type=\"String\"></Data></Cell>", sO));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine("  </Row>");
            sw.WriteLine("  <Row>");
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:Index=\"3\" ss:StyleID=\"col1r\"><Data ss:Type=\"String\">{0}</Data></Cell>", payways.MaxTarrif[3].TimeRus));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\"><Data ss:Type=\"String\">{0}</Data></Cell>", payways.MaxTarrif[3].DaysRus));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\" ss:Formula=\"=R[0]C[11]\"><Data ss:Type=\"Number\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\" ss:Formula=\"=R[0]C[12]\"><Data ss:Type=\"Number\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col1r\" ss:Formula=\"=RC[-3]-RC[-2]\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:MergeAcross=\"4\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\" ss:Formula=\"={0}\"><Data ss:Type=\"String\"></Data></Cell>", sO));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\"><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:StyleID=\"col0\" ss:Formula=\"={0}\"><Data ss:Type=\"String\"></Data></Cell>", sO));
            sw.WriteLine("  </Row>");

            sw.WriteLine("  <Row>");
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine("  </Row>");

            sw.WriteLine("  <Row>");
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:MergeAcross=\"11\"><Data ss:Type=\"String\">Grab Url: {0}</Data></Cell>", payways.url));
            sw.WriteLine("  </Row>");
            sw.WriteLine("  <Row>");
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:MergeAcross=\"11\"><Data ss:Type=\"String\">ƒ‡ÌÌ˚Â ÔÓÎÛ˜ÂÌ˚: {0}</Data></Cell>", payways.grabbed.ToString("dd.MM.yyyy HH:mm")));
            sw.WriteLine("  </Row>");
            sw.WriteLine("  <Row>");
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell><Data ss:Type=\"String\">{0}</Data></Cell>", ""));
            sw.WriteLine(String.Format(System.Globalization.CultureInfo.InvariantCulture, "    <Cell ss:MergeAcross=\"11\"><Data ss:Type=\"String\">Created by: {0} (version {1})</Data></Cell>", "Avtodor-Tr M4 Grabber by milokz@gmail.com", avtodor_tr_m4.Program.version));
            sw.WriteLine("  </Row>");

            sw.WriteLine("</Table>");
            sw.WriteLine("<WorksheetOptions xmlns=\"urn:schemas-microsoft-com:office:excel\">");
            sw.WriteLine("  <FreezePanes/>");
            sw.WriteLine("  <FrozenNoSplit/>");
            sw.WriteLine("  <SplitHorizontal>2</SplitHorizontal>");
            sw.WriteLine("  <TopRowBottomPane>2</TopRowBottomPane>");
            sw.WriteLine("  <SplitVertical>1</SplitVertical>");
            sw.WriteLine("  <LeftColumnRightPane>1</LeftColumnRightPane>");
            sw.WriteLine("  <ActivePane>0</ActivePane>");
            sw.WriteLine("</WorksheetOptions>");
            sw.WriteLine("</Worksheet>");
        }

        public static void Export2ExcelXML(PayWays payways, string fname)
        {
            FileStream fs = new FileStream(fname, FileMode.Create, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            sw.WriteLine("<?xml version=\"1.0\"?>");
            sw.WriteLine("<?mso-application progid=\"Excel.Sheet\"?>");
            sw.WriteLine("<Workbook xmlns=\"urn:schemas-microsoft-com:office:spreadsheet\" xmlns:ss=\"urn:schemas-microsoft-com:office:spreadsheet\">");
            sw.WriteLine("<DocumentProperties>\r\n  <Created>Avtodor-Tr M4 Grabber by milokz@gmail.com</Created>\r\n</DocumentProperties>");
            sw.WriteLine("<Styles>");
            Export2ExcelXML(payways, sw);            
            sw.WriteLine("</Workbook>");
            sw.Close();
            fs.Close();
        }
    }

    public class PayWay
    {
        [XmlAttribute("id")]
        public string ID = String.Empty;
        [XmlAttribute("name")]
        public string Name = String.Empty;
        [XmlElement("tarrifs")]
        public List<Tarrif> Tarrifs = new List<Tarrif>();

        public void AddCost(string day, string time, int cat, int casht, float cost)
        {
            string sday = ParseDays(day);
            if (time == null) time = "";
            bool f24 = String.IsNullOrEmpty(time);
            bool f07 = sday.Length == 13;
            bool ex = false;
            for (int i = 0; i < Tarrifs.Count; i++)
            {
                if ((f07 && sday.Contains(Tarrifs[i].Days)) || (Tarrifs[i].Days == sday))
                {
                    if (f24 || (Tarrifs[i].Time == time))
                    {
                        ex = true;
                        Tarrifs[i].Costs[2 * cat - 2  + casht] = cost;
                    };
                };
            };
            if (!ex)
            {
                if (!f24)
                {
                    Tarrif t = new Tarrif();
                    t.Time = time;
                    t.Days = sday;
                    t.Costs = new float[8];
                    t.Costs[2 * cat - 2 + casht] = cost;
                    Tarrifs.Add(t);
                }
                else
                {
                    Tarrif t = new Tarrif();
                    t.Time = "DAY";
                    t.Days = sday;
                    t.Costs = new float[8];
                    t.Costs[2 * cat - 2 + casht] = cost;
                    Tarrifs.Add(t);
                    t = new Tarrif();
                    t.Time = "NIGHT";
                    t.Days = sday;
                    t.Costs = new float[8];
                    t.Costs[2 * cat - 2 + casht] = cost;
                    Tarrifs.Add(t);
                };
            };
        }

        public void SetName(string name, AvtodorTRWeb.Segments segments)
        {
            this.Name = name;
            if (String.IsNullOrEmpty(this.Name)) return;
            if (segments == null) return;
            if (segments.segments == null) return;
            if (segments.segments.Count == 0) return;

            foreach (AvtodorTRWeb.Segment sg in segments.segments)
            {
                int same = 0;
                foreach (string has in sg.Has)
                {
                    if (this.Name == has)
                        same++;
                    else if (this.Name.Contains(has))
                        same++;
                };
                if (same == sg.Has.Count)
                    this.ID = sg.ID;
            };
        }

        public void CheckID(string text, AvtodorTRWeb.Segments segments)
        {
            if (String.IsNullOrEmpty(text)) return;
            if (segments == null) return;
            if (segments.segments == null) return;
            if (segments.segments.Count == 0) return;

            foreach (AvtodorTRWeb.Segment sg in segments.segments)
            {
                int same = 0;
                foreach (string has in sg.Has)
                {
                    if (text == has)
                        same++;
                    else if (text.Contains(has))
                        same++;
                };
                if (same == sg.Has.Count)
                    this.ID = sg.ID;
            };
        }

        public override string ToString()
        {
            string tfs = "";
            foreach (Tarrif t in Tarrifs) tfs += t.Time + " " + t.Days + "; ";
            return String.Format("{0} - {1}: {2}", ID, Tarrifs.Count, tfs);
        }

        public static string ParseDays(string day)
        {
            if (String.IsNullOrEmpty(day)) return "1,2,3,4,5,6,7";

            List<int> dd = new List<int>();
            List<string> dof = new List<string>();
            dof.AddRange(Tarrif.DaysOfWeek);
            int prev = -1;
            char prel = '\0';
            
            string d = day.Replace(" ", "");
            while(d.Length > 0)
            {
                if ((d[0] == '-') || (d[0] == ':') || (d[0] == '.') || (d[0] == ','))
                {
                    prel = d[0];
                    d = d.Remove(0, 1);
                    continue;
                };
                string f = d.Substring(0, 2);
                d = d.Remove(0, 2);
                int curr = dof.IndexOf(f);
                if ((prel == '-') || (prel == ':') || (prel == '.'))
                {
                    int next = prev + 1;
                    if (next == 7) next = 0;
                    while (next != curr)
                    {
                        dd.Add(next);
                        ++next;
                        if (next == 7) next = 0;
                    };
                };
                dd.Add(prev = curr);
                prel = '\0';
            };
            dd.Sort();

            string res = "";  
            foreach(int i in dd) res += (res.Length > 0 ? "," : "") + Tarrif.DaysOfWeekExcel[i];
            return res;
        }
    }

    public class PayWaySorter : IComparer<PayWay>
    {
        public int Compare(PayWay a, PayWay b)
        {
            string va = a.ID;
            string vb = b.ID;
            Regex ex = new Regex(@"[\d]+");
            if (ex.Match(va).Success)
                va = int.Parse("0" + ex.Match(va).Value).ToString("00000") + va.Replace(ex.Match(va).Value, "");
            if (ex.Match(vb).Success)
                vb = int.Parse("0" + ex.Match(vb).Value).ToString("00000") + vb.Replace(ex.Match(vb).Value, "");
            return va.CompareTo(vb);
        }
    }

    public class Tarrif
    {
        public static string[] DaysOfWeek = new string[] { "œÕ", "¬“", "—–", "◊“", "œ“", "—¡", "¬—" };
        public static string[] DaysOfWeekExcel = new string[] { "2", "3", "4", "5", "6", "7", "1" };

        [XmlAttribute("time")]
        public string Time = String.Empty;
        [XmlAttribute("cost")]
        public float[] Costs = new float[16]; // CAT_1 - CAT_2 - CAT_3 - CAT_4
        [XmlAttribute("days")]
        public string Days = "1,2,3,4,5,6,7";

        [XmlIgnore]
        public string TimeRus
        {
            get
            {
                if (Time == "NIGHT") return "ÕŒ◊‹";
                if (Time == "DAY") return "ƒ≈Õ‹";
                return "???";
            }
        }

        [XmlIgnore]
        public string DaysRus
        {
            get
            {
                if (String.IsNullOrEmpty(this.Days)) return "???";
                List<string> dd = new List<string>(new string[] { "¬—", "œÕ", "¬“", "—–", "◊“", "œ“", "—¡" });
                string[] d = this.Days.Split(new char[] { ',' }, 7);
                int pd = Convert.ToInt32(d[0]) - 1;
                string res = dd[pd];
                int lc = 0;
                int cd = pd;
                for (int i = 1; i < d.Length; i++)
                {
                    cd = Convert.ToInt32(d[i]) - 1;
                    if (((cd - pd) == 1) || ((cd == 0) && (pd == 6)))
                        lc++;
                    else
                    {
                        if (lc == 0) res += "," + dd[cd];
                        if (lc == 1) res += "," + dd[pd] + "," + dd[cd];
                        if (lc > 1) res += "-" + dd[pd] + "," + dd[cd];
                        lc = 0;
                    };
                    pd = cd;
                };
                if (lc == 1) res += "," + dd[cd];
                if (lc > 1) res += "-" + dd[cd];
                return res;
            }
        }

        public Tarrif Clone()
        {
            Tarrif res = new Tarrif();
            res.Time = this.Time;
            res.Costs = (new List<float>(this.Costs)).ToArray();
            return res;
        }

        public void AddCost(string txt, ref int catub)
        {
            if (String.IsNullOrEmpty(txt)) return;
            Regex rx = new Regex(@"(?<cost>\d+[.,\d]+)");
            Match mx = rx.Match(txt);
            if (mx.Success)
            {
                float val = 0;
                if (float.TryParse(mx.Groups["cost"].Value.Replace(",", "."), System.Globalization.NumberStyles.Any, System.Globalization.CultureInfo.InvariantCulture, out val))
                    this.Costs[catub] = val;
                catub++;
            };
        }

        public float GetCost(byte cat, bool transponder)
        {
            if (transponder)
                return Math.Min(this.Costs[cat * 2 - 2], this.Costs[cat * 2 - 1]);
            else
                return Math.Max(this.Costs[cat * 2 - 2], this.Costs[cat * 2 - 1]);
        }        

        public bool HasDay(byte day)
        {
            return Days.IndexOf(day.ToString()) >= 0;
        }

        public bool HasDay(string day)
        {
            List<string> di = new List<string>(new string[] { "¬—", "œÕ", "¬“", "—–", "◊“", "œ“", "—¡" });
            int dir = Days.IndexOf(di.IndexOf(day.ToUpper()).ToString());
            List<string> ei = new List<string>(new string[] { "SU", "MO", "TU", "WE", "TH", "FR", "SA" });
            int eir = Days.IndexOf(di.IndexOf(day.ToUpper()).ToString());
            return Math.Max(dir, eir) >= 0;
        }

        public bool HasTime(DateTime time)
        {
            if (String.IsNullOrEmpty(this.Time)) return true;
            string dd = this.Time.Replace(" ", "").ToUpper().Trim();
            Regex rx = new Regex(@"(?<from>\w{1,2}:\w{1,2})-(?<to>\w{1,2}:\w{1,2})");
            Match mc = rx.Match(dd);
            if (!mc.Success) return true;
            string tf = mc.Groups["from"].Value;
            string tt = mc.Groups["to"].Value;
            DateTime dtf = DateTime.ParseExact(tf, "HH:mm", System.Globalization.CultureInfo.InvariantCulture);
            DateTime dtt = DateTime.ParseExact(tt, "HH:mm", System.Globalization.CultureInfo.InvariantCulture);
            time = new DateTime(dtf.Year, dtf.Month, dtf.Day, time.Hour, time.Minute, time.Second);
            List<DateTime[]> intervals = new List<DateTime[]>();
            if (dtt > dtf)
                intervals.Add(new DateTime[] { dtf, dtt });
            else
            {
                intervals.Add(new DateTime[] { dtf, dtf.Date.AddDays(1) });
                intervals.Add(new DateTime[] { dtt.Date, dtt });
            };
            foreach (DateTime[] dti in intervals)
                if ((dti[0] <= time) && (time <= dti[1]))
                    return true;
            return false;
        }

        public bool HasTime(double time)
        {
            return HasTime(DateTime.FromOADate(time));
        }
    }

    public class TarifSorter : IComparer<Tarrif>
    {
        public int Compare(Tarrif a, Tarrif b)
        {
            return -1 * (a.Days + " - " + a.Time).CompareTo(b.Days + " - " + b.Time);
            //return -1 * (a.Time + " - " + a.Days).CompareTo(b.Time + " - " + b.Days);
        }
    }

    public class TabledData
    {
        /// <summary>
        ///     Cell Address RC
        /// </summary>
        public class RCItem
        {
            /// <summary>
            ///     Row
            /// </summary>
            public int row = 0;
            /// <summary>
            ///     Column
            /// </summary>
            public int col = 0;
            /// <summary>
            ///     Address (A1,B2...)
            /// </summary>
            public string Address { get { return String.Format("{0}:{1}", ColumnIndex(col), row); } }

            /// <summary>
            ///     Basic constructor
            /// </summary>
            public RCItem() { }

            /// <summary>
            ///     Constructor
            /// </summary>
            /// <param name="row">Row</param>
            /// <param name="col">Column</param>
            public RCItem(int row, int col)
            {
                this.row = row;
                this.col = col;
            }

            /// <summary>
            ///     Constructor
            /// </summary>
            /// <param name="row">Row</param>
            /// <param name="col">Column (A,B...)</param>
            public RCItem(int row, string col)
            {
                this.row = row;
                this.col = ColumnIndex(col);
            }

            /// <summary>
            ///     Constructor
            /// </summary>
            /// <param name="reference">Address (A1,B2...)</param>
            public RCItem(/* A1, B2 ... */ string reference)
            {
                reference = reference.Trim();
                if (Char.IsDigit(reference[0]))
                {
                    string[] rc = reference.Split(new char[] { ':' });
                    row = int.Parse(rc[0]);
                    col = int.Parse(rc[1]);
                }
                else
                {
                    if (reference.IndexOf(":") > 0)
                    {
                        string[] cr = reference.Split(new char[] { ':' });
                        col = ColumnIndex(cr[0]);
                        row = int.Parse(cr[1]);
                    }
                    else
                    {
                        int sci = 0;
                        while (sci < reference.Length) if (char.IsDigit(reference[sci++])) break;
                        col = ColumnIndex(reference.Substring(0, sci));
                        row = int.Parse(reference.Substring(sci - 1));
                    };
                };
            }

            /// <summary>
            ///     Get Column Text Representation from Index
            /// </summary>
            /// <param name="reference">columnt text address</param>
            /// <returns>column index address</returns>
            public static int ColumnIndex(string reference)
            {
                int ci = 0;
                reference = reference.ToUpper();
                for (int ix = 0; ix < reference.Length && reference[ix] >= 'A'; ix++)
                    ci = (ci * 26) + ((int)reference[ix] - 64);
                return ci;
            }

            /// <summary>
            ///     Get Column Index Representation from Text
            /// </summary>
            /// <param name="reference">column index address</param>
            /// <returns>columnt text address</returns>
            public static string ColumnIndex(int reference)
            {
                string result = String.Empty;
                while (reference > 0)
                {
                    int mod = (reference - 1) % 26;
                    result = Convert.ToChar(65 + mod).ToString() + result;
                    reference = (int)((reference - mod) / 26);
                };
                return result;
            }
        }

        private uint MinRow = 0;
        private uint MaxRow = 0;
        private uint MinCol = 0;
        private uint MaxCol = 0;
        private string fileName;

        private Dictionary<ulong, TextHTML> data = new Dictionary<ulong, TextHTML>();
        private List<ulong> filled = new List<ulong>();
        private List<ulong> changed = new List<ulong>();

        public int MinX { get { return (int)MinRow; } }
        public int MaxX { get { return (int)MaxRow; } }
        public int MinY { get { return (int)MinCol; } }
        public int MaxY { get { return (int)MaxCol; } }
        public int Width { get { return (int)(MaxCol == 0 ? 0 : MaxCol - MinCol + 1); } }
        public int Height { get { return (int)(MaxRow == 0 ? 0 : MaxRow - MinRow + 1); } }
        public string FileName { get { return fileName; } set { fileName = value; } }

        /// <summary>
        ///     Clear Cells
        /// </summary>
        public void Clear()
        {
            data.Clear();
            filled.Clear();
            changed.Clear();
            MinRow = 0;
            MaxRow = 0;
            MinCol = 0;
            MaxCol = 0;
        }

        /// <summary>
        ///     Reset Last Changes in Cells
        /// </summary>
        public void ResetChanged()
        {
            changed.Clear();
        }

        private void Change(ulong item)
        {
            if (!filled.Contains(item)) filled.Add(item);
            if (!changed.Contains(item)) changed.Add(item);
        }

        /// <summary>
        ///     Get Filled Cell Address by index (zero-indexed)
        /// </summary>
        /// <param name="index">index (zero-indexed)</param>
        /// <returns>RC value</returns>
        public RCItem GetFilled(int index)
        {
            if (index < 0) return null;
            if (index >= this.filled.Count) return null;
            ulong item = this.filled[index];
            int row = (int)((item >> 32) & 0xFFFFFFFF);
            int col = (int)(item & 0xFFFFFFFF);
            return new RCItem(row, col);
        }

        /// <summary>
        ///     Get Last Changed Cell Address by index (zero-indexed)
        /// </summary>
        /// <param name="index">index (zero-indexed)</param>
        /// <returns>RC value</returns>
        public RCItem GetChanged(int index)
        {
            if (index < 0) return null;
            if (index >= this.changed.Count) return null;
            ulong item = this.changed[index];
            int row = (int)((item >> 32) & 0xFFFFFFFF);
            int col = (int)(item & 0xFFFFFFFF);
            return new RCItem(row, col);
        }

        /// <summary>
        ///     Get Filled Cells Count
        /// </summary>
        public int FilledCount
        {
            get
            {
                return data.Count;
            }
        }

        /// <summary>
        ///     Get Last Changed Cells Count
        /// </summary>
        public int ChangedCount
        {
            get
            {
                return changed.Count;
            }
        }

        /// <summary>
        ///     Get Filled Cells Bounds (MinRow, MaxRow, MinCol, MaxCol)
        /// </summary>
        /// <returns></returns>
        public int[] GetFilledBounds()
        {
            return new int[] { (int)MinRow, (int)MaxRow, (int)MinCol, (int)MaxCol };
        }

        /// <summary>
        ///     Get Changed Cells Bounds (MinRow, MaxRow, MinCol, MaxCol)
        /// </summary>
        /// <returns></returns>
        public int[] GetChangedBounds()
        {
            if (this.changed.Count == 0) return new int[] { 0, 0, 0, 0 };
            int[] res = new int[] { int.MaxValue, int.MinValue, int.MaxValue, int.MinValue };
            foreach (long item in this.changed)
            {
                int row = (int)((item >> 32) & 0xFFFFFFFF);
                int col = (int)(item & 0xFFFFFFFF);
                if (row < res[0]) res[0] = row;
                if (row > res[1]) res[1] = row;
                if (col < res[2]) res[2] = col;
                if (col > res[3]) res[3] = col;
            };
            return res;
        }

        /// <summary>
        ///     Set Cell Value & Formula
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="value">value</param>
        /// <param name="formula">formula</param>
        public void Set(uint row, uint col, string text, string html)
        {
            if (String.IsNullOrEmpty(text)) text = "";
            if (String.IsNullOrEmpty(html)) html = "";
            SetMinMax(row, col);
            ulong index = (((ulong)row) << 32) + (ulong)col;
            Change(index);
            if (data.ContainsKey(index))
                data[index] = new TextHTML(text, html);
            else
                data.Add(index, new TextHTML(text, html));
        }

        /// <summary>
        ///     Set Cell Formula
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="formula">formula</param>
        public void Set(int row, int col, string html) { Set((uint)row, (uint)col, html); }

        /// <summary>
        ///     Set Cell Formula
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="formula">formula</param>
        public void Set(uint row, uint col, string html)
        {
            if (String.IsNullOrEmpty(html)) html = "";
            SetMinMax(row, col);
            ulong index = (((ulong)row) << 32) + (ulong)col;
            Change(index);
            if (data.ContainsKey(index))
                data[index] = new TextHTML("", html);
            else
                data.Add(index, new TextHTML("", html));
        }

        /// <summary>
        ///     Set Cell Value & Formula
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="ValueFormula">Value & Formula</param>
        public void Set(int row, int col, TextHTML html) { Set((uint)row, (uint)col, html); }

        /// <summary>
        ///     Set Cell Value & Formula
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <param name="ValueFormula">Value & Formula</param>
        public void Set(uint row, uint col, TextHTML vf)
        {
            if (String.IsNullOrEmpty(vf.text)) vf.text = "";
            if (String.IsNullOrEmpty(vf.html)) vf.html = "";
            SetMinMax(row, col);
            ulong index = (((ulong)row) << 32) + (ulong)col;
            Change(index);
            if (data.ContainsKey(index))
                data[index] = vf;
            else
                data.Add(index, vf);
        }

        /// <summary>
        ///     Get Cell Value & Formula
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <returns>Value & Formula</returns>
        public TextHTML Get(int row, int col) { return Get((uint)row, (uint)col); }

        /// <summary>
        ///     Get Cell Value & Formula
        /// </summary>
        /// <param name="row">row</param>
        /// <param name="col">column</param>
        /// <returns>Value & Formula</returns>
        public TextHTML Get(uint row, uint col)
        {
            ulong index = (((ulong)row) << 32) + (ulong)col;
            if (data.ContainsKey(index))
                return data[index];
            else
                return new TextHTML("", "");
        }

        private void SetMinMax(uint row, uint col)
        {
            if (MinRow == 0) MinRow = row;
            if (MaxRow == 0) MinRow = row;
            if (MinCol == 0) MinCol = col;
            if (MaxCol == 0) MinCol = col;

            if (row < MinRow) MinRow = row;
            if (row > MaxRow) MaxRow = row;
            if (col < MinCol) MinCol = col;
            if (col > MaxCol) MaxCol = col;
        }
    }

    public class TextHTML
    {
        public string text;
        public string html;
        public TextHTML(string text, string html) { this.text = text; this.html = html; }
    }

    [Serializable]
    public class XMLSaved<T>
    {
        /// <summary>
        ///     —Óı‡ÌÂÌËÂ ÒÚÛÍÚÛ˚ ‚ Ù‡ÈÎ
        /// </summary>
        /// <param name="file">œÓÎÌ˚È ÔÛÚ¸ Í Ù‡ÈÎÛ</param>
        /// <param name="obj">—ÚÛÍÚÛ‡</param>
        public static void Save(T obj)
        {
            string fname = System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase.ToString();
            fname = fname.Replace("file:///", "");
            fname = fname.Replace("/", @"\");
            fname = fname.Substring(0, fname.LastIndexOf(@"\") + 1);
            fname += @"segments.xml";

            System.Xml.Serialization.XmlSerializer xs = new System.Xml.Serialization.XmlSerializer(typeof(T));
            System.IO.StreamWriter writer = System.IO.File.CreateText(fname);
            xs.Serialize(writer, obj);
            writer.Flush();
            writer.Close();
        }

        /// <summary>
        ///     œÓ‰ÍÎ˛˜ÂÌËÂ ÒÚÛÍÚÛ˚ ËÁ Ù‡ÈÎ‡
        /// </summary>
        /// <param name="file">œÓÎÌ˚È ÔÛÚ¸ Í Ù‡ÈÎÛ</param>
        /// <returns>—ÚÛÍÚÛ‡</returns>
        public static T Load()
        {
            string fname = System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase.ToString();
            fname = fname.Replace("file:///", "");
            fname = fname.Replace("/", @"\");
            fname = fname.Substring(0, fname.LastIndexOf(@"\") + 1);
            fname += @"segments.xml";

            // if couldn't create file in temp - add credintals
            // http://support.microsoft.com/kb/908158/ru
            System.Xml.Serialization.XmlSerializer xs = new System.Xml.Serialization.XmlSerializer(typeof(T));
            System.IO.StreamReader reader = System.IO.File.OpenText(fname);
            T c = (T)xs.Deserialize(reader);
            reader.Close();
            return c;
        }
    }
}
