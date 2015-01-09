using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.IO;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.Data;
using System.Windows.Forms;
using NPOI.HSSF.UserModel;//excel 97
using NPOI.SS.UserModel;

namespace taobaoINFO
{
    class getInfo
    {
        static public string getKEY(string html)
        {
            Regex reg = new Regex("&keys=(?<key>.*)\"", RegexOptions.IgnoreCase);
            Match m = reg.Match(html);
            return m.Groups["key"].Value;
        }
        static public string getSBN(string html)
        {
            Regex reg = new Regex("&sbn=(?<sbn>.*)&is", RegexOptions.IgnoreCase);
            Match m = reg.Match(html);
            return m.Groups["sbn"].Value;
        }
        static public string getSID(string html)
        {
            Regex reg = new Regex("sellerId:\"(?<sid>.*)\",", RegexOptions.IgnoreCase);
            Match m = reg.Match(html);
            return m.Groups["sid"].Value;
        }
        static public string getID(string url)
        {
            Regex reg = new Regex("id=(?<id>\\d*)", RegexOptions.IgnoreCase);
            Match m = reg.Match(url);
            return m.Groups["id"].Value;
        }
        static public string getHtml(string url)
        {
            WebClient wc = new WebClient();
            wc.Credentials = CredentialCache.DefaultCredentials;
            Stream resStream = wc.OpenRead(url);
            StreamReader sr = new StreamReader(resStream, Encoding.Default);
            string html = sr.ReadToEnd();
            resStream.Close();
            return html;
        }
        static public bool maijiagongju(string url, ref List<string> result, ref DataRow dr)
        {
            string postDataStr = "{\"_MING_ROOT_\":{\"title\":\"";
            postDataStr += url;
            postDataStr += "\"},\"_MING_CLASS_\":\"_MING_CLASS_\"}";
            string info = httprequest(URL.maijiagongju, "POST", "utf-8", postDataStr);
            if (info == "{}")
            {
                result.Add("链接失效，无法查询！:" + url);
                dr["title"] = "链接失效，无法查询！:" + url;
                return false;
            }
            info = "[" + info + "]";
            JArray ja = (JArray)JsonConvert.DeserializeObject(info);
            result.Add(ja[0]["desc"]["title"].ToString());//标题
            result.Add(ja[0]["desc"]["cname"].ToString());//所属类目
            result.Add(ja[0]["desc"]["desc"].ToString()); //标题分析结果
            result.Add(ja[0]["desc"]["score"].ToString());//标题关键词得分 
            result.Add(ja[0]["desc"]["catPrior"].ToString());//类目关键词数量
            result.Add(ja[0]["desc"]["wordNum"].ToString());//标题关键词数量
            result.Add(ja[0]["desc"]["highPVNum"].ToString());//核心关键词数量
            result.Add(ja[0]["desc"]["highCharNum"].ToString());//高质量关键词数量
            dr["title"] = ja[0]["desc"]["title"].ToString();
            dr["category"] = ja[0]["desc"]["cname"].ToString();
            dr["title_res"] = ja[0]["desc"]["desc"].ToString();
            dr["title_key_score"] = ja[0]["desc"]["score"].ToString();
            dr["cate_key_num"] = ja[0]["desc"]["catPrior"].ToString();
            dr["title_key_num"] = ja[0]["desc"]["wordNum"].ToString();
            dr["core_key_num"] = ja[0]["desc"]["highPVNum"].ToString();
            dr["high_key"] = ja[0]["desc"]["highCharNum"].ToString();
            return true;
        }
        static public void wangwangid(string html, ref List<string> result, ref DataRow dr)
        {
            Regex reg = new Regex("sellerNick:'(?<sellerNick>.\\w*)',", RegexOptions.IgnoreCase);
            Match m = reg.Match(html);
            result.Add(m.Groups["sellerNick"].Value);
            dr["id"] = m.Groups["sellerNick"].Value;
        }
        static public void getPic(string html, ref List<string> result, ref DataRow dr)
        {
            Regex reg = new Regex("pic:\\s*\"(?<pic>.*)\"", RegexOptions.IgnoreCase);
            Match m = reg.Match(html);
            result.Add(m.Groups["pic"].Value);
            dr["pic"] = m.Groups["pic"].Value;
        }
        static public void getPrice(string html, ref List<string> result, ref DataRow dr)
        {
            Regex reg = new Regex("price:\\s*(?<price>.*),", RegexOptions.IgnoreCase);
            Match m = reg.Match(html);
            result.Add(m.Groups["price"].Value);
            dr["price"] = m.Groups["price"].Value;
        }
        static public void zhekoujia(ref List<string> result, ref DataRow dr)
        {
            result.Add("0");
            dr["price_now"] = "0";
        }
        static public void paixiaMAX(ref List<string> result, ref DataRow dr)
        {
            result.Add("0");
            dr["max_price"] = "0";
        }
        static public void paixiaMIN(ref List<string> result, ref DataRow dr)
        {
            result.Add("0");
            dr["min_price"] = "0";
        }
        static public void location(string id, ref List<string> result, ref DataRow dr)
        {
            string url = URL.location + "&itemID=" + id;
            string res = httprequest(url, "GET", "gb2312");
            Regex reg = new Regex("location:'(?<location>.*)',", RegexOptions.IgnoreCase);
            Match m = reg.Match(res);
            result.Add(m.Groups["location"].Value);
            dr["location"] = m.Groups["location"].Value;
            if (res.IndexOf("免运费") == -1)
            {
                result.Add("不包邮");
                dr["freight"] = "不包邮";
            }
            else
            {
                result.Add("包邮");
                dr["freight"] = "包邮";
            }
        }
        static public void pingjia(string id, string sid, ref List<string> result, ref DataRow dr)
        {
            string url = URL.pingjia + "auctionNumId=" + id + "&userNumId=" + sid;
            string res = httprequest(url, "GET", "gb2312");
            Regex reg = new Regex("\"total\":(?<total>\\d*)", RegexOptions.IgnoreCase);
            Match m = reg.Match(res);
            result.Add(m.Groups["total"].Value);
            dr["review"] = m.Groups["total"].Value;
            reg = new Regex("\"good\":(?<good>\\d*)", RegexOptions.IgnoreCase);
            m = reg.Match(res);
            result.Add(m.Groups["good"].Value);
            dr["good_review"] = m.Groups["good"].Value;
            reg = new Regex("\"normal\":(?<normal>\\d*)", RegexOptions.IgnoreCase);
            m = reg.Match(res);
            result.Add(m.Groups["normal"].Value);
            dr["normal_review"] = m.Groups["normal"].Value;
            reg = new Regex("\"bad\":(?<bad>\\d*)", RegexOptions.IgnoreCase);
            m = reg.Match(res);
            result.Add(m.Groups["bad"].Value);
            dr["bad_review"] = m.Groups["bad"].Value;
            reg = new Regex("\"additional\":(?<add>\\d*)", RegexOptions.IgnoreCase);
            m = reg.Match(res);
            result.Add(m.Groups["add"].Value);
            dr["add_review"] = m.Groups["add"].Value;
            reg = new Regex("\"pic\":(?<pic>\\d*)", RegexOptions.IgnoreCase);
            m = reg.Match(res);
            result.Add(m.Groups["pic"].Value);
            dr["pic_review"] = m.Groups["pic"].Value;
        }
        static public string gettime(string timestamp)
        {
            DateTime time = DateTime.MinValue;
            DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new DateTime(1970, 1, 1));
            time = startTime.AddMilliseconds(long.Parse(timestamp));
            return time.ToString();
        }
        static public long gettimestamp(DateTime time)
        {
            System.DateTime startTime = TimeZone.CurrentTimeZone.ToLocalTime(new System.DateTime(1970, 1, 1));
            TimeSpan span = (TimeSpan)(time - startTime);
            return (long)span.TotalMilliseconds;
        }
        static public void jiaoyijinfo(string id, string sid, string sbn, string keys, string html, ref List<string> result, ref DataRow dr)
        {
            Regex reg = new Regex("dbst:(?<dbst>\\d*),", RegexOptions.IgnoreCase);
            Match m = reg.Match(html);
            string create_time = gettime(m.Groups["dbst"].Value);
            result.Add(create_time);
            dr["create_time"] = create_time;

            long dbst = long.Parse(m.Groups["dbst"].Value);

            reg = new Regex("ends=(?<ends>\\d*)&", RegexOptions.IgnoreCase);
            m = reg.Match(html);
            string end_time = gettime(m.Groups["ends"].Value);
            result.Add(end_time);
            dr["end_time"] = end_time;

            string url = URL.chengjiaoxinxi + "&id=" + id + "&sid=" + sid + "&sbn=" + sbn;
            string res = httprequest(url, "GET", "gb2312");
            reg = new Regex("confirmGoods: (?<success>.*),", RegexOptions.IgnoreCase);
            m = reg.Match(res);
            result.Add(m.Groups["success"].Value);
            dr["deal_success"] = m.Groups["success"].Value;

            reg = new Regex("quanity: (?<res>.*),", RegexOptions.IgnoreCase);
            m = reg.Match(res);
            result.Add(m.Groups["res"].Value);
            dr["success"] = m.Groups["res"].Value;

            int chengjiao = int.Parse(m.Groups["res"].Value);
            url = URL.shoucangliang + "&keys=" + keys;
            res = httprequest(url, "GET", "gb2312");
            reg = new Regex("\"ICVT\\w*\":(?<view>\\d*)", RegexOptions.IgnoreCase);
            m = reg.Match(res);
            result.Add(m.Groups["view"].Value);
            dr["view"] = m.Groups["view"].Value;

            int view = int.Parse(m.Groups["view"].Value);
            reg = new Regex("\"ICCP\\w*\":(?<collection>\\d*)", RegexOptions.IgnoreCase);
            m = reg.Match(res);
            result.Add(m.Groups["collection"].Value);
            dr["collection"] = m.Groups["collection"].Value;
            int collection = int.Parse(m.Groups["collection"].Value);
            long days = (gettimestamp(DateTime.Now) - dbst) / (24 * 60 * 60 * 1000);
            if (days == 0)
            {
                days = 1;
            }
            result.Add(days.ToString());
            result.Add((chengjiao / 30).ToString());
            result.Add((view / days).ToString());
            result.Add(((double)chengjiao / (double)view * 100).ToString("0.00"));
            result.Add(((double)collection / (double)view * 100).ToString("0.00"));
            dr["create_days"] = days.ToString();
            dr["deal_oneday"] = (chengjiao / 30).ToString();
            dr["view_oneday"] = (view / days).ToString();
            dr["con_rate"] = ((double)chengjiao / (double)view * 100).ToString("0.00");
            dr["coll_rate"] = ((double)collection / (double)view * 100).ToString("0.00");
        }
        static public void defen(string html, ref List<string> result, ref DataRow dr)
        {
            Regex reg = new Regex("低得分\\)\">\\s*(?<score>\\d\\.\\d)", RegexOptions.IgnoreCase);
            Match m = reg.Match(html);
            int no = 0;
            foreach (Match match in reg.Matches(html))
            {
                result.Add(match.Groups["score"].Value);
                if (no == 0)
                {
                    dr["describe"] = match.Groups["score"].Value;
                    no++;
                }
                else if (no == 1)
                {
                    dr["logistics"] = match.Groups["score"].Value;
                    no++;
                }
                else
                {
                    dr["service"] = match.Groups["score"].Value;
                    no++;
                }
            }
        }
        static public void shuxing(string html, ref List<string> result, ref DataRow dr)
        {
            Regex reg = new Regex("<ul class=\"attributes-list\">\\s*(?<shuxing>.*)\\s*</ul>", RegexOptions.IgnoreCase);
            Match m = reg.Match(html);
            string shuxing = m.Groups["shuxing"].Value;
            reg = new Regex("<li\\s*title=\"\\s*\\w*\">(?<shuxing>\\w*:\\s*\\w*)", RegexOptions.IgnoreCase);
            string res = "";
            foreach (Match match in reg.Matches(shuxing))
            {
                res += (match.Groups["shuxing"].Value + " / ");
            }
            result.Add(res);
            dr["attribute"] = res;
        }
        static private string httprequest(string url, string method, string encode, string data = "")
        {
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
            if (method == "POST")
            {
                request.Method = "POST";
                request.ContentType = "application/x-www-form-urlencoded";
                request.ContentLength = Encoding.UTF8.GetByteCount(data);
                request.Proxy = null;
                Stream myRequestStream = request.GetRequestStream();
                StreamWriter myStreamWriter = new StreamWriter(myRequestStream, Encoding.GetEncoding("gb2312"));
                myStreamWriter.Write(data);
                myStreamWriter.Close();
            }

            HttpWebResponse response = (HttpWebResponse)request.GetResponse();
            Stream myResponseStream = response.GetResponseStream();
            StreamReader myStreamReader = new StreamReader(myResponseStream, Encoding.GetEncoding(encode));
            string info = myStreamReader.ReadToEnd();
            myStreamReader.Close();
            myResponseStream.Close();
            return info;
        }
        static public void Updatelist(DataTable dt, ref Form1.ListViewNF listview)
        {

            ListViewItem item;
            int index = 1;
            foreach (DataRow row in dt.Rows)
            {
                item = new ListViewItem();
                item.Text = index++.ToString();
                int columns = dt.Columns.Count;
                for (int i = 0; i < columns; i++)
                {
                    string str = row[i].ToString();
                    item.SubItems.Add(str);
                }
                listview.Items.Add(item);
            }
        }
        static public string getPath(int mode, string filter, string filename = "")
        {
            if (mode == 0)
            {
                OpenFileDialog openFileDialog = new OpenFileDialog();
                openFileDialog.InitialDirectory = "c:\\";//注意这里写路径时要用c:\\而不是c:\
                openFileDialog.Filter = filter;
                openFileDialog.RestoreDirectory = true;
                openFileDialog.FilterIndex = 1;
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return openFileDialog.FileName;
                }
                else
                    return null;
            }
            else
            {
                SaveFileDialog saveFileDialog = new SaveFileDialog();
                saveFileDialog.Title = "保存";
                saveFileDialog.Filter = filter;
                saveFileDialog.FileName = filename;
                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    return saveFileDialog.FileName;
                }
                else
                    return null;
            }
        }
    }
}
