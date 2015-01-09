using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Threading;
using NPOI.HSSF.UserModel;//excel 97
using NPOI.SS.UserModel;

namespace taobaoINFO
{
    public delegate void xlsToDt(string filepath);
    public partial class Form1 : Form
    {
        private string filePath="";
        private Thread thread=null;
        private DataTable dt = new DataTable();
        private DataTable dt_2 = new DataTable();
        
        public Form1()
        {
            InitializeComponent();
        }

        private void button10_Click(object sender, EventArgs e)
        {
            thread = new Thread(this.start);
            thread.IsBackground = true;
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();
        }

        public class ListViewNF : System.Windows.Forms.ListView
        {
            public ListViewNF()
            {
                // 开启双缓冲
                this.SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint, true);

                // Enable the OnNotifyMessage event so we get a chance to filter out
                // Windows messages before they get to the form's WndProc
                this.SetStyle(ControlStyles.EnableNotifyMessage, true);
            }

            protected override void OnNotifyMessage(Message m)
            {
                //Filter out the WM_ERASEBKGND message
                if (m.Msg != 0x14)
                {
                    base.OnNotifyMessage(m);
                }
            }
        }

        void start()
        {
            button10.Enabled = false;
            button8.Enabled = false;
            string file="";
            if (textBox33.Text != "")
                file = textBox33.Text + "\n";
            if (filePath != "")
            {
                StreamReader sr = new StreamReader(filePath, false);
                file = sr.ReadToEnd().ToString();
                sr.Close();
            }
            if(file=="")
            {
                MessageBox.Show("请选择文件或输入url");
                button10.Enabled = true;
                button8.Enabled = true;
                thread.Abort();
            }
            string[] urls = file.Split('\n');
            clearresult();
            dt.Clear();
            int index = 1;
            int total = urls.Count();
            foreach (string url in urls)
            {
                if (url == "")
                    continue;
                List<string> result = new List<string>();
                ListViewItem list = new ListViewItem(index.ToString());
                DataRow dr = dt.NewRow();
                if (url.IndexOf("tmall") != -1)
                {
                    result.Add("天猫链接，无法查询！:"+url);
                    dr["title"] = "天猫链接，无法查询！:"+url;
                }
                else
                {
                    string html = getInfo.getHtml(url);
                    string id = getInfo.getID(url);
                    string sid = getInfo.getSID(html);
                    string sbn = getInfo.getSBN(html);
                    string keys = getInfo.getKEY(html);
                    if (getInfo.maijiagongju(url, ref result, ref dr) == true)//卖家工具信息
                    {
                        getInfo.wangwangid(html, ref result, ref dr); //旺旺id
                        getInfo.getPic(html, ref result, ref dr); //主图url
                        getInfo.getPrice(html, ref result, ref dr); //原价
                        getInfo.zhekoujia(ref result, ref dr); //折扣价
                        getInfo.paixiaMAX(ref result, ref dr); //交易价max
                        getInfo.paixiaMIN(ref result, ref dr);
                        getInfo.location(id, ref result, ref dr);//所在地及运费
                        getInfo.pingjia(id, sid, ref result, ref dr);//评价信息
                        getInfo.jiaoyijinfo(id, sid, sbn, keys, html, ref result, ref dr);//交易信息
                        getInfo.shuxing(html, ref result, ref dr);
                        getInfo.defen(html, ref result, ref dr);
                        result.Add(url); //标题url
                        dr["url"] = url;
                    }
                }
                dt.Rows.Add(dr);
                foreach (string res in result)
                    list.SubItems.Add(res);
                
                this.listView1.Items.Add(list);
                label1.Text = index++.ToString() + "/" + total.ToString();
                
            }
            dt_2 = dt.Copy();
            button10.Enabled = true;
            button8.Enabled = true;
            clearinput();
            clearfile();
        }
        void webstart()
        {
            this.webBrowser1.Navigate(listView1.SelectedItems[0].SubItems[38].Text);
        }
        private void clearinput()
        {
            textBox33.Text = "";
        }
        private void clearfile()
        {
            label24.Text = "已选文件：";
            filePath = "";
        }
        private void button1_Click(object sender, EventArgs e)
        {
            filePath = getInfo.getPath(0, "文本文件|*.txt");
            if (filePath == null)
            {
                return;
            }
            string fileName = filePath.Substring(filePath.LastIndexOf("\\") + 1);
            label24.Text = "已选文件：" + fileName;
        }

        private void pToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Clipboard.SetDataObject(listView1.SelectedItems[0].SubItems[1].Text);        
        }

        private void 打开淘宝ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Thread web = new Thread(this.webstart);
            web.IsBackground = true;
            web.SetApartmentState(ApartmentState.STA); 
            web.Start();
        }

        private void 删除该行ToolStripMenuItem_Click(object sender, EventArgs e)
        {
            dt_2.Rows.RemoveAt(listView1.SelectedItems[0].Index);
            this.listView1.Items.Clear();
            getInfo.Updatelist(dt_2,ref this.listView1);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (thread.ThreadState == ThreadState.Background)
            {
                button11.Text = "继续";
                thread.Suspend();
            }
            else
            {
                button11.Text = "暂停";
                thread.Resume();
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            clearinput();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            clearfile();
        }
        public void clearresult()
        {
            this.listView1.Items.Clear();
            dt.Clear();
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            this.listView1.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
            table.init(ref dt);
            this.listView1.Columns[1].Width=380;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            clearresult();

        }

        private void button8_Click(object sender, EventArgs e)
        {
            string outpath = getInfo.getPath(1, "excel文件.xls|*.xls", "结果.xls");
            if (outpath == null)
                return;
            HSSFWorkbook wk = new HSSFWorkbook();
            ISheet tb = wk.CreateSheet("查询结果");
            IRow exrow = tb.CreateRow(0);
            ICell cell;
            string[] head = new string[] { "序号", "宝贝标题", "所属类目", "标题分析结果", "标题关键词得分", "类目关联关键词词数", "标题关键词数量", "核心关键词", "高质量关键词", "掌柜ID", "主图", "原价", "折扣价", "拍下最高价", "拍下最低价", "所在地", "运费", "累计评论数量", "好评数量", "中评数量", "差评数量", "追加评论数量", "评价有图片数量", "创建时间", "下架时间", "交易成功数量", "成交纪录数量", "浏览量", "收藏量", "创建天数", "日均成交件数", "日均浏览量", "转化率", "收藏率", "产品属性", "描述评分", "物流评分", "服务评分", "宝贝URL" };
            int count = head.Count();
            int i;
            int index = 1;
            for (i = 0; i < count; i++)
            {
                cell = exrow.CreateCell(i);
                cell.SetCellValue(head[i]);
            }
            foreach (DataRow dtrow in dt_2.Rows)
            {
                exrow = tb.CreateRow(index);
                cell = exrow.CreateCell(0);
                cell.SetCellValue(index++);
                int columns = dt.Columns.Count;
                for (i = 1; i <= columns; i++)
                {
                    cell = exrow.CreateCell(i);
                    if (dt.Columns[i - 1].DataType == System.Type.GetType("System.String"))
                    {
                        cell.SetCellValue(dtrow[i - 1].ToString());
                    }
                    else if (dt.Columns[i - 1].DataType == System.Type.GetType("System.Int32"))
                    {
                        if(dtrow[i - 1].ToString()!="")
                            cell.SetCellValue(int.Parse(dtrow[i - 1].ToString()));
                    }
                    else if (dt.Columns[i - 1].DataType == System.Type.GetType("System.Double"))
                    {
                        if (dtrow[i - 1].ToString() != "")
                            cell.SetCellValue(double.Parse(dtrow[i - 1].ToString()));
                    }
                }
            }
            FileStream fs = File.OpenWrite(outpath);
            wk.Write(fs);
            fs.Close();
            MessageBox.Show("导出完成！");
        }

        private void button7_Click(object sender, EventArgs e)
        {
            string inpath = getInfo.getPath(0, "excel文件|*.xls");      
            if (inpath == null)
                return;
            xlsToDt task = excelToDatatable;
            IAsyncResult asyncResult = task.BeginInvoke(inpath, null, null);
            panel1.Visible = true;
            //while (!asyncResult.AsyncWaitHandle.WaitOne(100, false))
            //{
            //    panel1.Visible = true;
            //}

            //panel1.Visible = false;
            
            //excelToDatatable(inpath);
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (dt_2.Rows.Count == 0)
            {
                MessageBox.Show("没有数据，无法筛选！");
                return;
            }
            Form2 f2 = new Form2(dt);
            if (f2.ShowDialog() == DialogResult.OK)
            {
                List<string> temp = new List<string>();
                foreach (string word in f2.filter)
                {
                    if (word != "")
                    {
                        dt_2 = dt_2.Select("title not like '%" + word + "%'").CopyToDataTable();
                    }
                }
                if (f2.title_key_num_L != "")
                {
                    temp.Add("title_key_num>=" + f2.title_key_num_L);
                }
                if (f2.title_key_num_R != "")
                {
                    temp.Add("title_key_num<=" + f2.title_key_num_R);
                }
                if (f2.title_len_L != "")
                {
                    temp.Add("len(title)>=" + f2.title_len_L);
                }
                if (f2.title_len_R != "")
                {
                    temp.Add("len(title)<=" + f2.title_len_R);
                }
                if (f2.title_key_score_L != "")
                {
                    temp.Add("title_key_score>=" + f2.title_key_score_L);
                }
                if (f2.title_key_score_R != "")
                {
                    temp.Add("title_key_score<=" + f2.title_key_score_R);
                }
                if (f2.title_res != "")
                {
                    temp.Add("title_res='" + f2.title_res + "'");
                }
                if (f2.category != "")
                {
                    temp.Add("category='" + f2.category + "'");
                }
                if (f2.location != "")
                {
                    temp.Add("location='" + f2.location + "'");
                }
                if(f2.success_L!="")
                {
                    temp.Add("success>=" + f2.success_L);
                }
                if (f2.success_R != "")
                {
                    temp.Add("success<=" + f2.success_R);
                }
                if (f2.deal_oneday_L != "")
                {
                    temp.Add("deal_oneday>=" + f2.deal_oneday_L);
                }
                if (f2.deal_oneday_R != "")
                {
                    temp.Add("deal_oneday<=" + f2.deal_oneday_R);
                }
                if (f2.view_oneday_L != "")
                {
                    temp.Add("view_oneday>=" + f2.view_oneday_L);
                }
                if (f2.view_oneday_R != "")
                {
                    temp.Add("view_oneday<=" + f2.view_oneday_R);
                }
                if (f2.con_rate_L != "")
                {
                    temp.Add("con_rate>=" + f2.con_rate_L);
                }
                if (f2.con_rate_R != "")
                {
                    temp.Add("con_rate<=" + f2.con_rate_R);
                }
                if (f2.coll_rate_L!="")
                {
                    temp.Add("coll_rate>=" + f2.coll_rate_L);
                }
                if (f2.coll_rate_R != "")
                {
                    temp.Add("coll_rate<=" + f2.coll_rate_R);
                }
                if (f2.price_L != "")
                {
                    temp.Add("price>=" + f2.price_L);
                }
                if (f2.price_R != "")
                {
                    temp.Add("price<=" + f2.price_R);
                }

                if (f2.core_key_num_L != "")
                {
                    temp.Add("core_key_num>=" + f2.core_key_num_L);
                }
                if (f2.core_key_num_R != "")
                {
                    temp.Add("core_key_num<=" + f2.core_key_num_R);
                }
                if (f2.high_key_L != "")
                {
                    temp.Add("high_key>=" + f2.high_key_L);
                }
                if (f2.high_key_R != "")
                {
                    temp.Add("high_key<=" + f2.high_key_R);
                }
                if (f2.cate_key_num_L != "")
                {
                    temp.Add("cate_key_num>=" + f2.cate_key_num_L);
                }
                if (f2.cate_key_num_R != "")
                {
                    temp.Add("cate_key_num<=" + f2.cate_key_num_R);
                }
                if (f2.is_free_freight)
                {
                    temp.Add("freight='包邮'");
                }
                if (f2.not_free_freight)
                {
                    temp.Add("freight='不包邮'");
                }
                string[] queries = temp.ToArray();
                string query="";
                query = string.Join(" and ", queries);
                DataRow[] res = dt_2.Select(query);
                if (res.Count() > 0)
                    dt_2 = res.CopyToDataTable();
                else
                    dt_2.Clear();
                this.listView1.Items.Clear();
                getInfo.Updatelist(dt_2, ref this.listView1);
            }
            //f2.Show();
        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Form3 f3 = new Form3(listView1.SelectedItems[0]);
            f3.Show();
        }
        private void excelToDatatable(string filepath)
        {
            int i,j;
            ISheet sheet = null;
            FileStream fs = null;
            IWorkbook workbook = null;
            fs = new FileStream(filepath, FileMode.OpenOrCreate, FileAccess.ReadWrite);
            workbook = new HSSFWorkbook(fs);
            sheet = workbook.GetSheetAt(0);
            if (sheet != null)
            {
                int col_Count = sheet.GetRow(0).LastCellNum; ;
                int row_Count = sheet.LastRowNum;
                clearresult();
                for (i = 1; i <= row_Count; i++)
                {
                    IRow row = sheet.GetRow(i);
                    DataRow dr = dt.NewRow();
                    for (j = 1; j < col_Count; j++)
                    {
                        if(dt.Columns[j-1].DataType==System.Type.GetType("System.String"))
                            dr[j-1] = row.GetCell(j).ToString();
                        if(dt.Columns[j-1].DataType==System.Type.GetType("System.Int32"))
                            dr[j-1] = int.Parse(row.GetCell(j).ToString());
                        if(dt.Columns[j-1].DataType==System.Type.GetType("System.Double"))
                            dr[j-1] = double.Parse(row.GetCell(j).ToString());
                        if (row.GetCell(j).ToString().IndexOf("天猫链接") != -1 || row.GetCell(j).ToString().IndexOf("链接失效") != -1)
                            break;
                    }
                    dt.Rows.Add(dr);
                }
            }
            dt_2 = dt.Copy();
            getInfo.Updatelist(dt_2, ref this.listView1);
            panel1.Visible = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            dt_2 = dt.Copy();
            this.listView1.Items.Clear();
            getInfo.Updatelist(dt, ref this.listView1);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            textBox32.Text = "";
        }

    }
}