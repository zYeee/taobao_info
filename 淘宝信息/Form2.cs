using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace taobaoINFO
{
    public partial class Form2 : Form
    {
        public Form2(DataTable dt)
        {
            InitializeComponent();
            DataTable newdt = SelectDistinct(dt, "location");
            bool flag = false;
            foreach (DataRow dr in newdt.Rows)
            {
                comboBox1.Items.Add(dr[0]);
                if(dr[0].ToString()=="")
                    flag = true;
            }
            if (!flag)
            {
                comboBox1.Items.Add("");
            }

            flag = false;
            newdt = SelectDistinct(dt, "title_res");
            foreach (DataRow dr in newdt.Rows)
            {
                comboBox2.Items.Add(dr[0]);
                if (dr[0].ToString() == "")
                    flag = true;
            }
            if (!flag)
            {
                comboBox2.Items.Add("");
            }

            flag = false;
            newdt = SelectDistinct(dt, "category");
            foreach (DataRow dr in newdt.Rows)
            {
                comboBox3.Items.Add(dr[0]);
                if (dr[0].ToString() == "")
                    flag = true;
            }
            if (!flag)
            {
                comboBox3.Items.Add("");
            }
        }
        private DataTable SelectDistinct(DataTable SourceTable, params string[] FieldNames)
        {
            object[] lastValues;
            DataTable newTable;
            DataRow[] orderedRows;

            if (FieldNames == null || FieldNames.Length == 0)
                throw new ArgumentNullException("FieldNames");

            lastValues = new object[FieldNames.Length];
            newTable = new DataTable();

            foreach (string fieldName in FieldNames)
                newTable.Columns.Add(fieldName, SourceTable.Columns[fieldName].DataType);

            orderedRows = SourceTable.Select("", string.Join(",", FieldNames));

            foreach (DataRow row in orderedRows)
            {
                if (!fieldValuesAreEqual(lastValues, row, FieldNames))
                {
                    newTable.Rows.Add(createRowClone(row, newTable.NewRow(), FieldNames));

                    setLastValues(lastValues, row, FieldNames);
                }
            }

            return newTable;
        }

        private bool fieldValuesAreEqual(object[] lastValues, DataRow currentRow, string[] fieldNames)
        {
            bool areEqual = true;

            for (int i = 0; i < fieldNames.Length; i++)
            {
                if (lastValues[i] == null || !lastValues[i].Equals(currentRow[fieldNames[i]]))
                {
                    areEqual = false;
                    break;
                }
            }

            return areEqual;
        }

        private DataRow createRowClone(DataRow sourceRow, DataRow newRow, string[] fieldNames)
        {
            foreach (string field in fieldNames)
                newRow[field] = sourceRow[field];

            return newRow;
        }

        private void setLastValues(object[] lastValues, DataRow sourceRow, string[] fieldNames)
        {
            for (int i = 0; i < fieldNames.Length; i++)
                lastValues[i] = sourceRow[fieldNames[i]];
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.OK;
            pass();
            this.Close();
        }
        private void pass()
        {
            if (comboBox4.SelectedIndex > 0)
            {
                switch (comboBox4.SelectedIndex)
                {
                    case 1: title_len_L = "20"; title_len_R = "30"; break;
                    case 2: title_len_L = "25"; title_len_R = "30"; break;
                    case 3: title_len_L = "28"; title_len_R = "30"; break;
                    case 4: title_len_L = "29"; title_len_R = "30"; break;
                }
            }
            title_key_num_L = textBox21.Text;
            title_key_num_R = textBox20.Text;
            title_key_score_L = textBox17.Text;
            title_key_score_R = textBox16.Text;
            title_res = comboBox2.Text;
            location = comboBox1.Text;
            category = comboBox3.Text;
            if (comboBox6.SelectedIndex > 0)
            {
                switch (comboBox6.SelectedIndex)
                {
                    case 1: success_L = "0"; success_R = "10"; break;
                    case 2: success_L = "10"; success_R = "50"; break;
                    case 3: success_L = "0"; success_R = "100"; break;
                    case 4: success_L = "100"; success_R = "300"; break;
                    case 5: success_L = "300"; success_R = "500"; break;
                    case 6: success_L = "500"; success_R = "1000"; break;
                    case 7: success_L = "1000"; success_R = "2000"; break;
                    case 8: success_L = "2000"; success_R = "5000"; break;
                    case 9: success_L = "5000";  break;
                }
            }

            if (comboBox5.SelectedIndex > 0)
            {
                switch (comboBox5.SelectedIndex)
                {
                    case 1: deal_oneday_L = "0"; deal_oneday_R = "5"; break;
                    case 2: deal_oneday_L = "5"; deal_oneday_R = "10"; break;
                    case 3: deal_oneday_L = "10"; deal_oneday_R = "15"; break;
                    case 4: deal_oneday_L = "15"; deal_oneday_R = "20"; break;
                    case 5: deal_oneday_L = "20"; deal_oneday_R = "30"; break;
                    case 6: deal_oneday_L = "30"; deal_oneday_R = "50"; break;
                    case 7: deal_oneday_L = "50"; break;
                }
            }
            view_oneday_L = textBox11.Text;
            view_oneday_R = textBox10.Text;
            if (comboBox7.SelectedIndex > 0)
            {
                switch (comboBox7.SelectedIndex)
                {
                    case 1: con_rate_L = "0"; con_rate_R = "1"; break;
                    case 2: con_rate_R = "1"; con_rate_R = "2"; break;
                    case 3: con_rate_L = "2"; con_rate_R = "3"; break;
                    case 4: con_rate_L = "3"; con_rate_R = "4"; break;
                    case 5: con_rate_L = "4"; con_rate_R = "5"; break;
                    case 6: con_rate_L = "5"; con_rate_R = "6"; break;
                    case 7: con_rate_L = "6"; break;
                }
            }
            if (comboBox8.SelectedIndex > 0)
            {
                switch (comboBox8.SelectedIndex)
                {
                    case 1: coll_rate_L = "0"; coll_rate_R = "2"; break;
                    case 2: coll_rate_L = "2"; coll_rate_R = "5"; break;
                    case 3: coll_rate_L = "5"; coll_rate_R = "8"; break;
                    case 4: coll_rate_L = "8"; coll_rate_R = "10"; break;
                    case 5: coll_rate_L = "10"; coll_rate_R = "15"; break;
                    case 6: coll_rate_L = "15"; coll_rate_R = "20"; break;
                    case 7: coll_rate_L = "20"; break;
                }
            }
            price_L = textBox4.Text;
            price_R = textBox5.Text;
            core_key_num_L = textBox25.Text;
            core_key_num_R = textBox24.Text;
            high_key_L = textBox23.Text;
            high_key_R = textBox22.Text;
            cate_key_num_L = textBox19.Text;
            cate_key_num_R = textBox18.Text;
            is_free_freight = radioButton1.Checked;
            not_free_freight = radioButton2.Checked;
            filter = textBox1.Text.Replace("\r\n", "@").Split('@');
        }
        public string title_len_L="";
        public string title_len_R="";
        public string title_key_num_L="";
        public string title_key_num_R="";
        public string title_key_score_L="";
        public string title_key_score_R="";
        public string title_res="";
        public string success_L="";
        public string success_R="";
        public string deal_oneday_L="";
        public string deal_oneday_R="";
        public string view_oneday_L="";
        public string view_oneday_R="";
        public string con_rate_L="";
        public string con_rate_R="";
        public string coll_rate_L="";
        public string coll_rate_R="";
        public string price_L="";
        public string price_R="";
        public string cate_key_num_L=""; //类目关联关键词
        public string cate_key_num_R="";
        public string core_key_num_L=""; //核心关键词数量
        public string core_key_num_R="";
        public string high_key_L=""; //高质量关键词数量
        public string high_key_R="";
        public string freight="";//运费
        public string create_time_L="";
        public string create_time_R="";
        public string location="";
        public string category="";
        public bool is_free_freight = false;
        public bool not_free_freight = false;
        public string[] filter;

        private void button3_Click(object sender, EventArgs e)
        {
            string inpath = getInfo.getPath(0, "文本文件|*.txt");
            if (inpath == null)
                return;
            StreamReader sr = new StreamReader(inpath, Encoding.Default);
            string word = sr.ReadToEnd().ToString();
            sr.Close();
            textBox1.Text = word;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            string outpath = getInfo.getPath(1, "文本文件|*.txt","词表");
            if (outpath == null)
                return;
            StreamWriter sw = new StreamWriter(outpath, false, Encoding.Default);
            sw.Write(textBox1.Text);
            sw.Close();
        }
    }
}
