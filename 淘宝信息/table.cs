using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace taobaoINFO
{
    class table
    {
        static public void init(ref DataTable dt)
        {
            dt.Columns.Add("title", typeof(string));
            dt.Columns.Add("category", typeof(string));
            dt.Columns.Add("title_res", typeof(string));
            dt.Columns.Add("title_key_score", typeof(int));
            dt.Columns.Add("cate_key_num", typeof(int));
            dt.Columns.Add("title_key_num", typeof(int));
            dt.Columns.Add("core_key_num", typeof(int));
            dt.Columns.Add("high_key", typeof(int));
            dt.Columns.Add("id", typeof(string));
            dt.Columns.Add("pic", typeof(string));
            dt.Columns.Add("price", typeof(double));
            dt.Columns.Add("price_now", typeof(double));
            dt.Columns.Add("max_price", typeof(double));
            dt.Columns.Add("min_price", typeof(double));
            dt.Columns.Add("location", typeof(string));
            dt.Columns.Add("freight", typeof(string));
            dt.Columns.Add("review", typeof(int));
            dt.Columns.Add("good_review", typeof(int));
            dt.Columns.Add("normal_review", typeof(int));
            dt.Columns.Add("bad_review", typeof(int));
            dt.Columns.Add("add_review", typeof(int));
            dt.Columns.Add("pic_review", typeof(int));
            dt.Columns.Add("create_time", typeof(string));
            dt.Columns.Add("end_time", typeof(string));
            dt.Columns.Add("deal_success", typeof(int));
            dt.Columns.Add("success", typeof(int));
            dt.Columns.Add("view", typeof(int));
            dt.Columns.Add("collection", typeof(int));
            dt.Columns.Add("create_days", typeof(int));
            dt.Columns.Add("deal_oneday", typeof(int));
            dt.Columns.Add("view_oneday", typeof(int));
            dt.Columns.Add("con_rate", typeof(double));
            dt.Columns.Add("coll_rate", typeof(double));
            dt.Columns.Add("attribute", typeof(string));
            dt.Columns.Add("describe", typeof(double));
            dt.Columns.Add("logistics", typeof(double));
            dt.Columns.Add("service", typeof(double));
            dt.Columns.Add("url", typeof(string));
        }
    }
}
