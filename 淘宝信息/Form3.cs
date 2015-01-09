using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace taobaoINFO
{
    public partial class Form3 : Form
    {
        public Form3(ListViewItem list)
        {
            InitializeComponent();
            pictureBox1.ImageLocation = list.SubItems[10].Text;
            label1.Text += list.SubItems[1].Text;
            label2.Text += list.SubItems[2].Text;
            label3.Text += list.SubItems[9].Text;
            label4.Text += list.SubItems[11].Text;
            label5.Text += list.SubItems[15].Text;
            label6.Text += list.SubItems[16].Text;
            label7.Text += list.SubItems[23].Text;
            label8.Text += list.SubItems[24].Text;
            label9.Text += list.SubItems[27].Text;
            label10.Text += list.SubItems[28].Text;
            label11.Text += list.SubItems[25].Text;
            label12.Text += list.SubItems[26].Text;
            label13.Text += list.SubItems[30].Text;
            label14.Text += list.SubItems[31].Text;
            label15.Text += list.SubItems[32].Text;
            label16.Text += list.SubItems[33].Text;
            label17.Text += list.SubItems[17].Text;
            label18.Text += list.SubItems[18].Text;
            label19.Text += list.SubItems[19].Text;
            label20.Text += list.SubItems[20].Text;
            label21.Text += list.SubItems[21].Text;
            label22.Text += list.SubItems[22].Text;
            label23.Text += list.SubItems[35].Text;
            label24.Text += list.SubItems[36].Text;
            label25.Text += list.SubItems[37].Text;
            label26.Text += list.SubItems[3].Text;
            label27.Text += list.SubItems[4].Text;
            label28.Text += list.SubItems[5].Text;
            label29.Text += list.SubItems[6].Text;
            label30.Text += list.SubItems[7].Text;
            label31.Text += list.SubItems[8].Text;
            label32.Text = list.SubItems[34].Text.Replace("/", "    ");
        }

    }
}
