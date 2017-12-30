using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using CsharpHttpHelper;
using HtmlAgilityPack;
using System.Net;
using System.IO;

namespace SearchTitlesImgBack
{
    public partial class SelectWords : Form
    {
        private static Dictionary<Button, string> btnStrs = new Dictionary<Button, string>();
        public static List<string> titles1 = new List<string>();//点击时长尾词存放集合
        public static List<string> titles2 = new List<string>();//点击时标题存放集合
        public static List<string> jxTitles = new List<string>();//解析时标题存放集合
        public static string[] jxRelated;
        public static string addButtenTag = "";
        public SelectWords()
        {
            InitializeComponent();
        }

        private void SelectWords_Load(object sender, EventArgs e)
        {
            btnStrs.Add(button2, "SerachTitles_Color");
            btnStrs.Add(button3, "SerachTitles_Category");
            btnStrs.Add(button4, "SerachTitles_Purpose");
            btnStrs.Add(button5, "SerachTitles_Scene");
            btnStrs.Add(button6, "SerachTitles_Style");
            btnStrs.Add(button7, "SerachTitles_Industry");
            foreach (var btnStr in btnStrs)
            {
                (btnStr.Key).Click += new EventHandler(upBtnClick);
            }
        }
        /// <summary>
        /// 上部按钮点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void upBtnClick(object sender, EventArgs e)
        {
            foreach (var btnStr in btnStrs)
            {
                if (btnStr.Key.Name == (sender as Button).Name)
                {
                    Tools.ChangeForeColor(btnStr.Key);
                    addButtenTag = btnStr.Value;
                    string sqlStr = "select distinct * from " + btnStr.Value;
                    DataTable dt = DBHelper.ExecuteQuery(sqlStr);
                    addLabel_p1_p2(dt);
                }
            }
        }
        /// <summary>
        /// 将查询到的数据变成lable显示到界面1和界面2
        /// </summary>
        /// <param name="dt"></param>
        private void addLabel_p1_p2(DataTable dt) 
        {
            this.flowLayoutPanel1.Controls.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                Label onlineUser1 = new Label();
                onlineUser1.Width = 160;
                onlineUser1.Margin = new System.Windows.Forms.Padding(1, 1, 1, 1);
                onlineUser1.BackColor = Color.Black;
                onlineUser1.ForeColor = Color.Gold;
                onlineUser1.TextAlign = ContentAlignment.MiddleCenter;
                onlineUser1.Text += dt.Rows[i][0].ToString();
                onlineUser1.Click += new EventHandler(onlineUser_click_p1);
                this.flowLayoutPanel1.Controls.Add(onlineUser1);
            }
            this.flowLayoutPanel2.Controls.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                Label onlineUser2 = new Label();
                onlineUser2.Width = 160;
                onlineUser2.Margin = new System.Windows.Forms.Padding(1, 1, 1, 1);
                onlineUser2.BackColor = Color.Black;
                onlineUser2.ForeColor = Color.Gold;
                onlineUser2.TextAlign = ContentAlignment.MiddleCenter;
                onlineUser2.Text += dt.Rows[i][0].ToString();
                onlineUser2.Click += new EventHandler(onlineUser_click_p2);
                this.flowLayoutPanel2.Controls.Add(onlineUser2);
            }
        }
        /// <summary>
        /// 标题框中的label点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void onlineUser_click_p1(object sender, EventArgs e)
        {
            string titleStr = "";
            Label lbl = (Label)(sender);
            if (lbl.BackColor == Color.Black)
            {
                lbl.BackColor = Color.Red;
                titles2.Add(lbl.Text);
                foreach (var title in titles2)
                {
                    titleStr = titleStr + title;
                }
                this.textBox2.Text = titleStr;
            }
            else
            {
                lbl.BackColor = Color.Black;
                titles2.Remove(lbl.Text);
                foreach (var title in titles2)
                {
                    titleStr = titleStr + title;
                }
                this.textBox2.Text = titleStr;
            }
        }
        /// <summary>
        /// 长尾词框中的label点击事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void onlineUser_click_p2(object sender, EventArgs e)
        {
            string titleStr = "";
            Label lbl = (Label)(sender);
            if (lbl.BackColor == Color.Black)
            {
                lbl.BackColor = Color.Red;
                titles1.Add(lbl.Text);
                foreach (var title in titles1)
                {
                    titleStr = titleStr + title + " ";
                }
                this.textBox3.Text = titleStr;
            }
            else
            {
                lbl.BackColor = Color.Black;
                titles1.Remove(lbl.Text);
                foreach (var title in titles1)
                {
                    titleStr = titleStr + title + " ";
                }
                this.textBox3.Text = titleStr;
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Tools.ChangeForeColor(this.button1);
            addCGLabe(jxRelated, jxTitles);
        }
        /// <summary>
        /// 将搜索到的数据变成lable显示到界面1和界面2
        /// </summary>
        /// <param name="BaiduXLList"></param>
        /// <param name="BaiduSSList"></param>
        private void addCGLabe(string[] BaiduXLList, List<string> BaiduSSList) 
        {
            if (BaiduXLList == null || BaiduSSList == null)
            {
                return;
            }
            this.flowLayoutPanel1.Controls.Clear();
            this.flowLayoutPanel2.Controls.Clear();
            for (int i = 0; i < BaiduXLList.Length; i++)
            {
                Label onlineUser = new Label();
                onlineUser.Width = 160;
                onlineUser.Margin = new System.Windows.Forms.Padding(1, 1, 1, 1);
                onlineUser.BackColor = Color.Black;
                onlineUser.ForeColor = Color.White;
                onlineUser.TextAlign = ContentAlignment.MiddleCenter;
                onlineUser.Text += BaiduXLList[i];
                onlineUser.Click += new EventHandler(onlineUser_click_p1);
                this.flowLayoutPanel1.Controls.Add(onlineUser);
            }
            for (int i = 0; i < BaiduSSList.Count; i++)
            {
                Label onlineUser = new Label();
                onlineUser.Width = 160;
                onlineUser.Margin = new System.Windows.Forms.Padding(1, 1, 1, 1);
                onlineUser.BackColor = Color.Black;
                onlineUser.ForeColor = Color.White;
                onlineUser.TextAlign = ContentAlignment.MiddleCenter;
                onlineUser.Text += BaiduSSList[i];
                onlineUser.Click += new EventHandler(onlineUser_click_p2);
                this.flowLayoutPanel2.Controls.Add(onlineUser);
            }
        }

        private void button8_Click(object sender, EventArgs e)//搜索
        {
            Tools.ChangeForeColor(this.button8);
            jxTitles.Clear();
            this.flowLayoutPanel1.Controls.Clear();
            this.flowLayoutPanel2.Controls.Clear();
            string baiduXLUrl = "http://nssug.baidu.com/?&wd=" + this.textBox1.Text + "&prod=image";//cb=jQuery111102218013105683504_1480310835355&ie=utf-8&t=0.9119490526483085&_=1580310835360
            string baiduSSUrl = "http://image.baidu.com/search/index?tn=baiduimage&word=" + this.textBox1.Text;//&ipn=r&ct=201326592&cl=2&lm=-1&st=-1&fm=result&fr=&sf=1&fmq=1480312719899_R&pv=&ic=0&nc=1&z=&se=1&showtab=0&fb=0&width=&height=&face=0&istype=2&ie=utf-8&ctd=1480312719899^00_1903X182
            string baiduXLRes = "";
            string baiduSSRes = "";
            if(this.checkBox1.Checked)
            {
                baiduXLRes = GetHtml(baiduXLUrl,"百度下拉词2秒未返回请求");
            } 
            if (this.checkBox2.Checked)
            {
                baiduSSRes = GetHtml(baiduSSUrl,"百度搜索词2秒未返回请求");
            } 
            
            string[] BaiduXLList = BaiduXLToList(baiduXLRes);
            BaiduSSToList(baiduSSRes);

            
            string XL360URL = "http://sug.image.so.com/suggest/word?callback=suggest_so&encodein=utf-8&encodeout=utf-8&word=" + this.textBox1.Text;
            string SS360URL = "http://image.so.com/i?q=" + this.textBox1.Text + "&src=srp";
            string XL360Res = "";
            string SS360Res = "";
            if (this.checkBox3.Checked)
            {
                XL360Res = GetHtml(XL360URL, "360下拉词2秒未返回请求");
            }
            if (this.checkBox4.Checked)
            {
                SS360Res = GetHtml(SS360URL, "360搜索词2秒未返回请求");
            }
            string[] XL360List = BaiduXLToList(XL360Res);
            SS360ToList(SS360Res);

            string[] allList = BaiduXLList.Concat(XL360List).ToArray();
            allList = allList.Distinct().ToArray();
            jxRelated = allList;
            if (allList == null || jxTitles == null)
            {
                MessageBox.Show("未获取到数据，请重试！");
                return;
            }
            addCGLabe(allList, jxTitles);
        }
        /// <summary>
        /// 请求发送，获取源码
        /// </summary>
        /// <param name="url"></param>
        /// <returns></returns>
        public static string GetHtml(string url,string mes)
        {
            HttpHelper http = new HttpHelper();
            HttpItem item = new HttpItem()
            {
                URL = url,
                Allowautoredirect = false,
                KeepAlive = false,
                Timeout = 2000,//连接超时时间
            };
            try
            {
                HttpResult result = http.GetHtml(item);
                string htmlText = result.Html;
                return htmlText;
            }
            catch 
            {
                MessageBox.Show(mes);
                return "";
            }
        }
        /// <summary>
        /// 解析百度下拉词
        /// </summary>
        /// <param name="baiduXL"></param>
        /// <returns></returns>
        private string[] BaiduXLToList(string baiduXL)
        {

            baiduXL = baiduXL.Replace("\"", "");
            string[] baiduXLStr = baiduXL.Split(new char[] { '[', ']' });
            try
            {
                string[] baiduXLList = baiduXLStr[1].Split(new char[] { ',' });
                return baiduXLList;
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
                return baiduXLStr;
            }
        }
        /// <summary>
        /// 解析百度搜索词
        /// </summary>
        /// <param name="baiduSS">页面源码</param>
        /// <returns></returns>
        private List<string> BaiduSSToList(string baiduSS)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(baiduSS);
            HtmlNode node = doc.DocumentNode;
            var aTags = node.SelectNodes("//a[@class='pull-rs']");
            if (aTags == null)
            {
                return jxTitles;
            }
            foreach (var tag in aTags)
            {
                jxTitles.Add(tag.InnerHtml);
            }
            return jxTitles;
        }
        /// <summary>
        /// 解析360搜索词
        /// </summary>
        /// <param name="SS360">页面源码</param>
        /// <returns></returns>
        private List<string> SS360ToList(string SS360)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(SS360);
            HtmlNode node = doc.DocumentNode;
            var aTags = node.SelectNodes("//a");
            if (aTags == null)
            {
                return jxTitles;
            }
            foreach (var aTag in aTags)
            {
                if (aTag.Attributes["data-index"] != null)
                {
                    jxTitles.Add(aTag.InnerHtml);
                }
            }
            return jxTitles;
        }
        private void button9_Click(object sender, EventArgs e)//清空搜索
        {
            titles1.Clear();
            titles2.Clear();
            this.textBox2.Text = "";
            this.textBox3.Text = "";
            button8_Click(null, null);
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (this.textBox2.Text != "")
                Clipboard.SetDataObject(this.textBox2.Text);
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (this.textBox3.Text != "")
                Clipboard.SetDataObject(this.textBox3.Text);
        }

        private void button11_Click(object sender, EventArgs e)
        {
            titles2.Clear();
            this.textBox2.Text = "";
            foreach (Control tmpControl in this.flowLayoutPanel1.Controls)
            {
                if (tmpControl is Label)
                {
                    tmpControl.BackColor = Color.Black;
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            titles1.Clear();
            this.textBox3.Text = "";
            foreach (Control tmpControl in this.flowLayoutPanel2.Controls)
            {
                if (tmpControl is Label)
                {
                    tmpControl.BackColor = Color.Black;
                }
            }
        }
    }
}
