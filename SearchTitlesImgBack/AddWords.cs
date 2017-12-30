using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SearchTitlesImgBack
{
    public partial class AddWords : Form
    {
        private static Dictionary<Button, string> btnStrs = new Dictionary<Button, string>();
        private static string addButtenTag = "";
        public AddWords()
        {
            InitializeComponent();
        }

        private void AddWords_Load(object sender, EventArgs e)
        {
            btnStrs.Add(button1, "SerachTitles_Color");
            btnStrs.Add(button2, "SerachTitles_Category");
            btnStrs.Add(button3, "SerachTitles_Purpose");
            btnStrs.Add(button4, "SerachTitles_Scene");
            btnStrs.Add(button5, "SerachTitles_Style");
            btnStrs.Add(button6, "SerachTitles_Industry");
            foreach (var btnStr in btnStrs) 
            {
                (btnStr.Key).Click += new EventHandler(upBtnClick);
            }
        }
        private void button8_Click(object sender, EventArgs e)//添加按钮
        {
            string text = this.textBox2.Text.Replace(" ", "");
            text = text.Trim();
            if (text == "")
            {
                MessageBox.Show("请输入添加内容");
            }
            int resultCount = AddTag(text);
            if (resultCount != 0)
            {
                MessageBox.Show("添加成功");
                this.textBox2.Text = "";
                foreach (var btnStr in btnStrs) 
                {
                    if (btnStr.Value == addButtenTag) 
                    {
                        (btnStr.Key).PerformClick();
                    }
                }
            }
            else
            {
                MessageBox.Show("添加失败");
            }
        }
        private void button7_Click(object sender, EventArgs e)//删除按钮
        {
            string[] lists = this.textBox1.Text.Split(new char[] { '|' });
            string[] colum = addButtenTag.Split(new char[] { '_' });
            string sqlStr = "delete from " + addButtenTag + " where " + colum[1] + " in" + " (";
            foreach (var list in lists)
            {
                sqlStr = sqlStr + "'" + list + "',";
            }
            sqlStr = sqlStr.Substring(0, sqlStr.Length - 4) + ")";
            int resultCount = DBHelper.ExecuteUpdate(sqlStr);
            if (resultCount == 0)
            {
                MessageBox.Show("删除失败");
            }
            foreach (var btnStr in btnStrs)
            {
                if (btnStr.Value == addButtenTag)
                {
                    (btnStr.Key).PerformClick();
                }
            }
            this.textBox1.Text = "";
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
                    addLabel_p1(dt);
                }
            }
        }
        /// <summary>
        /// 将查询到的数据变成lable显示到界面
        /// </summary>
        /// <param name="dt"></param>
        private void addLabel_p1(DataTable dt) 
        {
            this.textBox1.Text = "";
            this.flowLayoutPanel1.Controls.Clear();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                Label onlineUser = new Label();
                onlineUser.Width = 160;
                onlineUser.Margin = new System.Windows.Forms.Padding(1, 1, 1, 1);
                onlineUser.BackColor = Color.Black;
                onlineUser.ForeColor = Color.Gold;
                onlineUser.TextAlign = ContentAlignment.MiddleCenter;
                onlineUser.Text += dt.Rows[i][0].ToString();
                onlineUser.Click += new EventHandler(onlineUser_click_p1);
                this.flowLayoutPanel1.Controls.Add(onlineUser);
            }
        }
        /// <summary>
        /// 加词界面，label点击事件 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void onlineUser_click_p1(object sender, EventArgs e)
        {
            Label lbl = (Label)(sender);
            this.textBox1.Text = this.textBox1.Text + lbl.Text + "|";
        }
        /// <summary>
        /// 加词检测
        /// </summary>
        /// <param name="tag"></param>
        /// <returns></returns>
        private int AddTag(string tag)
        {
            if (tag.Contains("|"))
            {
                MessageBox.Show("字符中不允许有： |  ");
                return 0;
            }
            if (addButtenTag != "")
            {
                string sqlstr = "INSERT INTO " + addButtenTag + " VALUES ('" + tag + "')";
                return DBHelper.ExecuteUpdate(sqlstr);
            }
            else
            {
                return 0;
            }
        }
    }
}
