using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace SearchTitlesImgBack
{
    public partial class DeductionCount : Form
    {
        public DeductionCount()
        {
            InitializeComponent();
        }
        private void DeductionCount_Load(object sender, EventArgs e)
        {
            this.dateTimePicker1.Value = System.DateTime.Now.AddDays(-7);
            this.dateTimePicker2.Value = System.DateTime.Now;
            if (Login.userRoles == "IT传媒部-组长")
            {
                DataTable dtUserSmall = frmMain.dtUser.Copy();
                dtUserSmall.Rows.Add("");
                this.comboBox1.DataSource = dtUserSmall;
                this.comboBox1.DisplayMember = "UserId";
                this.comboBox1.ValueMember = "UserId";
            }
            else 
            {
                this.linkLabel1.Visible = false;
                this.comboBox1.Text = Login.userName;
                this.comboBox1.Enabled = false;
                this.button1.Visible = false;
                this.textBox1.Enabled = false;
                this.textBox2.Enabled = false;
                this.textBox3.Enabled = false;
                this.textBox4.Enabled = false;
                this.textBox5.Enabled = false;
                this.textBox6.Enabled = false;

                this.label1.Location = new Point(this.label1.Location.X + 80, this.label1.Location.Y);
                this.label2.Location = new Point(this.label2.Location.X + 80, this.label2.Location.Y);
                this.label3.Location = new Point(this.label3.Location.X + 80, this.label3.Location.Y);
                this.label4.Location = new Point(this.label4.Location.X + 80, this.label4.Location.Y);
                this.label5.Location = new Point(this.label5.Location.X + 80, this.label5.Location.Y);
                this.label6.Location = new Point(this.label6.Location.X + 80, this.label6.Location.Y);
                this.label7.Location = new Point(this.label7.Location.X + 80, this.label7.Location.Y);
                this.label8.Location = new Point(this.label8.Location.X + 80, this.label8.Location.Y);
                this.textBox1.Location = new Point(this.textBox1.Location.X + 80, this.textBox1.Location.Y);
                this.textBox2.Location = new Point(this.textBox2.Location.X + 80, this.textBox2.Location.Y);
                this.textBox3.Location = new Point(this.textBox3.Location.X + 80, this.textBox3.Location.Y);
                this.textBox4.Location = new Point(this.textBox4.Location.X + 80, this.textBox4.Location.Y);
                this.textBox5.Location = new Point(this.textBox5.Location.X + 80, this.textBox5.Location.Y);
                this.textBox6.Location = new Point(this.textBox6.Location.X + 80, this.textBox6.Location.Y);
            }
            
            baseSalaryNumShow();
        }
        
        /// <summary>
        /// 奖励基数框刷新
        /// </summary>
        private void baseSalaryNumShow() 
        {
            this.textBox1.Text = frmMain.PrimaryBaseNum.ToString();
            this.textBox2.Text = frmMain.PrimaryReward.ToString();
            this.textBox3.Text = frmMain.MiddleBaseNum.ToString();
            this.textBox4.Text = frmMain.MiddleReward.ToString();
            this.textBox5.Text = frmMain.SeniorBaseNum.ToString();
            this.textBox6.Text = frmMain.SeniorReward.ToString();
        }
        private void button1_Click(object sender, EventArgs e)//提交基数
        {
            this.errorProvider1.Clear();
            if (this.textBox1.Text.Trim() == "" || this.textBox2.Text.Trim() == "" || this.textBox3.Text.Trim() == "" ||
                this.textBox4.Text.Trim() == "" || this.textBox5.Text.Trim() == "" || this.textBox6.Text.Trim() == "")
            {
                MessageBox.Show("你想干嘛？？？");
                return;
            }
            try
            {
                int.Parse(this.textBox1.Text.Trim());
            }
            catch
            {
                this.errorProvider1.SetError(this.textBox1, "只能填写数字");
                return;
            }
            try
            {
                double.Parse(this.textBox2.Text.Trim());
            }
            catch
            {
                this.errorProvider1.SetError(this.textBox2, "只能填写数字");
                return;
            }
            try
            {
                int.Parse(this.textBox3.Text.Trim());
            }
            catch
            {
                this.errorProvider1.SetError(this.textBox3, "只能填写数字");
                return;
            }
            try
            {
                double.Parse(this.textBox4.Text.Trim());
            }
            catch
            {
                this.errorProvider1.SetError(this.textBox4, "只能填写数字");
                return;
            }
            try
            {
                int.Parse(this.textBox5.Text.Trim());
            }
            catch
            {
                this.errorProvider1.SetError(this.textBox5, "只能填写数字");
                return;
            }
            try
            {
                double.Parse(this.textBox6.Text.Trim());
            }
            catch
            {
                this.errorProvider1.SetError(this.textBox6, "只能填写数字");
                return;
            }
            if (MessageBox.Show("你确定要修改绩效基数吗?", "修改提醒", MessageBoxButtons.YesNo) == DialogResult.No)
                return;
            string sql1 = string.Format("update BaseDate set PrimaryBaseNum = {0},PrimaryReward = {1},MiddleBaseNum = {2},MiddleReward = {3},SeniorBaseNum = {4},SeniorReward = {5}",
                this.textBox1.Text.Trim(),
                this.textBox2.Text.Trim(),
                this.textBox3.Text.Trim(),
                this.textBox4.Text.Trim(),
                this.textBox5.Text.Trim(),
                this.textBox6.Text.Trim());
            int resultNum = DBHelper.ExecuteUpdate(sql1);
            if (resultNum > 0)
            {
                MessageBox.Show("修改成功");
                DataTable dtBaseNum = DBHelper.ExecuteQuery("select * from BaseDate");
                frmMain.PrimaryBaseNum = int.Parse(dtBaseNum.Rows[0][0].ToString());
                this.textBox1.Text = dtBaseNum.Rows[0][0].ToString();
                frmMain.PrimaryReward = double.Parse(dtBaseNum.Rows[0][1].ToString());
                this.textBox2.Text = dtBaseNum.Rows[0][1].ToString();
                frmMain.MiddleBaseNum = int.Parse(dtBaseNum.Rows[0][2].ToString());
                this.textBox3.Text = dtBaseNum.Rows[0][2].ToString();
                frmMain.MiddleReward = double.Parse(dtBaseNum.Rows[0][3].ToString());
                this.textBox4.Text = dtBaseNum.Rows[0][3].ToString();
                frmMain.SeniorBaseNum = int.Parse(dtBaseNum.Rows[0][4].ToString());
                this.textBox5.Text = dtBaseNum.Rows[0][4].ToString();
                frmMain.SeniorReward = double.Parse(dtBaseNum.Rows[0][5].ToString());
                this.textBox6.Text = dtBaseNum.Rows[0][5].ToString();
            }
        }

        private void button2_Click(object sender, EventArgs e)//查询
        {
            string sql1 = "";
            string name_temp = "";
            if (Login.userRoles == "IT传媒部-组长")
            {
                name_temp = this.comboBox1.Text.Trim();
            }
            else
            {
                name_temp = Login.userName;
            }
            sql1 = string.Format("select UpTimes as 日期,datename(weekday, UpTimes) as 星期,AssumeName as 花名,coalesce(sum(ActualUpNum),0) as 有效上传数, " +
                        "(case when coalesce(sum(ActualUpNum),0)<={0} then 0 " +
                        "when coalesce(sum(ActualUpNum),0)>{0} and coalesce(sum(ActualUpNum),0)<={2} then (coalesce(sum(ActualUpNum),0)-{0})*{1} " +
                        "when coalesce(sum(ActualUpNum),0)>{2} then {1}*({2}-{0}) " +
                        "else 0 end) as 初级激励, " +
                        "(case when coalesce(sum(ActualUpNum),0)<={2} then 0 " +
                        "when coalesce(sum(ActualUpNum),0)>{2} and coalesce(sum(ActualUpNum),0)<={4} then (coalesce(sum(ActualUpNum),0)-{2})*{3} " +
                        "when coalesce(sum(ActualUpNum),0)>{4} then {3}*({4}-{2}) " +
                        "else 0 end) as 中级激励, " +
                        "(case when coalesce(sum(ActualUpNum),0)<={4} then 0 " +
                        "when coalesce(sum(ActualUpNum),0)>{4} then {5}*(coalesce(sum(ActualUpNum),0)-{4}) " +
                        "else 0 end) as 高级激励, " +
                        "((case when coalesce(sum(ActualUpNum),0)<={0} then 0 " +
                        "when coalesce(sum(ActualUpNum),0)>{0} and coalesce(sum(ActualUpNum),0)<={2} then (coalesce(sum(ActualUpNum),0)-{0})*{1} " +
                        "when coalesce(sum(ActualUpNum),0)>{2} then {1}*({2}-{0}) " +
                        "else 0 end)+ " +
                        "(case when coalesce(sum(ActualUpNum),0)<={2} then 0 " +
                        "when coalesce(sum(ActualUpNum),0)>{2} and coalesce(sum(ActualUpNum),0)<={4} then (coalesce(sum(ActualUpNum),0)-{2})*{3} " +
                        "when coalesce(sum(ActualUpNum),0)>{4} then {3}*({4}-{2}) " +
                        "else 0 end)+ " +
                        "(case when coalesce(sum(ActualUpNum),0)<={4} then 0 " +
                        "when coalesce(sum(ActualUpNum),0)>{4} then {5}*(coalesce(sum(ActualUpNum),0)-{4}) " +
                        "else 0 end)) as 当日总绩效 " +
                        "from UpPerformance  " +
                        "where AssumeName like '%{6}%' and UpTimes between '{7}' and '{8}' " +
                        "group by UpTimes, AssumeName " +
                        "order by UpTimes desc",
                        frmMain.PrimaryBaseNum, frmMain.PrimaryReward, frmMain.MiddleBaseNum, frmMain.MiddleReward, frmMain.SeniorBaseNum, frmMain.SeniorReward,
                        name_temp, this.dateTimePicker1.Value.ToShortDateString(), this.dateTimePicker2.Value.ToShortDateString());
            DataTable dt1 = DBHelper.ExecuteQuery(sql1);
            this.dataGridView1.DataSource = dt1;
            //此处将什么低于550的单元格标红，星期天标绿
            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                if (int.Parse(dt1.Rows[i][3].ToString()) < 550)
                {
                    this.dataGridView1.Rows[i].Cells[3].Style.ForeColor = Color.Red;
                }
                if (dt1.Rows[i][1].ToString() == "星期日")
                {
                    this.dataGridView1.Rows[i].Cells[1].Style.ForeColor = Color.FromArgb(1,144,26);
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("你确定要导出数据至excel表格吗?", "导出提醒", MessageBoxButtons.YesNo) == DialogResult.No)
                return;
            string sql1 = string.Format("select CONVERT(varchar(100), UpTimes, 23) as 日期,datename(weekday, UpTimes) as 星期,AssumeName as 花名,coalesce(sum(ActualUpNum),0) as 有效上传数, " +
                        "(case when coalesce(sum(ActualUpNum),0)<={0} then 0 " +
                        "when coalesce(sum(ActualUpNum),0)>{0} and coalesce(sum(ActualUpNum),0)<={2} then (coalesce(sum(ActualUpNum),0)-{0})*{1} " +
                        "when coalesce(sum(ActualUpNum),0)>{2} then {1}*({2}-{0}) " +
                        "else 0 end) as 初级激励, " +
                        "(case when coalesce(sum(ActualUpNum),0)<={2} then 0 " +
                        "when coalesce(sum(ActualUpNum),0)>{2} and coalesce(sum(ActualUpNum),0)<={4} then (coalesce(sum(ActualUpNum),0)-{2})*{3} " +
                        "when coalesce(sum(ActualUpNum),0)>{4} then {3}*({4}-{2}) " +
                        "else 0 end) as 中级激励, " +
                        "(case when coalesce(sum(ActualUpNum),0)<={4} then 0 " +
                        "when coalesce(sum(ActualUpNum),0)>{4} then {5}*(coalesce(sum(ActualUpNum),0)-{4}) " +
                        "else 0 end) as 高级激励, " +
                        "((case when coalesce(sum(ActualUpNum),0)<={0} then 0 " +
                        "when coalesce(sum(ActualUpNum),0)>{0} and coalesce(sum(ActualUpNum),0)<={2} then (coalesce(sum(ActualUpNum),0)-{0})*{1} " +
                        "when coalesce(sum(ActualUpNum),0)>{2} then {1}*({2}-{0}) " +
                        "else 0 end)+ " +
                        "(case when coalesce(sum(ActualUpNum),0)<={2} then 0 " +
                        "when coalesce(sum(ActualUpNum),0)>{2} and coalesce(sum(ActualUpNum),0)<={4} then (coalesce(sum(ActualUpNum),0)-{2})*{3} " +
                        "when coalesce(sum(ActualUpNum),0)>{4} then {3}*({4}-{2}) " +
                        "else 0 end)+ " +
                        "(case when coalesce(sum(ActualUpNum),0)<={4} then 0 " +
                        "when coalesce(sum(ActualUpNum),0)>{4} then {5}*(coalesce(sum(ActualUpNum),0)-{4}) " +
                        "else 0 end)) as 当日总绩效 " +
                        "from UpPerformance  " +
                        "where AssumeName like '%{6}%' and UpTimes between '{7}' and '{8}' " +
                        "group by UpTimes, AssumeName " +
                        "order by UpTimes desc",
                        frmMain.PrimaryBaseNum, frmMain.PrimaryReward, frmMain.MiddleBaseNum, frmMain.MiddleReward, frmMain.SeniorBaseNum, frmMain.SeniorReward,
                        this.comboBox1.Text.Trim(), this.dateTimePicker1.Value.ToShortDateString(), this.dateTimePicker2.Value.ToShortDateString());
            DataTable dt1 = DBHelper.ExecuteQuery(sql1);
            if (dt1.Rows.Count > 0)
            {
                if (Directory.Exists(@"D:\\日数据表(软件导出)") == false)//如果不存在就创建file文件夹
                {
                    Directory.CreateDirectory(@"D:\\日数据表(软件导出)");
                }
                try
                {
                    Tools.dataTableToCsv(dt1, @"D:\\日数据表(软件导出)\" +
                    this.dateTimePicker1.Value.ToString("yyyyMMdd") + "至" + this.dateTimePicker2.Value.ToString("yyyyMMdd") +
                    this.comboBox1.Text + "日数据表" + ".xls");
                    MessageBox.Show("导出完成,路径为：D:\\\\日数据表(软件导出)\\" + 
                    this.dateTimePicker1.Value.ToString("yyyyMMdd") + "至" + this.dateTimePicker2.Value.ToString("yyyyMMdd") +
                    this.comboBox1.Text + "日数据表" + ".xls");
                    return;
                }
                catch
                {
                    MessageBox.Show("请先关闭对应导出日期的excel表程序");
                    return;
                }
            }
        }

        
    }
}
