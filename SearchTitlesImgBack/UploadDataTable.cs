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
    public partial class UploadDataTable : Form
    {
        public UploadDataTable()
        {
            InitializeComponent();
        }
        private void UploadDataTable_Load(object sender, EventArgs e)
        {
            this.dateTimePicker1.Value = System.DateTime.Now.AddDays(-7);
            this.dateTimePicker2.Value = System.DateTime.Now;
            this.dateTimePicker3.Value = System.DateTime.Now;
            if (Login.userRoles == "IT传媒部-组长")
            {
                DataTable dtUserSmall = frmMain.dtUser.Copy();
                dtUserSmall.Rows.Add("");
                this.comboBox1.DataSource = dtUserSmall.Copy();
                this.comboBox1.DisplayMember = "UserId";
                this.comboBox1.ValueMember = "UserId";

                this.comboBox2.DataSource = frmMain.dtUser.Copy();
                this.comboBox2.DisplayMember = "UserId";
                this.comboBox2.ValueMember = "UserId";
            }
            else
            {
                this.linkLabel1.Visible = false;
                this.dateTimePicker3.Enabled = false;

                this.comboBox1.Text = Login.userName;
                this.comboBox1.Enabled = false;

                this.comboBox2.Text = Login.userName;
                this.comboBox2.Enabled = false;
            }
        }
        private void button1_Click(object sender, EventArgs e)//查询
        {
            string sql1 = "";
            if (Login.userRoles == "IT传媒部-组员")
            {
                sql1 = string.Format("select UpTimes,datename(weekday, UpTimes) as Week,AssumeName,UpId,UpNum,InvalidNum,InvalidReason,ActualUpNum from UpPerformance "+
                    "where UpTimes between '{0}' and '{1}' and AssumeName = '{2}' order by UpTimes desc",
                    this.dateTimePicker1.Value.ToShortDateString(),
                    this.dateTimePicker2.Value.ToShortDateString(),
                    Login.userName);
            }
            else
            {
                sql1 = string.Format("select UpTimes,datename(weekday, UpTimes) as Week,AssumeName,UpId,UpNum,InvalidNum,InvalidReason,ActualUpNum from UpPerformance "+
                    "where UpTimes between '{0}' and '{1}' and AssumeName like '%{2}%' order by UpTimes desc",
                    this.dateTimePicker1.Value.ToShortDateString(),
                    this.dateTimePicker2.Value.ToShortDateString(),
                    this.comboBox1.Text);
            }
            DataTable dt1 = DBHelper.ExecuteQuery(sql1);
            this.dataGridView1.DataSource = dt1;
        }

        private void button2_Click(object sender, EventArgs e)//修改
        {
            this.errorProvider1.Clear();
            if (Login.userRoles == "IT传媒部-组长")
            {
                if (this.dataGridView1.SelectedRows.Count == 0)
                {
                    MessageBox.Show("你想干嘛？？？");
                    return;
                }
                try
                {
                    int.Parse(this.textBox2.Text.Trim());
                }
                catch
                {
                    this.errorProvider1.SetError(this.textBox2,"只能填写数字");
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
                string sql1 = string.Format("delete from UpPerformance where UpTimes = '{0}' and AssumeName = '{1}' and UpId = '{2}'",
                    this.dataGridView1.SelectedRows[0].Cells[0].Value.ToString(),
                    this.dataGridView1.SelectedRows[0].Cells[2].Value.ToString(),
                    this.dataGridView1.SelectedRows[0].Cells[3].Value.ToString());
                int resultNum1 = DBHelper.ExecuteUpdate(sql1);
                string sql2 = string.Format("insert into UpPerformance values('{0}','{1}','{2}',{3},{4},'{5}',{3}-{4})",
                    this.dateTimePicker3.Value.ToShortDateString(),
                    this.comboBox2.Text,
                    this.textBox1.Text.Trim(),
                    this.textBox2.Text.Trim(),
                    this.textBox3.Text.Trim(),
                    this.textBox4.Text);
                int resultNum2 = DBHelper.ExecuteUpdate(sql2);
                if (resultNum1>0&&resultNum2 > 0)
                {
                    this.dataGridView1.SelectedRows[0].Cells[0].Value = this.dateTimePicker3.Value.ToShortDateString();
                    this.dataGridView1.SelectedRows[0].Cells[1].Value = Tools.getWeek(DateTime.Parse(this.dateTimePicker3.Value.ToShortDateString()));
                    this.dataGridView1.SelectedRows[0].Cells[2].Value = this.comboBox2.Text.Trim();
                    this.dataGridView1.SelectedRows[0].Cells[3].Value = this.textBox1.Text.Trim();
                    this.dataGridView1.SelectedRows[0].Cells[4].Value = this.textBox2.Text.Trim();
                    this.dataGridView1.SelectedRows[0].Cells[5].Value = this.textBox3.Text.Trim();
                    this.dataGridView1.SelectedRows[0].Cells[6].Value = this.textBox4.Text.Trim();
                    this.dataGridView1.SelectedRows[0].Cells[7].Value = (int.Parse(this.textBox2.Text.Trim()) - int.Parse(this.textBox3.Text.Trim())).ToString();
                    MessageBox.Show("修改成功");
                }
                else 
                {
                    MessageBox.Show("修改失败");
                }
            }
            else
            {
                if (this.textBox2.Text.Trim() == "")
                {
                    MessageBox.Show("你想干嘛？？？");
                    return;
                }
                try
                {
                    int.Parse(this.textBox2.Text.Trim());
                }
                catch
                {
                    this.errorProvider1.SetError(this.textBox2, "只能填写数字");
                    return;
                }
                //删除今天的id记录
                string sql1 = string.Format("delete from UpPerformance where UpTimes = CONVERT(varchar(12) , getdate(), 111 ) and AssumeName = '{0}' and UpId = '{1}'",
                     Login.userName, this.textBox1.Text.Trim());
                DBHelper.ExecuteUpdate(sql1);
                //增加今天的ID记录
                string sql2 = string.Format("insert into UpPerformance values(CONVERT(varchar(12) , getdate(), 111 ),'{0}','{1}',{2},0,'',{2})",
                    Login.userName, this.textBox1.Text.Trim(), this.textBox2.Text.Trim());
                int resultNum = DBHelper.ExecuteUpdate(sql2);
                if (resultNum > 0)
                {
                    MessageBox.Show("提交成功");
                    this.button1_Click(null, null);
                }
                else 
                {
                    MessageBox.Show("提交失败");
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)//删除
        {
            DataTable dt1 = DBHelper.ExecuteQuery("select getdate()");
            //服务器时间
            DateTime nowtime = DateTime.Parse(((DateTime.Parse(dt1.Rows[0][0].ToString())).ToShortDateString()));
            if (this.dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("请先选定记录");
                return;
            }
            if (Login.userRoles == "IT传媒部-组员")
            {
                if (nowtime > DateTime.Parse(this.dataGridView1.SelectedRows[0].Cells[0].Value.ToString()))
                {
                    MessageBox.Show("请联系组长删除");
                    return;
                }
            }
            if (MessageBox.Show("你确定要删除【" + this.dataGridView1.SelectedRows[0].Cells[3].Value.ToString() + "】吗?", "删除提醒", MessageBoxButtons.YesNo) == DialogResult.No)
                return;
            string sql1 = string.Format("delete from UpPerformance where UpTimes = '{0}' and AssumeName = '{1}' and UpId = '{2}'",
                this.dataGridView1.SelectedRows[0].Cells[0].Value.ToString(),
                this.dataGridView1.SelectedRows[0].Cells[2].Value.ToString(),
                this.dataGridView1.SelectedRows[0].Cells[3].Value.ToString());
            int resultNum = DBHelper.ExecuteUpdate(sql1);
            if (resultNum > 0)
            {
                MessageBox.Show("删除成功");
                this.dataGridView1.Rows.Remove(this.dataGridView1.SelectedRows[0]);
            }
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (this.dataGridView1.SelectedRows.Count == 0)
                return;
            this.textBox1.Text = this.dataGridView1.SelectedRows[0].Cells[3].Value.ToString();
            this.textBox2.Text = this.dataGridView1.SelectedRows[0].Cells[4].Value.ToString();
            this.textBox3.Text = this.dataGridView1.SelectedRows[0].Cells[5].Value.ToString();
            this.textBox4.Text = this.dataGridView1.SelectedRows[0].Cells[6].Value.ToString();
            if (Login.userRoles != "IT传媒部-组长")
            {
                if (DateTime.Parse(this.dataGridView1.SelectedRows[0].Cells[0].Value.ToString()) < DateTime.Parse(System.DateTime.Now.ToShortDateString()))
                { this.button2.Enabled = false; this.button3.Enabled = false; }
                else
                { this.button2.Enabled = true; this.button3.Enabled = true; }
                return;
            }
            this.dateTimePicker3.Value = DateTime.Parse(this.dataGridView1.SelectedRows[0].Cells[0].Value.ToString());
            this.comboBox2.Text = this.dataGridView1.SelectedRows[0].Cells[2].Value.ToString();
            this.textBox3.Text = this.dataGridView1.SelectedRows[0].Cells[5].Value.ToString();
            this.textBox4.Text = this.dataGridView1.SelectedRows[0].Cells[6].Value.ToString();
        }

        private void button4_Click(object sender, EventArgs e)//增加
        {
            this.errorProvider1.Clear();
            if (Login.userRoles == "IT传媒部-组长")
            {
                try
                {
                    int.Parse(this.textBox2.Text.Trim());
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
                string sql1 = string.Format("insert into UpPerformance values('{0}','{1}','{2}',{3},{4},'{5}',{3}-{4})",
                    this.dateTimePicker3.Value.ToShortDateString(),
                    this.comboBox2.Text,
                    this.textBox1.Text.Trim(),
                    this.textBox2.Text.Trim(),
                    this.textBox3.Text.Trim(),
                    this.textBox4.Text);
                int resultNum1 = DBHelper.ExecuteUpdate(sql1);
                if (resultNum1 > 0)
                {
                    this.button1_Click(null, null);
                    MessageBox.Show("添加成功");
                }
                else
                {
                    MessageBox.Show("添加失败");
                }
            }
            else 
            {
                if (this.textBox2.Text.Trim() == "")
                {
                    MessageBox.Show("你想干嘛？？？");
                    return;
                }
                try
                {
                    int.Parse(this.textBox2.Text.Trim());
                }
                catch
                {
                    this.errorProvider1.SetError(this.textBox2, "只能填写数字");
                    return;
                }
                //增加今天的ID记录
                string sql1 = string.Format("insert into UpPerformance values(CONVERT(varchar(12) , getdate(), 111 ),'{0}','{1}',{2},0,'',{2})",
                    Login.userName, this.textBox1.Text.Trim(), this.textBox2.Text.Trim());
                int resultNum1 = DBHelper.ExecuteUpdate(sql1);
                if (resultNum1 > 0)
                {
                    MessageBox.Show("添加成功");
                    this.button1_Click(null, null);
                }
                else
                {
                    MessageBox.Show("添加失败");
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("你确定要导出数据至excel表格吗?", "导出提醒", MessageBoxButtons.YesNo) == DialogResult.No)
                return;
            string sql1 = string.Format("select CONVERT(varchar(100), UpTimes, 23) as 日期,datename(weekday, UpTimes) as 星期,AssumeName as 花名,UpId as 上传ID,UpNum as 上传数,InvalidNum as 无效数,InvalidReason as 原因,ActualUpNum as 实际数 from UpPerformance " +
                    "where UpTimes between '{0}' and '{1}' and AssumeName like '%{2}%' order by UpTimes desc",
                    this.dateTimePicker1.Value.ToShortDateString(),
                    this.dateTimePicker2.Value.ToShortDateString(),
                    this.comboBox1.Text);
            DataTable dt1 = DBHelper.ExecuteQuery(sql1);
            if (dt1.Rows.Count > 0)
            {
                if (Directory.Exists(@"D:\\上传明细表(软件导出)") == false)//如果不存在就创建file文件夹
                {
                    Directory.CreateDirectory(@"D:\\上传明细表(软件导出)");
                }
                try
                {
                    Tools.dataTableToCsv(dt1, @"D:\\上传明细表(软件导出)\" +
                    this.dateTimePicker1.Value.ToString("yyyyMMdd") + "至" + this.dateTimePicker2.Value.ToString("yyyyMMdd") +
                    this.comboBox1.Text + "上传明细表" + ".xls");
                    MessageBox.Show("导出完成,路径为：D:\\\\上传明细表(软件导出)\\" + this.dateTimePicker1.Value.ToString("yyyyMMdd") + "至" +
                    this.dateTimePicker2.Value.ToString("yyyyMMdd") + "上传明细表" + ".xls");
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
