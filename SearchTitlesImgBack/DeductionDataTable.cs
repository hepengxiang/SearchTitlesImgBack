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
    public partial class DeductionDataTable : Form
    {
        public DeductionDataTable()
        {
            InitializeComponent();
        }

        private void DeductionDataTable_Load(object sender, EventArgs e)
        {
            this.comboBox2.SelectedIndex = System.DateTime.Now.Month - 1;
            if (Login.userRoles == "IT传媒部-组长")
            {
                DataTable dtUserSmall = frmMain.dtUser.Copy();
                dtUserSmall.Rows.Add("");
                this.comboBox3.DataSource = dtUserSmall;
                this.comboBox3.DisplayMember = "UserId";
                this.comboBox3.ValueMember = "UserId";
            }
            else
            {
                this.linkLabel1.Visible = false;
                this.comboBox3.Text = Login.userName;
                this.comboBox3.Enabled = false;
                this.textBox1.Enabled = false;
                this.textBox2.Enabled = false;
                this.button2.Visible = false;
            }
        }

        private void button1_Click(object sender, EventArgs e)//查询
        {
            List<DateTime> datetimeList = getSundaysOfMonth(int.Parse(this.comboBox1.Text.Trim()), int.Parse(this.comboBox2.Text.Trim()));
            DateTime startTime = datetimeList[0];
            DateTime endTime = datetimeList[datetimeList.Count - 2];
            string sql1 = string.Format("declare @times datetime, @sql varchar(max), @name varchar(20), @i int, " +
                        "@PrimaryBaseNum int, @PrimaryReward float, @MiddleBaseNum int, @MiddleReward float, @SeniorBaseNum int, @SeniorReward float " +
                        "set @PrimaryBaseNum = {0} " +
                        "set @PrimaryReward = {1} " +
                        "set @MiddleBaseNum = {2} " +
                        "set @MiddleReward = {3} " +
                        "set @SeniorBaseNum = {4} " +
                        "set @SeniorReward = {5} " +
                        "set @sql = ''  set @name='{8}' set @i = 1 " +
                        "set @times = '{6}' " +
                          "while @times <='{7}' " +
                          "begin   " +
                            "set @sql = @sql+'select '+ convert(varchar,@i)+' as 周期,AssumeName as 花名,coalesce(sum(ActualUpNum),0) as 上传量, " +
                            "cast(coalesce(sum( " +
                                "case " +
                                "when ActualUpNum <='+ convert(varchar,@PrimaryBaseNum)+' then 0 " +
                                "when ActualUpNum >'+ convert(varchar,@PrimaryBaseNum)+' and ActualUpNum <='+ convert(varchar,@MiddleBaseNum)+' then (ActualUpNum-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@MiddleBaseNum)+' and ActualUpNum <='+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@SeniorBaseNum)+')*'+ convert(varchar,@SeniorReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+'+('+ convert(varchar,@SeniorBaseNum)+'-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+' " +
                                "else 0 end " +
                                "),0)as   decimal(10,   0)) as 绩效 , " +
                            "coalesce((select DeductionNum from DeductionPerformance where SubmitTime between '''+CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and '''+CONVERT(varchar(100), @times, 23)+'''  " +
                            "and PerformancePerson = AssumeName),0) as 扣除, " +
                            "coalesce((select DeductionReason from DeductionPerformance where SubmitTime between '''+CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and  '''+CONVERT(varchar(100), @times, 23)+'''  " +
                            "and PerformancePerson = AssumeName),''无'') as 原因, " +
                            "cast((cast(coalesce(sum( " +
                                "case  " +
                                "when ActualUpNum <='+ convert(varchar,@PrimaryBaseNum)+' then 0 " +
                                "when ActualUpNum >'+ convert(varchar,@PrimaryBaseNum)+' and ActualUpNum <='+ convert(varchar,@MiddleBaseNum)+' then (ActualUpNum-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@MiddleBaseNum)+' and ActualUpNum <='+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@SeniorBaseNum)+')*'+ convert(varchar,@SeniorReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+'+('+ convert(varchar,@SeniorBaseNum)+'-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+' " +
                                "else 0 end " +
                                "),0)as   decimal(10,   0))- " +
                            "coalesce((select DeductionNum from DeductionPerformance where SubmitTime between '''+CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and '''+CONVERT(varchar(100), @times, 23)+'''  " +
                            "and PerformancePerson = AssumeName),0))as   decimal(10,   0)) as 实际绩效, " +
                            "'''+CONVERT(varchar(100), @times, 23)+'''  as 日期 " +
                            " from UpPerformance " +
                            "where UpTimes between '''++CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and '''+CONVERT(varchar(100), @times, 23)+''' and AssumeName like ''%'+@name+'%'' " +
                            "group by AssumeName union all ' " +
                            "set @times = DATEADD(day,7,@times) " +
                            "set @i = @i+1 " +
                            "print @sql " +
                          "end " +
                            "set @sql = @sql+'select '+ convert(varchar,@i)+' as 周期,AssumeName as 花名,coalesce(sum(ActualUpNum),0) as 上传量, " +
                                "cast(coalesce(sum( " +
                                "case  " +
                                "when ActualUpNum <='+ convert(varchar,@PrimaryBaseNum)+' then 0 " +
                                "when ActualUpNum >'+ convert(varchar,@PrimaryBaseNum)+' and ActualUpNum <='+ convert(varchar,@MiddleBaseNum)+' then (ActualUpNum-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@MiddleBaseNum)+' and ActualUpNum <='+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@SeniorBaseNum)+')*'+ convert(varchar,@SeniorReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+'+('+ convert(varchar,@SeniorBaseNum)+'-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+' " +
                                "else 0 end " +
                                "),0)as   decimal(10,   0)) as 绩效 , " +
                            "coalesce((select DeductionNum from DeductionPerformance where SubmitTime between '''+CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and '''+CONVERT(varchar(100), @times, 23)+'''  " +
                            "and PerformancePerson = AssumeName),0) as 扣除, " +
                            "coalesce((select DeductionReason from DeductionPerformance where SubmitTime between '''+CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and  '''+CONVERT(varchar(100), @times, 23)+'''  " +
                            "and PerformancePerson = AssumeName),''无'') as 原因, " +
                            "cast((cast(coalesce(sum( " +
                                "case  " +
                                "when ActualUpNum <='+ convert(varchar,@PrimaryBaseNum)+' then 0 " +
                                "when ActualUpNum >'+ convert(varchar,@PrimaryBaseNum)+' and ActualUpNum <='+ convert(varchar,@MiddleBaseNum)+' then (ActualUpNum-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@MiddleBaseNum)+' and ActualUpNum <='+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@SeniorBaseNum)+')*'+ convert(varchar,@SeniorReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+'+('+ convert(varchar,@SeniorBaseNum)+'-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+' " +
                                "else 0 end " +
                                "),0)as   decimal(10,   0))- " +
                            "coalesce((select DeductionNum from DeductionPerformance where SubmitTime between '''+CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and '''+CONVERT(varchar(100), @times, 23)+'''  " +
                            "and PerformancePerson = AssumeName),0))as   decimal(10,   0)) as 实际绩效, " +
                            "'''+CONVERT(varchar(100), @times, 23)+''' as 日期 " +
                            " from UpPerformance " +
                            "where UpTimes between '''++CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and '''+CONVERT(varchar(100), @times, 23)+''' and AssumeName like ''%'+@name+'%'' " +
                            "group by AssumeName' " +
                        "exec(@sql)", 
                        frmMain.PrimaryBaseNum, frmMain.PrimaryReward, frmMain.MiddleBaseNum, 
                        frmMain.MiddleReward, frmMain.SeniorBaseNum, frmMain.SeniorReward, startTime, endTime, this.comboBox3.Text.Trim());
            DataTable dt1 = DBHelper.ExecuteQuery(sql1);
            this.dataGridView1.DataSource = dt1;
            //此处将什么低于3300的单元格标红
            for (int i = 0; i < dt1.Rows.Count; i++)
            {
                if (int.Parse(dt1.Rows[i][2].ToString()) < 3300)
                {
                    this.dataGridView1.Rows[i].Cells[2].Style.ForeColor = Color.Red;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)//提交
        {
            this.errorProvider1.Clear();
            if (this.dataGridView1.SelectedRows.Count == 0)
            {
                MessageBox.Show("你想干嘛？？？");
                return;
            }
            if (this.textBox1.Text.Trim() == "")
            {
                this.errorProvider1.SetError(this.textBox1,"必须填写内容");
                return;
            }
            if (this.textBox2.Text.Trim() == "")
            {
                this.errorProvider1.SetError(this.textBox2, "必须填写内容");
                return;
            }
            try
            {
                int.Parse(this.textBox1.Text.Trim());
            }
            catch
            {
                this.errorProvider1.SetError(this.textBox1, "必须填写数字");
                return;
            }
            string sql1 = string.Format("delete from DeductionPerformance where SubmitTime = '{0}' and PerformancePerson = '{1}'",
                this.dataGridView1.SelectedRows[0].Cells[7].Value.ToString(),
                this.dataGridView1.SelectedRows[0].Cells[1].Value.ToString());
            DBHelper.ExecuteUpdate(sql1);
            string sql2 = string.Format("insert into DeductionPerformance values('{0}',{1},'{2}','{3}')",
                this.dataGridView1.SelectedRows[0].Cells[7].Value.ToString(),
                this.textBox1.Text.Trim().Replace("-", ""),
                this.textBox2.Text.Trim(),
                this.dataGridView1.SelectedRows[0].Cells[1].Value.ToString());
            int resultNum2 = DBHelper.ExecuteUpdate(sql2);
            if (resultNum2 > 0)
            {
                MessageBox.Show("提交成功");
                this.dataGridView1.SelectedRows[0].Cells[4].Value = this.textBox1.Text.Trim().Replace("-", "");
                this.dataGridView1.SelectedRows[0].Cells[5].Value = this.textBox2.Text.Trim();
                this.dataGridView1.SelectedRows[0].Cells[6].Value = (int)(
                    double.Parse(this.dataGridView1.SelectedRows[0].Cells[3].Value.ToString()) -
                    double.Parse(this.dataGridView1.SelectedRows[0].Cells[4].Value.ToString()));
            }
        }
        /// <summary>
        /// 获取指定月份的周末的集合
        /// </summary>
        /// <param name="yearNum">年份</param>
        /// <param name="monthNum">月份</param>
        /// <returns></returns>
        private List<DateTime> getSundaysOfMonth(int yearNum, int monthNum)
        {
            List<DateTime> datetimeList = new List<DateTime>();
            DateTime startTime = DateTime.Parse(yearNum.ToString() + "-" + monthNum.ToString() + "-1");//这个月的第一天
            DateTime endTime = startTime.AddDays(1 - startTime.Day).AddMonths(1).AddDays(-1);//这个月最后一天
            TimeSpan ts = endTime.Subtract(startTime);//TimeSpan得到fromTime和toTime的时间间隔  
            int countday = ts.Days;//获取两个日期间的总天数  
            //循环用来扣除总天数中的双休日  
            for (int i = 0; i < countday; i++)
            {
                DateTime tempdt = startTime.Date.AddDays(i + 1);
                if (tempdt.DayOfWeek == System.DayOfWeek.Sunday)
                {
                    datetimeList.Add(tempdt);
                }
            }
            return datetimeList;
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            this.textBox1.Text = this.dataGridView1.SelectedRows[0].Cells[4].Value.ToString();
            this.textBox2.Text = this.dataGridView1.SelectedRows[0].Cells[5].Value.ToString();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("你确定要导出数据至excel表格吗?", "导出提醒", MessageBoxButtons.YesNo) == DialogResult.No)
                return;
            List<DateTime> datetimeList = getSundaysOfMonth(int.Parse(this.comboBox1.Text.Trim()), int.Parse(this.comboBox2.Text.Trim()));
            DateTime startTime = datetimeList[0];
            DateTime endTime = datetimeList[datetimeList.Count - 2];
            string sql1 = string.Format("declare @times datetime, @sql varchar(max), @name varchar(20), @i int, " +
                        "@PrimaryBaseNum int, @PrimaryReward float, @MiddleBaseNum int, @MiddleReward float, @SeniorBaseNum int, @SeniorReward float " +
                        "set @PrimaryBaseNum = {0} " +
                        "set @PrimaryReward = {1} " +
                        "set @MiddleBaseNum = {2} " +
                        "set @MiddleReward = {3} " +
                        "set @SeniorBaseNum = {4} " +
                        "set @SeniorReward = {5} " +
                        "set @sql = ''  set @name='{8}' set @i = 1 " +
                        "set @times = '{6}' " +
                          "while @times <='{7}' " +
                          "begin   " +
                            "set @sql = @sql+'select '+ convert(varchar,@i)+' as 周期,AssumeName as 花名,coalesce(sum(ActualUpNum),0) as 上传量, " +
                            "cast(coalesce(sum( " +
                                "case " +
                                "when ActualUpNum <='+ convert(varchar,@PrimaryBaseNum)+' then 0 " +
                                "when ActualUpNum >'+ convert(varchar,@PrimaryBaseNum)+' and ActualUpNum <='+ convert(varchar,@MiddleBaseNum)+' then (ActualUpNum-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@MiddleBaseNum)+' and ActualUpNum <='+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@SeniorBaseNum)+')*'+ convert(varchar,@SeniorReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+'+('+ convert(varchar,@SeniorBaseNum)+'-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+' " +
                                "else 0 end " +
                                "),0)as   decimal(10,   0)) as 绩效 , " +
                            "coalesce((select DeductionNum from DeductionPerformance where SubmitTime between '''+CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and '''+CONVERT(varchar(100), @times, 23)+'''  " +
                            "and PerformancePerson = AssumeName),0) as 扣除, " +
                            "coalesce((select DeductionReason from DeductionPerformance where SubmitTime between '''+CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and  '''+CONVERT(varchar(100), @times, 23)+'''  " +
                            "and PerformancePerson = AssumeName),''无'') as 原因, " +
                            "cast((cast(coalesce(sum( " +
                                "case  " +
                                "when ActualUpNum <='+ convert(varchar,@PrimaryBaseNum)+' then 0 " +
                                "when ActualUpNum >'+ convert(varchar,@PrimaryBaseNum)+' and ActualUpNum <='+ convert(varchar,@MiddleBaseNum)+' then (ActualUpNum-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@MiddleBaseNum)+' and ActualUpNum <='+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@SeniorBaseNum)+')*'+ convert(varchar,@SeniorReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+'+('+ convert(varchar,@SeniorBaseNum)+'-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+' " +
                                "else 0 end " +
                                "),0)as   decimal(10,   0))- " +
                            "coalesce((select DeductionNum from DeductionPerformance where SubmitTime between '''+CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and '''+CONVERT(varchar(100), @times, 23)+'''  " +
                            "and PerformancePerson = AssumeName),0))as   decimal(10,   0)) as 实际绩效, " +
                            "'''+CONVERT(varchar(100), @times, 23)+'''  as 日期 " +
                            " from UpPerformance " +
                            "where UpTimes between '''++CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and '''+CONVERT(varchar(100), @times, 23)+''' and AssumeName like ''%'+@name+'%'' " +
                            "group by AssumeName union all ' " +
                            "set @times = DATEADD(day,7,@times) " +
                            "set @i = @i+1 " +
                            "print @sql " +
                          "end " +
                            "set @sql = @sql+'select '+ convert(varchar,@i)+' as 周期,AssumeName as 花名,coalesce(sum(ActualUpNum),0) as 上传量, " +
                                "cast(coalesce(sum( " +
                                "case  " +
                                "when ActualUpNum <='+ convert(varchar,@PrimaryBaseNum)+' then 0 " +
                                "when ActualUpNum >'+ convert(varchar,@PrimaryBaseNum)+' and ActualUpNum <='+ convert(varchar,@MiddleBaseNum)+' then (ActualUpNum-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@MiddleBaseNum)+' and ActualUpNum <='+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@SeniorBaseNum)+')*'+ convert(varchar,@SeniorReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+'+('+ convert(varchar,@SeniorBaseNum)+'-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+' " +
                                "else 0 end " +
                                "),0)as   decimal(10,   0)) as 绩效 , " +
                            "coalesce((select DeductionNum from DeductionPerformance where SubmitTime between '''+CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and '''+CONVERT(varchar(100), @times, 23)+'''  " +
                            "and PerformancePerson = AssumeName),0) as 扣除, " +
                            "coalesce((select DeductionReason from DeductionPerformance where SubmitTime between '''+CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and  '''+CONVERT(varchar(100), @times, 23)+'''  " +
                            "and PerformancePerson = AssumeName),''无'') as 原因, " +
                            "cast((cast(coalesce(sum( " +
                                "case  " +
                                "when ActualUpNum <='+ convert(varchar,@PrimaryBaseNum)+' then 0 " +
                                "when ActualUpNum >'+ convert(varchar,@PrimaryBaseNum)+' and ActualUpNum <='+ convert(varchar,@MiddleBaseNum)+' then (ActualUpNum-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@MiddleBaseNum)+' and ActualUpNum <='+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+' " +
                                "when ActualUpNum >'+ convert(varchar,@SeniorBaseNum)+' then (ActualUpNum-'+ convert(varchar,@SeniorBaseNum)+')*'+ convert(varchar,@SeniorReward)+'+('+ convert(varchar,@MiddleBaseNum)+'-'+ convert(varchar,@PrimaryBaseNum)+')*'+ convert(varchar,@PrimaryReward)+'+('+ convert(varchar,@SeniorBaseNum)+'-'+ convert(varchar,@MiddleBaseNum)+')*'+ convert(varchar,@MiddleReward)+' " +
                                "else 0 end " +
                                "),0)as   decimal(10,   0))- " +
                            "coalesce((select DeductionNum from DeductionPerformance where SubmitTime between '''+CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and '''+CONVERT(varchar(100), @times, 23)+'''  " +
                            "and PerformancePerson = AssumeName),0))as   decimal(10,   0)) as 实际绩效, " +
                            "'''+CONVERT(varchar(100), @times, 23)+''' as 日期 " +
                            " from UpPerformance " +
                            "where UpTimes between '''++CONVERT(varchar(100), DATEADD(day,-6,@times), 23)+''' and '''+CONVERT(varchar(100), @times, 23)+''' and AssumeName like ''%'+@name+'%'' " +
                            "group by AssumeName' " +
                        "exec(@sql)",
                        frmMain.PrimaryBaseNum, frmMain.PrimaryReward, frmMain.MiddleBaseNum,
                        frmMain.MiddleReward, frmMain.SeniorBaseNum, frmMain.SeniorReward, startTime, endTime, this.comboBox3.Text.Trim());
            DataTable dt1 = DBHelper.ExecuteQuery(sql1);
            if (dt1.Rows.Count > 0)
            {
                if (Directory.Exists(@"D:\\月数据表(软件导出)") == false)//如果不存在就创建file文件夹
                {
                    Directory.CreateDirectory(@"D:\\月数据表(软件导出)");
                }
                try
                {
                    Tools.dataTableToCsv(dt1, @"D:\\月数据表(软件导出)\" +
                    this.comboBox3.Text+"--"+this.comboBox1.Text + "年" + this.comboBox2.Text + "月" +
                    "上传明细表" + ".xls");
                    MessageBox.Show("导出完成,路径为：D:\\\\月数据表(软件导出)\\" + this.comboBox3.Text + "--" + this.comboBox1.Text + "年" + 
                        this.comboBox2.Text + "月" + "上传明细表" + ".xls");
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
