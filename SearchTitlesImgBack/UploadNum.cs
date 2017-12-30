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
    public partial class UploadNum : Form
    {
        private static string DBTable = "";
        private static DataTable dt;
        public UploadNum()
        {
            InitializeComponent();
        }
        private void UploadNum_Load(object sender, EventArgs e)
        {
            this.dataGridView2.ColumnHeadersHeight = 30;

            this.dateTimePicker1.Value = System.DateTime.Now.AddDays(-7);
            this.dateTimePicker2.Value = System.DateTime.Now;
            if (Login.userRoles == "IT传媒部-组长")
            {
                DataTable dtUserSmall = frmMain.dtUser.Copy();
                dtUserSmall.Rows.Add("");
                this.comboBox2.DataSource = dtUserSmall;
                this.comboBox2.DisplayMember = "UserId";
                this.comboBox2.ValueMember = "UserId";
            }
            else
            {
                this.linkLabel1.Visible = false;
                this.comboBox2.Text = Login.userName;
                this.comboBox2.Enabled = false;
            }
        }
        private void button2_Click(object sender, EventArgs e)//提交
        {
            this.dataGridView2.CurrentCell = null;
            if (this.dataGridView2.Rows.Count < 1)
            {
                MessageBox.Show("你想干嘛？？？");
                return;
            }
            if (Login.userRoles == "IT传媒部-组员")
            {
                DataTable dt1 = DBHelper.ExecuteQuery("select CONVERT(varchar(12) , getdate(), 111 )");
                DateTime timeNow = DateTime.Parse(dt1.Rows[0][0].ToString());
                if (DateTime.Parse(this.dataGridView2.Rows[0].Cells[0].Value.ToString()) != timeNow) 
                {
                    MessageBox.Show("只能修改今天的数据，其他数据请联系组长删除");
                    return;
                }
            }
            string sqlstr1 = string.Format("delete from {0} where UpUser='{1}' and UpTime='{2}'",
               getDBName(this.comboBox1.Text),
               this.comboBox2.Text,
               this.dataGridView2.Rows[0].Cells[0].Value);

            DBHelper.ExecuteUpdate(sqlstr1);
            string sqlstr2 = string.Format("insert into {0} values('{1}', '{2}', '{3}', '{4}', '{5}', '{6}', '{7}', '{8}', '{9}', '{10}', '{11}', '{12}', '{13}', '{14}', '{15}','{16}','{17}')",
                getDBName(this.comboBox1.Text),
                this.comboBox2.Text,
                this.dataGridView2.Rows[0].Cells[0].Value,
                this.dataGridView2.Rows[0].Cells[1].Value,
                this.dataGridView2.Rows[0].Cells[2].Value,
                this.dataGridView2.Rows[0].Cells[3].Value,
                this.dataGridView2.Rows[0].Cells[4].Value,
                this.dataGridView2.Rows[0].Cells[5].Value,
                this.dataGridView2.Rows[0].Cells[6].Value,
                this.dataGridView2.Rows[0].Cells[7].Value,
                this.dataGridView2.Rows[0].Cells[8].Value,
                this.dataGridView2.Rows[0].Cells[9].Value,
                this.dataGridView2.Rows[0].Cells[10].Value,
                this.dataGridView2.Rows[0].Cells[11].Value,
                this.dataGridView2.Rows[0].Cells[12].Value,
                this.dataGridView2.Rows[0].Cells[14].Value,
                this.dataGridView2.Rows[0].Cells[13].Value,
                this.dataGridView2.Rows[0].Cells[15].Value);
            int resultNum = DBHelper.ExecuteUpdate(sqlstr2);
            if (resultNum > 0)
            {
                MessageBox.Show("增添成功~~~");
                button1_Click(null, null);
            }
            else
            {
                MessageBox.Show("增添失败~~~");
            }
        }

        private void button1_Click(object sender, EventArgs e)//查询
        {
            string sqlStr = "";
            DBTable = getDBName(this.comboBox1.Text);
            if (this.comboBox1.Text == "")
            {
                return;
            }
            if (DBTable == "SerachTitles_SuCai")
            {
                sqlStr = string.Format("select UpTime as 日期,sum(pfsc) as 漂浮素材, sum(jrys) as 节日元素, sum(xgsc) as 效果素材, sum(zsta) as 装饰图案, sum(ysz) as 艺术字体, "+
                "sum(cxbq) as 促销标签, sum(tbys) as 图标元素, sum(bkwl) as 边框纹理, sum(shkt) as 卡通手绘, sum(bgztx) as 不规则图, sum(pptys) as 幻灯片图, sum(cpys) as 产品元素, "+
                "sum(qt) as 其他, sum(beiyong1) as 备用一, sum(beiyong2) as 备用二  from {0} where UpUser like '%{1}%' and UpTime>'{2}' and UpTime<'{3}' "+
                "group by UpTime order by 日期 asc",
                    DBTable, 
                    this.comboBox2.Text,
                    this.dateTimePicker1.Value,
                    this.dateTimePicker2.Value);
                dt = DBHelper.ExecuteQuery(sqlStr);
            }
            else
            {
                sqlStr = string.Format("select UpTime as 日期, sum(ty) as 通用素材, sum(fzpj) as 服装配件, sum(jpxb) as 精品箱包, sum(hzjm) as 化妆健美, sum(spzb) as 饰品珠宝, "+
                    "sum(hwyd) as 户外运动, sum(zsjj) as 装饰家居, sum(smjd) as 数码家电, sum(myyp) as 母婴用品, sum(spbj) as 食品保健, sum(shxq) as 生活兴趣, sum(xncz) as 虚拟充值, "+
                    "sum(beiyong1) as 备用一, sum(beiyong2) as 备用二, sum(beiyong3) as 备用三 from {0} where UpUser like '%{1}%' and UpTime>'{2}' and UpTime<'{3}' "+
                    "group by UpTime order by 日期 asc",
                    DBTable, 
                    this.comboBox2.Text,
                    this.dateTimePicker1.Value,
                    this.dateTimePicker2.Value);
                dt = DBHelper.ExecuteQuery(sqlStr);
            }
            this.dataGridView1.DataSource = dt;
            //设置上面表格中列的宽度
            for (int i = 0; i < this.dataGridView1.ColumnCount; i++)
            {
                this.dataGridView1.Columns[i].Width = 54;
            }
            this.dataGridView1.Columns[0].Width = 100;
            //查询完成，添加一行汇总数据在最后面
            if (dt.Rows.Count != 0 && DBTable == "SerachTitles_SuCai")
            {
                dt.Rows.Add("1900/10/10",
                    dt.Compute("sum(漂浮素材)", "true"),
                    dt.Compute("sum(节日元素)", "true"),
                    dt.Compute("sum(效果素材)", "true"),
                    dt.Compute("sum(装饰图案)", "true"),
                    dt.Compute("sum(艺术字体)", "true"),
                    dt.Compute("sum(促销标签)", "true"),
                    dt.Compute("sum(图标元素)", "true"),
                    dt.Compute("sum(边框纹理)", "true"),
                    dt.Compute("sum(卡通手绘)", "true"),
                    dt.Compute("sum(不规则图)", "true"),
                    dt.Compute("sum(幻灯片图)", "true"),
                    dt.Compute("sum(产品元素)", "true"),
                    dt.Compute("sum(其他)", "true"),
                    dt.Compute("sum(备用一)", "true"),
                    dt.Compute("sum(备用二)", "true"));

                this.label5.Text = "总量："+(
                    int.Parse(dt.Rows[dt.Rows.Count - 1][1].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][2].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][3].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][4].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][5].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][6].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][7].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][8].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][9].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][10].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][11].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][12].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][13].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][14].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][15].ToString())).ToString();
            }
            if (dt.Rows.Count != 0 && DBTable != "SerachTitles_SuCai")
            {
                dt.Rows.Add("1900/10/10",
                    dt.Compute("sum(通用素材)", "true"),
                    dt.Compute("sum(服装配件)", "true"),
                    dt.Compute("sum(精品箱包)", "true"),
                    dt.Compute("sum(化妆健美)", "true"),
                    dt.Compute("sum(饰品珠宝)", "true"),
                    dt.Compute("sum(户外运动)", "true"),
                    dt.Compute("sum(装饰家居)", "true"),
                    dt.Compute("sum(数码家电)", "true"),
                    dt.Compute("sum(母婴用品)", "true"),
                    dt.Compute("sum(食品保健)", "true"),
                    dt.Compute("sum(生活兴趣)", "true"),
                    dt.Compute("sum(虚拟充值)", "true"),
                    dt.Compute("sum(备用一)", "true"),
                    dt.Compute("sum(备用二)", "true"),
                    dt.Compute("sum(备用三)", "true"));

                this.label5.Text = "总量：" + (
                    int.Parse(dt.Rows[dt.Rows.Count - 1][1].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][2].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][3].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][4].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][5].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][6].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][7].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][8].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][9].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][10].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][11].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][12].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][13].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][14].ToString()) +
                    int.Parse(dt.Rows[dt.Rows.Count - 1][15].ToString())).ToString();
            }
            //点击查询，在下面出现今天的上传框
            if (this.comboBox1.Text == "素材")
            {
                if (this.comboBox2.Text != "")
                {
                    string sqlStrAdd = string.Format("select top 1 UpTime as 日期,pfsc as 漂浮素材, jrys as 节日元素, xgsc as 效果素材, zsta as 装饰图案, ysz as 艺术字体, "+
                        "cxbq as 促销标签, tbys as 图标元素, bkwl as 边框纹理, shkt as 卡通手绘, bgztx as 不规则图, pptys as 幻灯片图, cpys as 产品元素, qt as 其他, "+
                        "beiyong1 as 备用一, beiyong2 as 备用二 from {0} where UpUser='{1}' and UpTime='{2}'",
                            getDBName(this.comboBox1.Text), this.comboBox2.Text,
                            System.DateTime.Now.ToShortDateString());
                    DataTable dtUpDown = DBHelper.ExecuteQuery(sqlStrAdd);
                    if (dtUpDown.Rows.Count == 0)
                    {
                        dtUpDown.Rows.Add(System.DateTime.Now.ToShortDateString(),0,0,0,0,0,0,0,0,0,0,0,0,0,0,0);
                    }
                    this.dataGridView2.DataSource = dtUpDown;
                    this.button2.Enabled = true;
                }
                else
                {
                    this.button2.Enabled = false;
                }
            }
            else
            {
                if (this.comboBox2.Text != "")
                {
                    string sqlStrAdd = string.Format("select top 1 UpTime as 日期, ty as 通用素材, fzpj as 服装配件, jpxb as 精品箱包, hzjm as 化妆健美, spzb as 饰品珠宝, " +
                        "hwyd as 户外运动, zsjj as 装饰家居, smjd as 数码家电, myyp as 母婴用品, spbj as 食品保健, shxq as 生活兴趣, xncz as 虚拟充值, beiyong1 as 备用一, "+
                        "beiyong2 as 备用二, beiyong3 as 备用三 from {0} where UpUser='{1}' and UpTime='{2}'",
                            getDBName(this.comboBox1.Text), this.comboBox2.Text,
                            System.DateTime.Now.ToShortDateString());
                    DataTable dtUpDown = DBHelper.ExecuteQuery(sqlStrAdd);
                    if (dtUpDown.Rows.Count == 0)
                    {
                        dtUpDown.Rows.Add(System.DateTime.Now.ToShortDateString(),0,0,0,0,0,0,0,0,0,0,0,0,0,0,0);
                    }
                    this.dataGridView2.DataSource = dtUpDown;
                    this.button2.Enabled = true;
                }
                else
                {
                    this.button2.Enabled = false;
                }
            }
            //设置下面表格中列的宽度
            for (int i = 0; i < this.dataGridView2.ColumnCount; i++)
            {

                this.dataGridView2.Columns[i].Width = 54;
            }
            if (this.dataGridView2.Rows.Count>0) 
            {
                this.dataGridView2.Columns[0].Width = 100;
            }
            for (int i = 0; i < this.dataGridView1.ColumnCount; i++)
            {
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            for (int i = 0; i < this.dataGridView2.ColumnCount; i++)
            {
                this.dataGridView2.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }
        /// <summary>
        /// 传入名称，返回数据库中对应的表名
        /// </summary>
        /// <param name="name"></param>
        /// <returns></returns>
        private string getDBName(string name)
        {
            string DBTable1 = "";
            switch (name)
            {
                case "素材":
                    DBTable1 = "SerachTitles_SuCai";
                    break;
                case "banner海报":
                    DBTable1 = "SerachTitles_BannerHB";
                    break;
                case "店铺首页":
                    DBTable1 = "SerachTitles_DPSY";
                    break;
                case "直通车主图钻展":
                    DBTable1 = "SerachTitles_ZTCZUZZ";
                    break;
                case "手机端背景":
                    DBTable1 = "SerachTitles_SJDBJ";
                    break;
                case "详情页":
                    DBTable1 = "SerachTitles_XQY";
                    break;
                case "专题背景":
                    DBTable1 = "SerachTitles_ZTBJ";
                    break;
                default:
                    DBTable1 = "SerachTitles_SuCai";
                    break;
            }
            return DBTable1;
        }
        //验证是否输入的是数字
        private void dataGridView2_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            if (e.RowIndex > -1 && e.ColumnIndex > 1)
            {
                DataGridView grid = (DataGridView)sender;
                grid.Rows[e.RowIndex].ErrorText = "";
                for (int i = 1; i < grid.ColumnCount;i++ ) 
                {
                    Int32 newInteger = 0;
                    if (!int.TryParse(e.FormattedValue.ToString(), out newInteger))
                    {
                        e.Cancel = true;
                        grid.Rows[e.RowIndex].ErrorText = "请输入数字";
                        MessageBox.Show("请输入数字!");
                        return;
                    } 
                }
            }
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (MessageBox.Show("你确定要导出数据至excel表格吗?", "导出提醒", MessageBoxButtons.YesNo) == DialogResult.No)
                return;
            string sqlStr = "";
            if (this.comboBox1.Text == "")
            {
                MessageBox.Show("类别不能为空");
                return;
            }
            string tabName = getDBName(this.comboBox1.Text);
            if (tabName == "SerachTitles_SuCai")
            {
                sqlStr = string.Format("select CONVERT(varchar(100), UpTime, 23) as 日期,sum(pfsc) as 漂浮素材, sum(jrys) as 节日元素, sum(xgsc) as 效果素材, sum(zsta) as 装饰图案, sum(ysz) as 艺术字体, " +
                "sum(cxbq) as 促销标签, sum(tbys) as 图标元素, sum(bkwl) as 边框纹理, sum(shkt) as 卡通手绘, sum(bgztx) as 不规则图, sum(pptys) as 幻灯片图, sum(cpys) as 产品元素, "+
                "sum(qt) as 其他, sum(beiyong1) as 备用一, sum(beiyong2) as 备用二  from {0} where UpUser like '%{1}%' and UpTime>'{2}' and UpTime<'{3}' "+
                "group by UpTime order by 日期 asc",
                    tabName, 
                    this.comboBox2.Text,
                    this.dateTimePicker1.Value,
                    this.dateTimePicker2.Value);
            }
            else
            {
                sqlStr = string.Format("select CONVERT(varchar(100), UpTime, 23) as 日期, sum(ty) as 通用素材, sum(fzpj) as 服装配件, sum(jpxb) as 精品箱包, sum(hzjm) as 化妆健美, sum(spzb) as 饰品珠宝, " +
                    "sum(hwyd) as 户外运动, sum(zsjj) as 装饰家居, sum(smjd) as 数码家电, sum(myyp) as 母婴用品, sum(spbj) as 食品保健, sum(shxq) as 生活兴趣, sum(xncz) as 虚拟充值, "+
                    "sum(beiyong1) as 备用一, sum(beiyong2) as 备用二, sum(beiyong3) as 备用三 from {0} where UpUser like '%{1}%' and UpTime>'{2}' and UpTime<'{3}' "+
                    "group by UpTime order by 日期 asc",
                    tabName, 
                    this.comboBox2.Text,
                    this.dateTimePicker1.Value,
                    this.dateTimePicker2.Value);
            }
            DataTable dt1 = DBHelper.ExecuteQuery(sqlStr);
            if (dt1.Rows.Count > 0)
            {
                if (Directory.Exists(@"D:\\分类数据表(软件导出)") == false)//如果不存在就创建file文件夹
                {
                    Directory.CreateDirectory(@"D:\\分类数据表(软件导出)");
                }
                try
                {
                    Tools.dataTableToCsv(dt1, @"D:\\分类数据表(软件导出)\" +
                    this.dateTimePicker1.Value.ToString("yyyyMMdd") + "至" + this.dateTimePicker2.Value.ToString("yyyyMMdd") +"--"+
                    this.comboBox1.Text + "--" + this.comboBox2.Text + "--"+
                    "分类数据表" + ".xls");
                    MessageBox.Show("导出完成,路径为：D:\\\\分类数据表(软件导出)\\" +
                    this.dateTimePicker1.Value.ToString("yyyyMMdd") + "至" + this.dateTimePicker2.Value.ToString("yyyyMMdd") + "--" +
                    this.comboBox1.Text + "--" + this.comboBox2.Text + "--" +
                    "分类数据表" + ".xls");
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
