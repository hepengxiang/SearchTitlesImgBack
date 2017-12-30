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
    public partial class UserManage : Form
    {
        public UserManage()
        {
            InitializeComponent();
        }
        private void UserManage_Load(object sender, EventArgs e)
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
                this.comboBox1.Text = Login.userName;
                this.comboBox1.Enabled = false;
                this.button1.Enabled = false;
                this.button2.Enabled = false;
                this.button4.Enabled = false;

                this.textBox1.Text = Login.userName;
                this.comboBox2.Text = Login.userRoles;
                this.comboBox2.Enabled = false;
            }
        }
        private void button6_Click(object sender, EventArgs e)//查询
        {

            string sqlstr1 = string.Format("select a.UpTime as 上传时间," +
                " coalesce(sum(a.pfsc + a.jrys+ a.xgsc+ a.zsta+a.ysz+a.cxbq+a.tbys+a.bkwl+a.shkt+a.bgztx+a.pptys+a.cpys+a.qt+a.beiyong1+a.beiyong2),0) " +
                "as 素材总数 ," +
                " coalesce(sum(b.ty+b.fzpj+b.jpxb+b.hzjm+b.spzb+b.hwyd+b.zsjj+b.smjd+b.myyp+b.spbj+b.shxq+b.xncz+b.beiyong1+b.beiyong2+b.beiyong3),0) " +
                "as 海报总数 ," +
                " coalesce(sum(c.ty+c.fzpj+c.jpxb+c.hzjm+c.spzb+c.hwyd+c.zsjj+c.smjd+c.myyp+c.spbj+c.shxq+c.xncz+c.beiyong1+c.beiyong2+c.beiyong3),0) " +
                "as 店铺首页 ," +
                " coalesce(sum(d.ty+d.fzpj+d.jpxb+d.hzjm+d.spzb+d.hwyd+d.zsjj+d.smjd+d.myyp+d.spbj+d.shxq+d.xncz+d.beiyong1+d.beiyong2+d.beiyong3),0) " +
                "as 直通车图 ," +
                " coalesce(sum(e.ty+e.fzpj+e.jpxb+e.hzjm+e.spzb+e.hwyd+e.zsjj+e.smjd+e.myyp+e.spbj+e.shxq+e.xncz+e.beiyong1+e.beiyong2+e.beiyong3),0) " +
                "as 手机端图 ," +
                " coalesce(sum(f.ty+f.fzpj+f.jpxb+f.hzjm+f.spzb+f.hwyd+f.zsjj+f.smjd+f.myyp+f.spbj+f.shxq+f.xncz+f.beiyong1+f.beiyong2+f.beiyong3),0) " +
                "as 详情页图 ," +
                " coalesce(sum(g.ty+g.fzpj+g.jpxb+g.hzjm+g.spzb+g.hwyd+g.zsjj+g.smjd+g.myyp+g.spbj+g.shxq+g.xncz+g.beiyong1+g.beiyong2+g.beiyong3),0) " +
                "as 专题页图  " +
                "from " +
                "SerachTitles_SuCai a " +
                "full join SerachTitles_BannerHB b on a.UpTime = b.UpTime and a.UpUser = b.UpUser " +
                "full join SerachTitles_DPSY c on a.UpTime = c.UpTime and a.UpUser = c.UpUser " +
                "full join SerachTitles_ZTCZUZZ d on a.UpTime = d.UpTime and a.UpUser = d.UpUser " +
                "full join SerachTitles_SJDBJ e on a.UpTime = e.UpTime and a.UpUser = e.UpUser " +
                "full join SerachTitles_XQY f on a.UpTime = f.UpTime and a.UpUser = f.UpUser " +
                "full join SerachTitles_ZTBJ g on a.UpTime = g.UpTime  and a.UpUser = g.UpUser " +
                " where a.UpUser like '%{0}%' and a.UpTime>'{1}' and a.UpTime<'{2}'" +
                " group by a.UpTime" +
                " order by a.UpTime asc",
                this.comboBox1.Text,
                this.dateTimePicker1.Value.ToShortDateString(),
                this.dateTimePicker2.Value.ToShortDateString());
            DataTable dtOther = DBHelper.ExecuteQuery(sqlstr1);
            dtOther.Rows.Add("1900/10/10",
                    dtOther.Compute("sum(素材总数)", "true"),
                    dtOther.Compute("sum(海报总数)", "true"),
                    dtOther.Compute("sum(店铺首页)", "true"),
                    dtOther.Compute("sum(直通车图)", "true"),
                    dtOther.Compute("sum(手机端图)", "true"),
                    dtOther.Compute("sum(详情页图)", "true"),
                    dtOther.Compute("sum(专题页图)", "true"));
            dtOther = rowsAllNum(dtOther);
            this.dataGridView1.DataSource = dtOther;
            for (int i = 0; i < this.dataGridView1.ColumnCount; i++)
            {
                this.dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
        }
        /// <summary>
        /// 给datatable添加一列汇总
        /// </summary>
        /// <param name="dtss"></param>
        /// <returns></returns>
        private DataTable rowsAllNum(DataTable dtss)
        {
            int tem = 0;
            int rNum = dtss.Rows.Count;//行
            int cNum = dtss.Columns.Count;//列
            dtss.Columns.Add("总量", typeof(int));
            for (int i = 0; i < rNum; i++)
            {
                for (int j = 1; j < cNum; j++)
                {
                    if (!String.IsNullOrEmpty(dtss.Rows[i][j].ToString()))
                    {
                        tem += int.Parse(dtss.Rows[i][j].ToString());
                    }
                    string a = dtss.Rows[i][j].ToString();
                    int b = cNum + 1;
                }
                dtss.Rows[i][cNum] = tem;
                tem = 0;
            }
            return dtss;
        }

        private void button1_Click(object sender, EventArgs e)//增加
        {
            string sqlstr = string.Format("insert into SerachTitles_Users values('{0}','{1}','{2}')", this.textBox1.Text, this.textBox2.Text, this.comboBox2.Text);
            int resultNum = DBHelper.ExecuteUpdate(sqlstr);
            if (resultNum > 0)
            {
                MessageBox.Show("新增成功");
            }
            else
            {
                MessageBox.Show("插入失败，用户名重复或用户名有特殊字符~~~");
            }
        }

        private void button2_Click(object sender, EventArgs e)//删除
        {
            if (MessageBox.Show("你确定要删除【" + this.textBox1.Text + "】吗?", "删除提醒", MessageBoxButtons.YesNo) == DialogResult.No)
                return;
            string sqlstr = string.Format("delete from SerachTitles_Users where UserId='{0}'", this.textBox1.Text);
            int resultNum = DBHelper.ExecuteUpdate(sqlstr);
            if (resultNum > 0)
            {
                MessageBox.Show("删除成功");
            }
            else
            {
                MessageBox.Show("删除失败，用户名是否正确~~~");
            }
        }

        private void button3_Click(object sender, EventArgs e)//修改
        {
            //if (MessageBox.Show("你确定要删除成员[" + this.textBox10.Text + "]吗?", "删除提醒", MessageBoxButtons.YesNo) == DialogResult.No)
            //return;
            string sqlstr = string.Format("update SerachTitles_Users set PSW = '{1}' where UserId = '{0}'", this.textBox1.Text, this.textBox2.Text, this.comboBox2.Text);
            int resultNum = DBHelper.ExecuteUpdate(sqlstr);
            if (resultNum > 0)
            {
                MessageBox.Show("更新成功");
            }
            else
            {
                MessageBox.Show("更新失败，用户名是否重复~~~");
            }
        }

        private void button4_Click(object sender, EventArgs e)//查询
        {
            string sqlstr = string.Format("select UserId as 用户名, PSW as 密码, Roles as 职位 from SerachTitles_Users where UserId like '%{0}%' and Roles like '%{1}%'", this.textBox1.Text, this.comboBox2.Text);
            DataTable dts = DBHelper.ExecuteQuery(sqlstr);
            this.dataGridView2.DataSource = dts;
        }
    }
}
