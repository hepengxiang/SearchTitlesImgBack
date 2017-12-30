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
    public partial class Login : Form
    {
        public static bool isLogin = false;//暂时去掉登陆界面false
        public static string userName = "";
        public static string userRoles = "";
        private static int submitNum = 0;
        public Login()
        {
            InitializeComponent();
        }

        private void button1_MouseEnter(object sender, EventArgs e)
        {
            if ((sender as Button).BackgroundImage != null)
                (sender as Button).BackgroundImage.Dispose();
            (sender as Button).BackgroundImage = imageListLoginBtn.Images[1];
        }

        private void button1_MouseLeave(object sender, EventArgs e)
        {
            if ((sender as Button).BackgroundImage != null)
                (sender as Button).BackgroundImage.Dispose();
            (sender as Button).BackgroundImage = imageListLoginBtn.Images[0];
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.errorProvider1.Clear();
            string userStr = this.textBox1.Text.Trim().Replace("'","").Replace(" ","");
            if(userStr == "")
            {
                this.errorProvider1.SetError(this.textBox1,"用户名不能为空");
                return;
            }
            string pswStr = this.textBox2.Text.Trim().Replace("'", "").Replace(" ", "");
            if (pswStr == "")
            {
                this.errorProvider1.SetError(this.textBox2, "密码不能为空");
                return;
            }
            string sqlstr = string.Format("select * from SerachTitles_Users where UserId = '{0}' and PSW = '{1}'", userStr, pswStr);
            DataTable dt = DBHelper.ExecuteQuery(sqlstr);
            if (dt.Rows.Count > 0)
            {
                Login.userName = dt.Rows[0][0].ToString();
                Login.userRoles = dt.Rows[0][2].ToString();
                Login.isLogin = true;
                //添加素材空数据
                string sqlsc = string.Format("insert into SerachTitles_SuCai values('{0}',CONVERT(varchar(100), getdate(), 23),0,0,0,0,0,0,0,0,0,0,0,0,0,0,0)", Login.userName);
                DBHelper.ExecuteUpdate(sqlsc); 
                this.Parent.Hide();
            }
            else
            {
                submitNum++;
                this.errorProvider1.SetError(this.button1,"用户名或密码错误！");
                if (submitNum>2) 
                {
                    System.Environment.Exit(0);
                }
            }
        }
    }
}
