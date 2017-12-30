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
    public partial class frmMain : Form
    {

        public static int PrimaryBaseNum;
        public static double PrimaryReward;
        public static int MiddleBaseNum;
        public static double MiddleReward;
        public static int SeniorBaseNum;
        public static double SeniorReward;
        public static DataTable dtUser;
        public static Dictionary<string, int> TopButtonKY;
        public frmMain()
        {
            InitializeComponent();
        }
        private void frmMain_Load(object sender, EventArgs e)
        {
            
            TopButtonKY = new Dictionary<string, int>();
            TopButtonKY.Add("SelectWordsBtn",0);
            TopButtonKY.Add("AddWordsBtn",1);
            TopButtonKY.Add("UploadNumBtn",2);
            TopButtonKY.Add("UploadDataTableBtn",3);
            TopButtonKY.Add("DeductionDataTableBtn",4);
            TopButtonKY.Add("DeductionCountBtn",5);
            TopButtonKY.Add("UserManageBtn",6);
            frmMain.dtUser = DBHelper.ExecuteQuery("select UserId from SerachTitles_Users");
            getBaseSalaryNum();
            Button[] btns = { SelectWordsBtn, AddWordsBtn, UploadNumBtn, UploadDataTableBtn, DeductionDataTableBtn, DeductionCountBtn, UserManageBtn };
            addWindowToPanel(btns);
            addLoginToPanel();//加入登陆界面
        }
        
        /// <summary>
        /// 获取奖励基数
        /// </summary>
        private void getBaseSalaryNum() 
        {
            DataTable dtBaseNum = DBHelper.ExecuteQuery("select * from BaseDate");
            frmMain.PrimaryBaseNum = int.Parse(dtBaseNum.Rows[0][0].ToString());
            frmMain.PrimaryReward = double.Parse(dtBaseNum.Rows[0][1].ToString());
            frmMain.MiddleBaseNum = int.Parse(dtBaseNum.Rows[0][2].ToString());
            frmMain.MiddleReward = double.Parse(dtBaseNum.Rows[0][3].ToString());
            frmMain.SeniorBaseNum = int.Parse(dtBaseNum.Rows[0][4].ToString());
            frmMain.SeniorReward = double.Parse(dtBaseNum.Rows[0][5].ToString());
        }
        /// <summary>
        /// 加入登陆界面
        /// </summary>
        private void addLoginToPanel() 
        {
            Type t = Type.GetType("SearchTitlesImgBack.Login");
            object obj = System.Activator.CreateInstance(t);
            (obj as Form).TopLevel = false;
            (obj as Form).Parent = this.panelLogin;
            (obj as Form).Dock = DockStyle.Fill;
            panelLogin.Controls.Add((obj as Form));
            (obj as Form).Show();
        }
        /// <summary>
        /// 将子界面加入到主界面中的panel中
        /// </summary>
        /// <param name="btns"></param>
        private void addWindowToPanel(Button[] btns) 
        {
            foreach (Button btn in btns)
            {
                btn.Click += new EventHandler(buttonclick);
                btn.MouseEnter += new EventHandler(buttonenter);
                btn.MouseLeave += new EventHandler(buttonleave);
                Type t = Type.GetType("SearchTitlesImgBack." + btn.Name.Replace("Btn", ""));
                object obj = System.Activator.CreateInstance(t);//创建t类的实例 "obj"
                (obj as Form).TopLevel = false;
                (obj as Form).Parent = this.panel1;
                (obj as Form).Dock = DockStyle.Fill;
                foreach (Control control in (obj as Form).Controls)
                {
                    if (control is Button)
                    {
                        Button btnbefore = (Button)control;
                        changeButtonStyle(btnbefore);
                    }
                    else if (control is GroupBox)
                    {
                        GroupBox gboxs = (GroupBox)control;
                        foreach (Control gbox in gboxs.Controls)
                        {
                            if (gbox is Button)
                            {
                                Button gboxbtn = (Button)gbox;
                                changeButtonStyle(gboxbtn);
                            }
                        }
                    }
                    else if (control is DataGridView) 
                    {
                        DataGridView dgv = (DataGridView)control;
                        changeDGVStyle(dgv);
                    }
                }
                panel1.Controls.Add((obj as Form));
                (obj as Form).Hide();
            }
        }
        /// <summary>
        /// 设置子按钮的显示样式和绑定事件
        /// </summary>
        /// <param name="btn"></param>
        private void changeButtonStyle(Button btn) 
        {
            btn.BackgroundImage = imageListSmallBtn.Images[0];//添加页面内部按钮图片
            btn.ForeColor = Color.Gold;
            btn.MouseEnter += new EventHandler(buttonenterbefore);
            btn.MouseLeave += new EventHandler(buttonleavebefore);
            btn.MouseDown += new MouseEventHandler(buttondownbefore);
            btn.MouseUp += new MouseEventHandler(buttondownup);
        }
        /// <summary>
        /// 设置DataGridView的显示样式
        /// </summary>
        /// <param name="dgv"></param>
        private void changeDGVStyle(DataGridView dgv) 
        {
            //设置不能点击排序

            //设置基数行的背景色
            dgv.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(244, 232, 184);
            //设置内容背景色
            dgv.BackgroundColor = Color.FromArgb(211, 199, 147);
            //去掉所有边框线条
            dgv.BorderStyle = BorderStyle.None;
            dgv.CellBorderStyle = DataGridViewCellBorderStyle.None;
            dgv.RowHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            dgv.ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.None;
            //设置列标题样式
            dgv.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(211, 199, 147);
            dgv.ColumnHeadersDefaultCellStyle.Font = new Font("楷体", 10, FontStyle.Bold);
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            //设置行标题样式
            dgv.RowHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.RowHeadersDefaultCellStyle.BackColor = Color.FromArgb(211, 199, 147);
            dgv.RowHeadersDefaultCellStyle.Font = new Font("楷体", 12, FontStyle.Bold);
            dgv.RowHeadersDefaultCellStyle.SelectionBackColor = Color.Black;
            dgv.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            //设置行单元格的默认样式
            dgv.RowsDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dgv.RowsDefaultCellStyle.SelectionBackColor = Color.Black;
            dgv.RowsDefaultCellStyle.SelectionForeColor = Color.Gold;
            dgv.RowsDefaultCellStyle.Font = new Font("楷体", 10, FontStyle.Bold);
            //关闭默认使用用户主题
            dgv.EnableHeadersVisualStyles = false;
            //设置未设置其他样式的默认样式(偶数行背景色)
            dgv.DefaultCellStyle.BackColor = Color.FromArgb(211, 199, 147);
        }
        private void buttonleavebefore(object sender, EventArgs e)//页面内部按钮离开
        {
            if ((sender as Button).BackgroundImage != null)
                (sender as Button).BackgroundImage.Dispose();
            (sender as Button).BackgroundImage = imageListSmallBtn.Images[0];
        }
        private void buttonenterbefore(object sender, EventArgs e)//页面内部按钮进入
        {
            if ((sender as Button).BackgroundImage != null)
                (sender as Button).BackgroundImage.Dispose();
            (sender as Button).BackgroundImage = imageListSmallBtn.Images[1];
        }
        private void buttondownbefore(object sender, EventArgs e)//页面内部按钮按下
        {
            if ((sender as Button).BackgroundImage != null)
                (sender as Button).BackgroundImage.Dispose();
            (sender as Button).BackgroundImage = imageListSmallBtn.Images[2];
        }
        private void buttondownup(object sender, EventArgs e)//页面内部按钮松开
        {
            if ((sender as Button).BackgroundImage != null)
                (sender as Button).BackgroundImage.Dispose();
            (sender as Button).BackgroundImage = imageListSmallBtn.Images[0];
        }
        private void buttonleave(object sender, EventArgs e)//顶部按钮离开
        {
            (sender as Button).Width = 80;
            (sender as Button).Height = 80;
            (sender as Button).BackgroundImage = imageListTopLeave.Images[TopButtonKY[(sender as Button).Name]];
        }
        private void buttonenter(object sender, EventArgs e)//顶部按钮进入
        {
            (sender as Button).Width = 100;
            (sender as Button).Height = 100;
            (sender as Button).BackgroundImage = imageListTopEnter.Images[TopButtonKY[(sender as Button).Name]];
        }
        private void buttonclick(object sender, EventArgs e)//顶部按钮点击
        {
            if (Login.isLogin)
            {
                foreach (Control control in panel1.Controls)
                {
                    if (control.Name == (sender as Button).Name.Replace("Btn", ""))
                    {
                        control.Show();
                    }
                    else
                    {
                        control.Hide();
                    }
                }
            }
        }
        //退出程序
        private void BtnExit_Click(object sender, EventArgs e)
        {
            System.Environment.Exit(0);
        }
        //最小化程序
        private void BtnWinMin_Click(object sender, EventArgs e)
        {
            this.WindowState = FormWindowState.Minimized;
        }
        Point downPoint;
        private void frmMain_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                this.Location = new Point(this.Location.X + e.X - downPoint.X,
                    this.Location.Y + e.Y - downPoint.Y);
            }
        }

        private void frmMain_MouseDown(object sender, MouseEventArgs e)
        {
            downPoint = new Point(e.X, e.Y);
        }
        /// <summary>
        /// 解决闪屏问题
        /// </summary>
        protected override CreateParams CreateParams
        {
            get
            {
                CreateParams paras = base.CreateParams;
                paras.ExStyle |= 0x02000000;

                const int WS_MINIMIZEBOX = 0x00020000;  // Winuser.h中定义
                paras.Style = paras.Style | WS_MINIMIZEBOX;   // 允许最小化操作

                return paras;
            }
        }
    }
}
