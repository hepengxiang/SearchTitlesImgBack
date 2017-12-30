using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Drawing;
using CsharpHttpHelper;
using HtmlAgilityPack;
using System.IO;
using System.Data;

namespace SearchTitlesImgBack
{
    class Tools
    {
        /// <summary>
        /// 遍历所有按钮的颜色，重置其他按钮的颜色
        /// </summary>
        /// <param name="clickButton"></param>
        public static void ChangeForeColor(Button clickButton)
        {
            foreach (Control tmpControl in clickButton.Parent.Controls)
            {
                if (tmpControl is Button)
                {
                    tmpControl.ForeColor = Color.Gold;
                }
            }
            clickButton.ForeColor = Color.Red;
        }
        /// <summary>
        /// 导出数据到excel
        /// </summary>
        /// <param name="table"></param>
        /// <param name="file"></param>
        public static void dataTableToCsv(DataTable table, string file)
        {
            string title = "";
            FileStream fs = new FileStream(file, FileMode.OpenOrCreate);
            //FileStream fs1 = File.Open(file, FileMode.Open, FileAccess.Read);
            StreamWriter sw = new StreamWriter(new BufferedStream(fs), System.Text.Encoding.Default);
            for (int i = 0; i < table.Columns.Count; i++)
            {
                title += table.Columns[i].ColumnName + "\t"; //栏位：自动跳到下一单元格
            }
            if (title != "")
            {
                title = title.Substring(0, title.Length - 1) + "\n";
                sw.Write(title);
                foreach (DataRow row in table.Rows)
                {
                    string line = "";
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        line += row[i].ToString().Trim() + "\t"; //内容：自动跳到下一单元格
                    }
                    line = line.Substring(0, line.Length - 1) + "\n";
                    sw.Write(line);
                }
            }
            sw.Close();
            fs.Close();
        }
        public static string getWeek(DateTime dtTime) 
        {
            string weekstr = dtTime.DayOfWeek.ToString();
            switch (weekstr)
            {
                case "Monday": weekstr = "星期一"; break;
                case "Tuesday": weekstr = "星期二"; break;
                case "Wednesday": weekstr = "星期三"; break;
                case "Thursday": weekstr = "星期四"; break;
                case "Friday": weekstr = "星期五"; break;
                case "Saturday": weekstr = "星期六"; break;
                case "Sunday": weekstr = "星期日"; break;
            }
            return weekstr;
        }
    }
}
