using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApplication2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            //if (Directory.GetCurrentDirectory() != "")
            //{
            //    MessageBox.Show("檔案放置路徑錯誤");
            //    this.Close();
            //    Environment.Exit(Environment.ExitCode);
            //}
            InitializeComponent();
            for (int i = 0; i < this._type.Length; i++)
            {
                this.comboBox1.Items.Add(this._type[i]);
            }
        }
        String _dirdirectorypath = @".\\";//測試用的路徑 到時候要改
        String[] _type = new string[1] {"M:同時符合多項或其他" };
        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private string _select
        {
            get
            {
                string[] array = this._type[this.comboBox1.SelectedIndex].Split(':');
                return array[0];
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1) { MessageBox.Show("請選擇類型"); return; }
            if (textBox1.Text == "") { MessageBox.Show("請輸入名稱"); return; }
            String ntime = DateTime.Now.ToString("yyMMdd");
            String[] directories = Directory.GetDirectories(_dirdirectorypath);
            for (int i = 0; i < directories.Length; i++)
                directories[i] = Path.GetFileNameWithoutExtension(directories[i]);
            List<String> list = new List<String>();
            String newname = "";
            foreach (string selectArray in directories)
            {
                if (selectArray.StartsWith(this._select))
                {
                    String[] _selectArray = selectArray.Split('_');
                    list.Add((_selectArray[0]));
                }

            }
            if (list.Count == 0)
            {
                newname = ntime + "01";
            }
            else
            {
                list.Sort();
                String SelectTimeCounter = list[list.Count - 1];
                if (SelectTimeCounter.Substring(1, 2) == ntime.Substring(0, 2))
                {
                    if (SelectTimeCounter.Substring(3, 2) == ntime.Substring(2, 2))
                    {
                        newname = (Convert.ToInt32(SelectTimeCounter.Substring(1, SelectTimeCounter.Length - 1)) + 1).ToString();
                    }
                    else newname = ntime + "01";
                }
                else newname = ntime + "01";
            }
            try
            {
                if (MES_Check.Checked == true)                //是否為MES
                    textBox1.Text = "MES" + textBox1.Text;
                Directory.CreateDirectory(".\\" + this._select + newname + "_" + textBox1.Text);//創資料夾
                Process.Start(".\\" + this._select + newname + "_" + textBox1.Text);//開資料夾
                File.Copy(".\\售服處理form.docx", ".\\" + this._select + newname + "_" + textBox1.Text + //移動Form表
                          "\\" + this._select + newname + "_" + textBox1.Text + ".docx", true);//Rename Form表
                Process.Start(".\\" + this._select + newname + "_" + textBox1.Text + //開啟 Form表
                              "\\" + this._select + newname + "_" + textBox1.Text + ".docx");
                Application.Exit();
            }
            catch (ArgumentException)
            {
                MessageBox.Show("偵測到輸入不法字元\r\n請確認名稱是否正確", "error");
                return;
            }
            catch (NotSupportedException)
            {
                MessageBox.Show("偵測到輸入不法字元\r\n請確認名稱是否正確", "error");
                return;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
