using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace VS_code_copy_to_Word_clutter_removal_tool
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //获取剪切板内容
            IDataObject dataObject = Clipboard.GetDataObject();
            if (dataObject.GetDataPresent(DataFormats.Rtf))
            {
                //去除RTF格式
                string rtf = dataObject.GetData(DataFormats.Rtf) as string;
                //以Regex.Replace去除多余字符(注: 不管是否有问题，一律强制处理)
                string fixedRtf =Regex.Replace(rtf, @"\\uinput2(?<uc>\\u-?\d*)\s..",(m) =>
                {
                    return m.Groups["uc"].Value + "?";
                });
                //另建新DataObject物件
                DataObject newDataObject = new DataObject();
                //RTF格式用修正后的字符串，其余依原值
                foreach (String t in dataObject.GetFormats())
                    newDataObject.SetData(t,
                    t == "Rich Text Format" ? fixedRtf :
                    dataObject.GetData(t));
                //将修正后内容写入剪贴簿
                Clipboard.SetDataObject(newDataObject, true);
                MessageBox.Show("Take Out Pessy Code successful!");
            }
        }
    }
}
