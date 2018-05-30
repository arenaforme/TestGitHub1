using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace PowerPointAddIn1
{
    public partial class UserControl2 : UserControl
    {
        public UserControl2()
        {
            InitializeComponent();
        }

        private void UserControl2_Load(object sender, EventArgs e)
        {

        }

        private void webBrowser1_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.textBox1.Text = "";
        }
        private PowerPoint.Shape textbox;
        private void button2_Click(object sender, EventArgs e)
        {
            MessageBox.Show(this.textBox1.Text);
            PowerPoint.Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
            PowerPoint.Slide slide = slides[1];
            textbox = slide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 50, 100, 600, 50);//向当前PPT添加文本框
            textbox.TextFrame.TextRange.Text = this.textBox1.Text;//设置文本框的内容
            textbox.TextFrame.TextRange.Font.Size = 48;//设置文本字体大小
            textbox.TextFrame.TextRange.Font.Color.RGB = Color.DarkViolet.ToArgb();//设置文本颜色
        }
    }
}
