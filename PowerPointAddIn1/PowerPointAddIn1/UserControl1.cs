﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PowerPointAddIn1
{
    public partial class UserControl1 : UserControl
    {
        public UserControl1()
        {
            InitializeComponent();
        }

        public UserControl1(string v)
        {
        }

        private void button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Test");
        }

        private void UserControl1_Load(object sender, EventArgs e)
        {

        }
    }
}
