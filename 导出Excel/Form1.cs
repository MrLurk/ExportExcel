﻿using OutPutExcel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace 导出Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            new ExcelCommand().ExcelOut($@"{DateTime.Now.ToString("yyyyMMddhhmmss")}.xlsx", @"Template/template.xlsx");
        }
    }
}
