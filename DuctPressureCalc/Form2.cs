using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DuctPressureCalc
{
    public partial class Form2 : Form
    {
        public Form2()
        {
            InitializeComponent();
        }

        private Form1 mainForm = null;
        public Form2(Form callingForm)
        {
            mainForm = callingForm as Form1;
            InitializeComponent();
        }
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            string radio = "-1";
            if (this.radioButton1.Checked)
            {
                radio = "1";
            } 
            else if (this.radioButton2.Checked)
            {
                radio = "2";
            }
            else if (this.radioButton3.Checked)
            {
                radio = "3";
            }
            else if (this.radioButton4.Checked)
            {
                radio = "4";
            }
            else if (this.radioButton5.Checked)
            {
                radio = "5";
            }
            else if (this.radioButton6.Checked)
            {
                radio = "6";
            }
            else if (this.radioButton7.Checked)
            {
                radio = "7";
            }
            else if (this.radioButton8.Checked)
            {
                radio = "8";
            }
            else if (this.radioButton9.Checked)
            {
                radio = "9";
            }
            else if (this.radioButton10.Checked)
            {
                radio = "10";
            }
            else if (this.radioButton11.Checked)
            {
                radio = "11";
            }
            else if (this.radioButton12.Checked)
            {
                radio = "12";
            }
            else if (this.radioButton13.Checked)
            {
                radio = "13";
            }
            this.mainForm.radioSelect = radio;
            this.Close();
           

        }
    }
}
