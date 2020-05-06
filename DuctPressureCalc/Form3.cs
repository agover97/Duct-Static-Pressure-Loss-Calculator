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
    public partial class Form3 : Form
    {
        public Form3()
        {
            InitializeComponent();
        }

        private Form1 mainForm = null;
        public Form3(Form callingForm)
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
            else if (this.radioButton14.Checked)
            {
                radio = "14";
            }
            else if (this.radioButton15.Checked)
            {
                radio = "15";
            }
            else if (this.radioButton16.Checked)
            {
                radio = "16";
            }
            else if (this.radioButton17.Checked)
            {
                radio = "17";
            }
            else if (this.radioButton18.Checked)
            {
                radio = "18";
            }
            else if (this.radioButton19.Checked)
            {
                radio = "19";
            }
            else if (this.radioButton20.Checked)
            {
                radio = "20";
            }
            else if (this.radioButton21.Checked)
            {
                radio = "21";
            }
            else if (this.radioButton22.Checked)
            {
                radio = "22";
            }
            this.mainForm.componentSelect = radio;
            this.Close();


        }
    }
}
