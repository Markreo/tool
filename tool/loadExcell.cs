using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace tool
{
    public partial class loadExcell : Form
    {
        string _fileName = "";
        public string _sheet = "0";
        public string _type = "B";
        public string _number = "C";
        public string _la;
        public string _long;

        public loadExcell(string fileName = "")
        {
            InitializeComponent();
            if(fileName != "")
            {
                lFileName.Text = fileName;
                Console.WriteLine(fileName);
                _fileName = fileName;
                txtSheet.Text = "2";
                txtType.Text = "3";
                txtNumber.Text = "2";
                txtLong.Text = "4";
                txtLa.Text = "5";
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            _sheet = txtSheet.Text;
            _type = txtType.Text;
            _number = txtNumber.Text;
            _la = txtLa.Text;
            _long = txtLong.Text;
            this.DialogResult = DialogResult.OK;
        }
    }
}
