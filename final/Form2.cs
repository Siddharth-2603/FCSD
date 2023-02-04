using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


//public static string SetValueForText2 = "";
//public static string SetValueForText3 = "";

namespace final
{
    //public event EventHandler DataPassed;

    public partial class Form2 : Form
    {
        public static string SetValueForText1 = "";
        public Form2()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            //DataPassed ?.Invoke(this, EventArgs.Empty);
            SetValueForText1 = text1.Text;
            Form f3 = new Form3();
            f3.Show();
        }

        private void text1_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
