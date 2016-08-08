using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace outlook_duplicate_remover
{
    public partial class Progress : Form
    {
        public string ProgressMessage
        {
            set
            {
                message.Text = value;
                message.Update();
            }
        }
        public void UpdateProgressBar()
        {
            progressBar1.Value += 1;
        }
        public void ResetProgressBar (int maximumLength)
        {
            progressBar1.Maximum = maximumLength;
            progressBar1.Minimum = 0;
            progressBar1.Value = 0;
        }
        public Progress(int maximumLengthOfProgressBar)
        {
            InitializeComponent();
            ResetProgressBar(maximumLengthOfProgressBar);
        }
    }
}
