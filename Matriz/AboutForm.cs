using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Diagnostics;

namespace Matriz
{
    public partial class AboutForm : Form
    {
        public static string urlOfficial = "https://angeloeyez.github.io/Matriz-MatrixBOMTool/";
        public AboutForm()
        {
            InitializeComponent();
        }

        private void AboutForm_Load(object sender, EventArgs e)
        {
            LinkLabel.Link link = new LinkLabel.Link();
            link.LinkData = urlOfficial;
            linkLabel1.Links.Add(link);

            var version = System.Windows.Forms.Application.ProductVersion;
            LabelVersion.Text = string.Format("Ver: {0}", version);
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Process.Start(e.Link.LinkData as string);
        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(urlOfficial);
        }
    }
}
