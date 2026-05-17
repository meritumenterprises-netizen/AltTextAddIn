using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordAddIn1
{
	public partial class AltTextPaneControl : UserControl
	{
		private TextBox txtAltText;
        private bool _updatingFromWord;

        public event EventHandler<string> AltTextChangedByUser;

        public AltTextPaneControl()
		{
			InitializeComponent();
			BuildUi();
		}

		private void BuildUi()
		{
			txtAltText = new TextBox
			{
				Multiline = true,
				ReadOnly = false,
				ScrollBars = ScrollBars.Vertical,
				WordWrap = true,
				Dock = DockStyle.Fill,
				Font = new Font("Segoe UI", 12),
			};

            txtAltText.TextChanged += TxtAltText_TextChanged;

            splitContainer1.Panel1.Controls.Add(txtAltText);
		}

        private void TxtAltText_TextChanged(object sender, EventArgs e)
        {
            if (_updatingFromWord)
                return;

            AltTextChangedByUser?.Invoke(this, txtAltText.Text);
        }

        public void SetAltText(string altText)
		{
            _updatingFromWord = true;

            try
            {
                if (string.IsNullOrWhiteSpace(altText))
                {
                    txtAltText.Text = "";
                    webBrowser1.DocumentText = "";
                }
                else
                {
                    altText = altText.Replace("\\n", Environment.NewLine);
                    altText = altText.Replace("\r\n", Environment.NewLine);
                    altText = altText.Replace("\n", Environment.NewLine);

                    txtAltText.Text = altText;
                    WebBrowserUtils.SetText(webBrowser1, txtAltText.Text);
                }
            }
            finally
            {
                _updatingFromWord = false;
            }
        }

        public void SetPlaceholder()
        {
            _updatingFromWord = true;

            try
            {
                txtAltText.Text = "[No graphic selected.]";
                webBrowser1.DocumentText = "";
            }
            finally
            {
                _updatingFromWord = false;
            }
        }
    }
}
