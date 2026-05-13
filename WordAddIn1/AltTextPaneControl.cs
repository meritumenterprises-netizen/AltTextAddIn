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
				ReadOnly = true,
				ScrollBars = ScrollBars.Vertical,
				WordWrap = true,
				Dock = DockStyle.Fill,
				Font = new Font("Segoe UI", 12),
			};

			Controls.Add(txtAltText);
		}

		public void SetAltText(string altText)
		{
			if (string.IsNullOrWhiteSpace(altText))
			{
				txtAltText.Text = "[Selected graphic has no Alt Text.]";
			}
			else
			{
				altText = altText.Replace("\\n", Environment.NewLine);
				altText = altText.Replace("\r\n", Environment.NewLine);
				altText = altText.Replace("\n", Environment.NewLine);

				txtAltText.Text = altText;
			}
		}
	}
}
