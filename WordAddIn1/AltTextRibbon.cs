using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace WordAddIn1
{
    public partial class AltTextRibbon
    {
        private void AltTextRibbon_Load(object sender, RibbonUIEventArgs e)
        {
            toggleButtonHideShowAltTextPane.Checked = PreferenceStore.GetDisplayOnStartup();
        }

        private void toggleButtonHideShowAltTextPane_Click(object sender, RibbonControlEventArgs e)
        {
            var hwnd = Globals.ThisAddIn.Application.ActiveWindow.Hwnd;
            var customTaskPane = Globals.ThisAddIn.panes[hwnd];
            if (toggleButtonHideShowAltTextPane.Checked)
            {
                customTaskPane.Visible = true;
            }
            else
            {
                customTaskPane.Visible = false;
            }
        }

        private void checkBoxDisplayOnStartup_Click(object sender, RibbonControlEventArgs e)
        {
            PreferenceStore.SetDisplayOnStartup(checkBoxDisplayOnStartup.Checked);
        }
    }
}
