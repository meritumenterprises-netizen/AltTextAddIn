using System;
using System.Windows.Forms;
using Microsoft.Office.Tools;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;


namespace WordAddIn1
{
	public partial class ThisAddIn
	{
		private CustomTaskPane altTextTaskPane;
		private AltTextPaneControl altTextPaneControl;

		private void ThisAddIn_Startup(object sender, EventArgs e)
		{
			altTextPaneControl = new AltTextPaneControl();

			altTextTaskPane = this.CustomTaskPanes.Add(
				altTextPaneControl,
				"Selected Graphic Alt Text"
			);

			altTextTaskPane.Visible = true;
			altTextTaskPane.Width = 400;

			this.Application.WindowSelectionChange += Application_WindowSelectionChange;
		}

		private void Application_WindowSelectionChange(Word.Selection Sel)
		{
			try
			{
				string altText = GetSelectedGraphicAltText(Sel);
				altTextPaneControl.SetAltText(altText);
			}
			catch (Exception ex)
			{
				altTextPaneControl.SetAltText("Error reading Alt Text:\r\n" + ex.Message);
			}
		}

		private string GetSelectedGraphicAltText(Word.Selection selection)
		{
			if (selection == null)
				return "";

			if (selection.InlineShapes.Count > 0)
			{
				Word.InlineShape inlineShape = selection.InlineShapes[1];
				return inlineShape.AlternativeText ?? "";
			}

			if (selection.ShapeRange != null && selection.ShapeRange.Count > 0)
			{
				Word.Shape shape = selection.ShapeRange[1];
				return shape.AlternativeText ?? "";
			}

			return "[No graphic selected.]";
		}

		private void ThisAddIn_Shutdown(object sender, EventArgs e)
		{
			this.Application.WindowSelectionChange -= Application_WindowSelectionChange;
		}

		#region VSTO generated code

		private void InternalStartup()
		{
			this.Startup += new EventHandler(ThisAddIn_Startup);
			this.Shutdown += new EventHandler(ThisAddIn_Shutdown);
		}

		#endregion
	}
}
