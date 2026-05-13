using System;
using System.Collections.Generic;
using Microsoft.Office.Tools;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private Dictionary<int, CustomTaskPane> panes =
            new Dictionary<int, CustomTaskPane>();

        private Dictionary<int, AltTextPaneControl> controls =
            new Dictionary<int, AltTextPaneControl>();

        private void ThisAddIn_Startup(object sender, EventArgs e)
        {
            this.Application.WindowActivate += Application_WindowActivate;
            this.Application.WindowSelectionChange += Application_WindowSelectionChange;

            // Create pane for already-open windows
            foreach (Word.Window window in this.Application.Windows)
            {
                EnsurePaneForWindow(window);
            }
        }

        private void Application_WindowActivate(
            Word.Document Doc,
            Word.Window Wn)
        {
            EnsurePaneForWindow(Wn);
        }

        private void EnsurePaneForWindow(Word.Window window)
        {
            int hwnd = window.Hwnd;

            if (panes.ContainsKey(hwnd))
                return;

            var control = new AltTextPaneControl();

            var pane = this.CustomTaskPanes.Add(
                control,
                "Selected Graphic Alt Text",
                window);

            pane.Visible = true;
            pane.Width = 400;

            panes[hwnd] = pane;
            controls[hwnd] = control;
        }

        private void Application_WindowSelectionChange(Word.Selection Sel)
        {
            try
            {
                Word.Window activeWindow = this.Application.ActiveWindow;

                if (activeWindow == null)
                    return;

                int hwnd = activeWindow.Hwnd;

                if (!controls.ContainsKey(hwnd))
                    return;

                string altText = GetSelectedGraphicAltText(Sel);

                controls[hwnd].SetAltText(altText);
            }
            catch (Exception ex)
            {
                try
                {
                    Word.Window activeWindow = this.Application.ActiveWindow;

                    if (activeWindow == null)
                        return;

                    int hwnd = activeWindow.Hwnd;

                    if (controls.ContainsKey(hwnd))
                    {
                        controls[hwnd].SetAltText(
                            "Error reading Alt Text:\r\n" + ex.Message);
                    }
                }
                catch
                {
                }
            }
        }

        private string GetSelectedGraphicAltText(Word.Selection selection)
        {
            if (selection == null)
                return "";

            if (selection.InlineShapes.Count > 0)
            {
                Word.InlineShape inlineShape =
                    selection.InlineShapes[1];

                return inlineShape.AlternativeText ?? "";
            }

            try
            {
                if (selection.ShapeRange != null &&
                    selection.ShapeRange.Count > 0)
                {
                    Word.Shape shape =
                        selection.ShapeRange[1];

                    return shape.AlternativeText ?? "";
                }
            }
            catch
            {
            }

            return "[No graphic selected.]";
        }

        private void ThisAddIn_Shutdown(object sender, EventArgs e)
        {
            this.Application.WindowActivate -= Application_WindowActivate;
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