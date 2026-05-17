using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
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

            control.AltTextChangedByUser += Control_AltTextChangedByUser;

            var pane = this.CustomTaskPanes.Add(
                control,
                "Selected Graphic Alt Text",
                window);

            pane.Visible = true;
            pane.Width = 400;

            panes[hwnd] = pane;
            controls[hwnd] = control;
        }

        private void Control_AltTextChangedByUser(object sender, string newAltText)
        {
            try
            {
                Word.Selection selection = this.Application.Selection;

                if (selection == null)
                    return;

                if (selection.InlineShapes.Count > 0)
                {
                    Word.InlineShape inlineShape = selection.InlineShapes[1];
                    inlineShape.AlternativeText = newAltText;
                    (sender as AltTextPaneControl).webBrowser1.DocumentText = @"
                        <!DOCTYPE html>
                        <html>
                        <head>
                            <style>
                                body {
                                    font-family: 'Segoe UI';
                                    font-size: 12pt;
                                    margin: 8px;
                                }
                            </style>
                        </head>
                        <body>
                        " + newAltText + @"
                        </body>
                        </html>"; 
                    return;
                }

                try
                {
                    if (selection.ShapeRange != null && selection.ShapeRange.Count > 0)
                    {
                        Word.Shape shape = selection.ShapeRange[1];
                        shape.AlternativeText = newAltText;
                    }
                }
                catch
                {
                    // ShapeRange throws if selection is not a floating shape.
                }
            }
            catch
            {
                // Optional: log error here.
            }
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
                if (!string.IsNullOrEmpty(altText) && altText != "[No graphic selected.]")
                {
                    //Word.Shape shape = Sel.ShapeRange[0];
                    if (Sel.InlineShapes.Count > 0)
                    {
                        Word.InlineShape inlineShape =
                            Sel.InlineShapes[1];
                        ApplyGlow(inlineShape.Glow);
                    }
                }
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

        private static void ApplyGlow(Word.GlowFormat glow)
        {
            // Green: #00B050
            glow.Color.RGB = 0x00B050;

            // Size: 5 pt
            glow.Radius = 5f;

            // Transparency: 60%
            // Word interop uses 0.0 = opaque, 1.0 = fully transparent
            glow.Transparency = 0.60f;
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