using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordAddIn1
{
    internal static class WebBrowserUtils
    {
        internal static void SetText(WebBrowser webBrowser, string text)
        {
            webBrowser.DocumentText = @"
                        <!DOCTYPE html>
                        <html>
                        <head>
                            <style>
                                body {
                                    font-family: 'Segoe UI';
                                    font-size: 12pt;
                                    margin: 8px;
                                }
                                .green {
                                    color: #00B050;
                                    font-weight: bold;
                                }
                                .red {
                                    color: red;
                                    font-weight: normal;
                                }
                                .yellow {
                                    background-color: #FFC000;
                                    font-weight: normal;
                                }
                            </style>
                        </head>
                        <body>
                        " + text.Replace(Environment.NewLine.ToString(), "<br/>") + @"
                        </body>
                        </html>";

        }
    }
}
