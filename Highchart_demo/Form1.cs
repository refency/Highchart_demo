using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Web.WebView2.WinForms;
using Aspose.Cells;
using System.Xml.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;


namespace Highchart_demo
{
    public partial class Form1 : Form
    {
        WebView2 WebBrowser_1 = new WebView2();
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            WebBrowser_1.Dock = DockStyle.Fill;
            this.Controls.Add(WebBrowser_1);
            WebBrowser_1.EnsureCoreWebView2Async();
            WebBrowser_1.CoreWebView2InitializationCompleted += WebBrowser_1_CoreWebView2InitializationCompleted;
        }

        private void ShowGraphic()
        {
            excel_reader();
        }

        private void CoreWebView2_DOMContentLoaded(object sender, Microsoft.Web.WebView2.Core.CoreWebView2DOMContentLoadedEventArgs e)
        {
            ShowGraphic();
        }

        private void excel_reader() {
            // Load Excel file
            Workbook wb = new Workbook(@"./parameters.xlsx");

            // Get all worksheets
            WorksheetCollection collection = wb.Worksheets;

            int code = 0;
            string name = "";
            string script = "";

            JArray all_data = new JArray();

            // Loop through all the worksheets
            for (int worksheetIndex = 0; worksheetIndex < collection.Count; worksheetIndex++)
            {
                // Get worksheet using its index
                Worksheet worksheet = collection[worksheetIndex];

                // Print worksheet name
                Console.WriteLine("Worksheet: " + worksheet.Name);

                // Get number of rows and columns
                int rows = worksheet.Cells.MaxDataRow;
                int cols = worksheet.Cells.MaxDataColumn;

                // Loop through rows
                for (int i = 1; i <= rows; i++) // Начинаю перебор со второго значения, поскольку в первом лежит имя столбца
                {
                    // Сначала приводим европейский вид времени с запятой, к другому, потому что не конвертится
                    // date = worksheet.Cells[i, 0].Value.ToString().Replace(",",".");
                    // Затем парсим эту строку в datetime
                    // DateTime.Parse(date)
                    // Для оптимизации, не использовал переменные, а сделал все разом
                    int code_of_value = Convert.ToInt32(worksheet.Cells[i, 1].Value);
                    long milliseconds = (long)(DateTime.Parse(worksheet.Cells[i, 0].Value.ToString().Replace(",", ".")).ToUniversalTime().Subtract(
                        new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc)
                        ).TotalMilliseconds);

                    double value = Convert.ToDouble(worksheet.Cells[i, 2].Value);
                    
                    // Итерация берется на 2 шага меньше, от значения реального файла
                    // То есть берем: i = 10000, а значение будет в file[10002]
                    // Все потому что в первом значении лежат названия столбцов, еще минус один,
                    // поскольку отсчет в массивах начинается с нуля

                    if (code != code_of_value || i == rows - 1) // Здесь херовая проверка, для добавления последнего элемента
                    {
                        switch (code) { // Временный вариант для наименования кодов
                            case 0:
                                name = "T пр(°C)| #0";
                                break;
                            case 1:
                                name = "T обр(°C)| #1";
                                break;
                            case 3:
                                name = "G пр(т)| #3";
                                break;
                            case 4:
                                name = "G обр(т)| #4";
                                break;
                            case 11:
                                name = "W пр(ГКал)| #11";
                                break;
                        }

                        code = code_of_value;

                        script = "elevationData.push(" + 
                            new JObject(
                            new JProperty("name", name),
                            new JProperty("data", all_data)
                        ) + ")";

                        WebBrowser_1.ExecuteScriptAsync(script);

                        all_data.Clear();
                    }

                    code = code_of_value;

                    all_data.Add(new JArray(
                        milliseconds,
                        value
                    ));
                }
            }

            script = @"for (let item of elevationData) {
                        chart.addSeries(item);
                    };";
            WebBrowser_1.ExecuteScriptAsync(script);
        }

        private void WebBrowser_1_CoreWebView2InitializationCompleted(object sender, Microsoft.Web.WebView2.Core.CoreWebView2InitializationCompletedEventArgs e)
        {
            // Это нужно будет в самом конце, для отключения функций браузера
            //WebBrowser_1.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;
            //WebBrowser_1.CoreWebView2.Settings.AreDevToolsEnabled = false;

            WebBrowser_1.CoreWebView2.DOMContentLoaded += CoreWebView2_DOMContentLoaded;

            string html_page = "./index.html";
            string HTML_String = "";

            using (StreamReader reader = new StreamReader(html_page))
            {
                HTML_String = reader.ReadToEnd();
            }

            WebBrowser_1.CoreWebView2.SetVirtualHostNameToFolderMapping("library", "./libraryes", Microsoft.Web.WebView2.Core.CoreWebView2HostResourceAccessKind.DenyCors);
            WebBrowser_1.CoreWebView2.NavigateToString(HTML_String);
        }
    }
}
