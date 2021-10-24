﻿using Google.Apis.Auth.OAuth2;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Microsoft.Office.Interop.Excel;
using org.mariuszgromada.math.mxparser;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using ZedGraph;
using Application = Microsoft.Office.Interop.Excel.Application;
using Color = System.Drawing.Color;

namespace appr
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
            zedGraphDesign();
        }

        private async void button1_Click(object sender, EventArgs e)
        {
            try
            {
                if (dataGridView1.Rows[0] == null || dataGridView1.Rows[1] == null)
                {
                    MessageBox.Show($"Нельзя построить график, если таблица пуста!", "Error!", MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }

                else
                {
                    GraphPane pane = zedGraphControl1.GraphPane;
                    pane.CurveList.Clear();
                    PointPairList list_points = new PointPairList();
                    List<double> values = new List<double>();

                    for (int i = 0; i < dataGridView1.Rows.Count; i++)
                    {
                        list_points.Add(Convert.ToDouble(dataGridView1[0, i].Value), Convert.ToDouble(dataGridView1[1, i].Value));
                        //ошибка неверный формат ввода если первый столбец Х пуст
                    }

                    LineItem line = pane.AddCurve("Точки", list_points, Color.Red, SymbolType.Circle);

                    line.Symbol.Fill.Color = Color.Red;
                    line.Symbol.Fill.Type = FillType.Solid;
                    line.Line.IsVisible = false;
                    line.Symbol.Size = 3;

                    zedGraphControl1.IsShowPointValues = true;

                    await buildasync();

                    pane.Title.Text = "Метод наименьших квадратов";
                    zedGraphControl1.AxisChange();
                    zedGraphControl1.Invalidate();
                    dataGridView1.AllowUserToAddRows = true;
                }
            }

            catch (Exception)
            {
                MessageBox.Show($"Нельзя построить график, если таблица пуста!", "Error!", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private void quadraticFunc()
        {
            double xi = 0;
            double yi = 0;
            double xy = 0;
            double x2 = 0;
            double x3 = 0;
            double x4 = 0;
            double x2y = 0;
            double yX = 0;
            double yXavg = 0;
            int n = dataGridView1.RowCount - 1;
            double avgY;
            double a, b, c;
            double det, deta, detb, detc;
            string exp = "";
            string quaddeterm = "";
            double determOut;
            double maximum = -1 * Math.Pow(2, 69);
            double minimum = Math.Pow(2, 69);
            PointPairList quadraticList = new PointPairList();

            for (int i = 0; i < n; i++)
            {
                xi += Convert.ToDouble(dataGridView1[0, i].Value);
                yi += Convert.ToDouble(dataGridView1[1, i].Value);
                xy += Convert.ToDouble(dataGridView1[0, i].Value) * Convert.ToDouble(dataGridView1[1, i].Value);
                x2 += Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value), 2);
                x3 += Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value), 3);
                x4 += Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value), 4);
                x2y += Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value), 2) * Convert.ToDouble(dataGridView1[1, i].Value);
            }

            avgY = yi / n;

            det = (x2 * x2 * x2) + (xi * xi * x4) + (x3 * x3 * n) - (n * x2 * x4) - (xi * x2 * x3) - (xi * x3 * x2);

            deta = (yi * x2 * x2) + (xi * xi * x2y) + (xy * x3 * n) - (x2y * x2 * n) - (xi * x2 * xy) - (yi * xi * x3);
            a = deta / det;

            detb = (x2 * x2 * xy) + (yi * xi * x4) + (x3 * x2y * n) - (x4 * xy * n) - (yi * x2 * x3) - (x2 * xi * x2y);
            b = detb / det;

            detc = (x2 * x2 * x2y) + (xi * xy * x4) + (x3 * x3 * yi) - (x4 * x2 * yi) - (xi * x2y * x3) - (x2 * xy * x3);
            c = detc / det;


            if (b >= 0)
            {
                exp += Convert.ToString(a.ToString("F" + 4)).Replace(',', '.') + "*x^2+" + Convert.ToString(b.ToString("F" + 4)).Replace(',', '.') + "*x";
                quaddeterm += Convert.ToString(Math.Round(a, 3)).Replace(',', '.') + "*x^2+" + Convert.ToString(Math.Round(b, 3)).Replace(',', '.') + "*x";

                if (c >= 0)
                {
                    exp += "+" + Convert.ToString(c.ToString("F" + 4)).Replace(',', '.');
                    quaddeterm += "+" + Convert.ToString(Math.Round(c, 3)).Replace(',', '.');
                }

                else
                {
                    exp += Convert.ToString(c.ToString("F" + 4)).Replace(',', '.');
                    quaddeterm += Convert.ToString(Math.Round(c, 3)).Replace(',', '.');
                }
            }

            else
            {
                exp = Convert.ToString(a.ToString("F" + 4)).Replace(',', '.') + "*x^2" + Convert.ToString(b.ToString("F" + 4)).Replace(',', '.') + "*x";
                quaddeterm = Convert.ToString(Math.Round(a, 3)).Replace(',', '.') + "*x^2" + Convert.ToString(Math.Round(b, 3)).Replace(',', '.') + "*x";
                if (c >= 0)
                {
                    exp += "+" + Convert.ToString(c.ToString("F" + 4)).Replace(',', '.');
                    quaddeterm += "+" + Convert.ToString(Math.Round(c, 3)).Replace(',', '.');
                }

                else
                {
                    exp += Convert.ToString(c.ToString("F" + 4)).Replace(',', '.');
                    quaddeterm += Convert.ToString(Math.Round(c, 3)).Replace(',', '.');
                }
            }
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (Convert.ToDouble(dataGridView1[0, i].Value) > maximum)
                {
                    maximum = Convert.ToDouble(dataGridView1[0, i].Value);
                }
                if (Convert.ToDouble(dataGridView1[0, i].Value) < minimum)
                {
                    minimum = Convert.ToDouble(dataGridView1[0, i].Value);
                }
            }

            for (int i = Convert.ToInt32(minimum); i <= maximum; i++)
            {

                quadraticList.Add(i, func(i, exp));
            }

            for (int i = 0; i < n; i++)
            {
                yX += Math.Pow(Convert.ToDouble(dataGridView1[1, i].Value) - func(Convert.ToDouble(dataGridView1[0, i].Value), exp), 2);
                yXavg += Math.Pow(Convert.ToDouble(dataGridView1[1, i].Value) - avgY, 2);
            }

            determOut = Math.Pow(Math.Sqrt(1 - yX / yXavg), 2);

            quadcorel.Invoke((MethodInvoker)delegate { quadcorel.Text = determOut.ToString("F" + 4); });
            quad.Invoke((MethodInvoker)delegate { quad.Text = quaddeterm; });

            Invoke((MethodInvoker)delegate { addPoints(quadraticList, "Квадратичная функция"); });
        }

        private async Task buildasync()
        {
            await Task.Run(() => linearFunc());
            await Task.Run(() => quadraticFunc());
        }

        private void linearFunc()
        {
            double xi = 0;
            double yi = 0;
            double xy = 0;
            double x2 = 0;
            double y2 = 0;
            int n = dataGridView1.RowCount - 1;
            double a, b;
            string exp;
            string lineardeterm;
            double maximum = -1 * Math.Pow(2, 69);
            double minimum = Math.Pow(2, 69);
            PointPairList linearList = new PointPairList();

            for (int i = 0; i < n; i++)
            {
                xi += Convert.ToDouble(dataGridView1[0, i].Value);
                yi += Convert.ToDouble(dataGridView1[1, i].Value);
                xy += Convert.ToDouble(dataGridView1[0, i].Value) * Convert.ToDouble(dataGridView1[1, i].Value);
                x2 += Math.Pow(Convert.ToDouble(dataGridView1[0, i].Value), 2);
                y2 += Math.Pow(Convert.ToDouble(dataGridView1[1, i].Value), 2);
            }

            a = ((xi * yi) - n * xy) / (Math.Pow(xi, 2) - n * x2);
            b = (xi * xy - x2 * yi) / (Math.Pow(xi, 2) - n * x2);

            double linC = ((n * xy) - (xi * yi)) / Math.Sqrt(((n * x2) - Math.Pow(xi, 2)) * ((n * y2) - Math.Pow(yi, 2)));
            //индекс детерминации
            double linD = Math.Pow(linC, 2);

            if (b >= 0)
            {
                exp = Convert.ToString(a).Replace(',', '.') + "*x+" + Convert.ToString(b).Replace(',', '.');
                lineardeterm = Convert.ToString(Math.Round(a, 3)).Replace(',', '.') + "*x+" + Convert.ToString(Math.Round(b, 3)).Replace(',', '.');
            }

            else
            {
                exp = Convert.ToString(a).Replace(',', '.') + "*x" + Convert.ToString(b).Replace(',', '.');
                lineardeterm = Convert.ToString(Math.Round(a, 3)).Replace(',', '.') + "*x" + Convert.ToString(Math.Round(b, 3)).Replace(',', '.');
            }
            for (int i = 0; i < dataGridView1.RowCount; i++)
            {
                if (Convert.ToDouble(dataGridView1[0, i].Value) > maximum)
                {
                    maximum = Convert.ToDouble(dataGridView1[0, i].Value);
                }
                if (Convert.ToDouble(dataGridView1[0, i].Value) < minimum)
                {
                    minimum = Convert.ToDouble(dataGridView1[0, i].Value);
                }
            }

            for (int i = Convert.ToInt32(minimum); i <= maximum; i++)
            {
                linearList.Add(i, func(i, exp));
            }

            linearcorel.Invoke((MethodInvoker)delegate { linearcorel.Text = linD.ToString("F" + 4); });
            linear.Invoke((MethodInvoker)delegate { linear.Text = lineardeterm; });

            Invoke((MethodInvoker)delegate { addPoints(linearList, "Линейная функция"); });
        }

        private double func(double x, string exp)
        {
            try
            {
                Argument xmain = new Argument("x");
                Expression y = new Expression(exp.ToLower(), xmain);

                xmain.setArgumentValue(x);
                return y.calculate();
            }

            catch (Exception)
            {
                return 0;
            }
        }

        private void addPoints(PointPairList l, string name)
        {
            GraphPane pane = zedGraphControl1.GraphPane;

            if (name == "Линейная функция")
            {
                LineItem linear = pane.AddCurve(name, l, Color.FromArgb(57, 198, 68), SymbolType.None);
            }
            else
            {
                LineItem quadro = pane.AddCurve(name, l, Color.FromArgb(44, 114, 205), SymbolType.None);
                quadro.Line.Width = 2.0F;
            }
            ((LineItem)zedGraphControl1.GraphPane.CurveList[1]).Line.Width = 2.0F;
            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            openFileDialog1.FileName = String.Empty;
            DialogResult res = openFileDialog1.ShowDialog();
            if (res != DialogResult.OK) return;

            try
            {
                dataGridView1.Rows.Clear();
                Application ObjWorkExcel = new Application();
                Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(openFileDialog1.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing); //открыть файл
                Worksheet ObjWorkSheet = (Worksheet)ObjWorkBook.Sheets[1];

                if (ObjWorkSheet.Rows.CurrentRegion.EntireRow.Count == 1)
                {
                    MessageBox.Show("No data found.");
                }

                else
                {
                    var lastCell = ObjWorkSheet.Cells.SpecialCells(XlCellType.xlCellTypeLastCell);

                    string sx = String.Empty;
                    string sy = String.Empty;

                    for (int i = 0; i < lastCell.Row; i++)
                    {
                        sx = ObjWorkSheet.Cells[i + 1, 1].Text.ToString();
                        sy = ObjWorkSheet.Cells[i + 1, 2].Text.ToString();

                        if (sx.Trim() != String.Empty && sy.Trim() != String.Empty)
                            dataGridView1.Rows.Add(sx, sy);
                    }
                    GC.Collect();
                }

                ObjWorkBook.Close(false, Type.Missing, Type.Missing);
                ObjWorkExcel.Quit();
            }

            catch (Exception exception)
            {
                MessageBox.Show($"При попытке загрузки из Excel произошла обшика!\n{exception.Message}", "Error!", MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        private static readonly string[] Scopes = { SheetsService.Scope.Spreadsheets };
        private const string SpreadsheetId = "1GoOUEb2OdQWLqPIJbEB_wMnfy4sJwc4cSkdMra2AEKM";
        private const string GoogleCredentialsFileName = "clients_secrets.json";
        private const string ReadRange = "sheet1!A:B";

        private async void googleDocsToolStripMenuItem_Click(object sender, EventArgs e)
        {
            await readAsync();
        }

        async Task readAsync()
        {
            try
            {
                var serviceValues = GetSheetsService().Spreadsheets.Values;
                await ReadAsync(serviceValues);
            }
            catch (Exception)
            {
                MessageBox.Show($"Количество элементов X != количеству элементам Y\nНекоторые элементы будут упущены", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private static SheetsService GetSheetsService()
        {
            using (var stream = new FileStream(GoogleCredentialsFileName, FileMode.Open, FileAccess.Read))
            {
                var serviceInitializer = new BaseClientService.Initializer
                {
                    HttpClientInitializer = GoogleCredential.FromStream(stream).CreateScoped(Scopes)
                };

                return new SheetsService(serviceInitializer);
            }
        }

        private async Task ReadAsync(SpreadsheetsResource.ValuesResource valuesResource)
        {
            var response = await valuesResource.Get(SpreadsheetId, ReadRange).ExecuteAsync();
            var values = response.Values;

            if (values == null || values.Count == 0)
            {
                MessageBox.Show("No data found.");
                return;
            }

            List<string> data = new List<string>();
            for (int i = 0; i < values.Count; i++)
            {
                string val0 = values[i][0].ToString();
                string val1 = values[i][1].ToString();

                if (val0 == String.Empty || val1 == String.Empty)
                {
                    MessageBox.Show("Количество элементов X != количеству элементам Y", "Ошибка");
                    return;
                }
                data.Add($"{val0};{val1}");

                if (data.Count == 0)
                {
                    MessageBox.Show("Количество элементов X != количеству элементам Y", "Ошибка");
                    return;
                }

                dataGridView1.Rows.Clear();

                foreach (string s in data)
                {
                    var res = s.Split(';');
                    dataGridView1.Rows.Add(res[0], res[1]);
                }
            }
        }
        private void очиститьToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();

            GraphPane pane = zedGraphControl1.GraphPane;
            zedGraphControl1.GraphPane.CurveList.Clear();
            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();

            quad.Text = "";
            linear.Text = "";
            linearcorel.Text = "";
            quadcorel.Text = "";
            textBox1.Text = "";

            pane.XAxis.Scale.MinAuto = true;
            pane.XAxis.Scale.MaxAuto = true;
            pane.YAxis.Scale.MinAuto = true;
            pane.YAxis.Scale.MaxAuto = true;
        }

        private void закрытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Close();
        }

        public const int WM_NCLBUTTONDOWN = 0xA1;
        public const int HT_CAPTION = 0x2;
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        public static extern bool ReleaseCapture();

        private void frmMain_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        private void menuStrip1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                ReleaseCapture();
                SendMessage(Handle, WM_NCLBUTTONDOWN, HT_CAPTION, 0);
            }
        }

        public void zedGraphDesign()
        {
            GraphPane graphfield = zedGraphControl1.GraphPane;
            graphfield.Border.Color = Color.Black;
            graphfield.Chart.Border.Color = Color.Black;
            graphfield.Fill.Type = FillType.Solid;
            graphfield.Fill.Color = Color.FromArgb(50, 49, 69);
            graphfield.Chart.Fill.Type = FillType.Solid;
            graphfield.Chart.Fill.Color = Color.Black;
            graphfield.XAxis.MajorGrid.IsZeroLine = true;
            graphfield.YAxis.MajorGrid.IsZeroLine = true;
            graphfield.XAxis.Color = Color.CornflowerBlue;
            graphfield.YAxis.Color = Color.CornflowerBlue;
            graphfield.XAxis.MajorGrid.IsVisible = true;
            graphfield.YAxis.MajorGrid.IsVisible = true;
            graphfield.XAxis.MajorGrid.Color = Color.FromArgb(62, 120, 138);
            graphfield.YAxis.MajorGrid.Color = Color.FromArgb(62, 120, 138);
            graphfield.XAxis.Title.FontSpec.FontColor = Color.Silver;
            graphfield.YAxis.Title.FontSpec.FontColor = Color.Silver;
            graphfield.XAxis.Scale.FontSpec.FontColor = Color.Silver;
            graphfield.YAxis.Scale.FontSpec.FontColor = Color.Silver;
            graphfield.Title.FontSpec.FontColor = Color.White;
        }

        private void textBox1_KeyPress(object sender, KeyPressEventArgs e)
        {
            char number = e.KeyChar;

            if (!Char.IsDigit(number) && number != 8)
            {
                e.Handled = true;
            }
        }

        //генератор через N
        private void button2_Click(object sender, EventArgs e)
        {
            dataGridView1.Rows.Clear();

            if (textBox1.Text == "" || textBox1.Text == "0" || textBox1.Text == "1" || textBox1.Text == "2")
            {
                MessageBox.Show("Введите количество строк N\nДля построения квадратичной фуникции должно быть не меньше 3 точек!");
                return;
            }

            if (double.Parse(textBox1.Text) >= 20000)
            {
                MessageBox.Show("При таком большом количестве точек программа будет работать медленнее", "Предупреждение");
            }

            button1.Enabled = false;
            dataGridView1.AllowUserToAddRows = false;
            int Yn = Convert.ToInt32(textBox1.Text);
            dataGridView1.Rows.Add(Yn);
            Random rnd = new Random();

            for (int row = 0; row < dataGridView1.Rows.Count; row++)
            {
                for (int col = 0; col < dataGridView1.Columns.Count; col++)
                {
                    dataGridView1[col, row].Value = rnd.Next(100);
                }
            }
            dataGridView1.AllowUserToAddRows = true;
            button1.Enabled = true;
        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            MessageBox.Show("1. Для правильного построения графика требуется не менее 3 точек в таблице\n2.Если много точек будут повторяться в значениях, расчеты могут быть неверны\n3.Количество элементов X должно быть равно количеству элементов Y");
        }

    }
}
