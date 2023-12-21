using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace Antoshin_IP
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double a1, b1, c1, a2, b2, c2;

            if (!double.TryParse(textBox1.Text, out a1) ||
            !double.TryParse(textBox2.Text, out b1) ||
            !double.TryParse(textBox3.Text, out c1) ||
            !double.TryParse(textBox4.Text, out a2) ||
            !double.TryParse(textBox5.Text, out b2) ||
            !double.TryParse(textBox6.Text, out c2))
            {
                MessageBox.Show("Please enter valid coefficients for the lines.");
                return;
            }

            double angleGeneral = CalculateSharpAngleGeneral(a1, b1, c1, a2, b2, c2);
            double angleCoefficients = CalculateSharpAngleWithCoefficients(a1, b1, a2, b2);
            double angleCanonical = CalculateSharpAngleCanonical(a1, b1, c1, a2, b2, c2);
            label7.Text = $"Результат:\nУгол между линиями 1 = {angleGeneral}\nУгол между линиями 2 = {angleCoefficients}\nУгол между линиями 3 = {angleCanonical}";

            string details = $"Линия коэффициентов 1 (a1, b1, c1): {a1}, {b1}, {c1}\n" +
            $"Линия коэффициентов 2 (a2, b2, c2): {a2}, {b2}, {c2}\n" +
            $"Угол между линиями 1 (Общее уравнение): {angleGeneral} degrees\n" +
            $"Угол между линиями 2 (Уравнение с коэффициентами): {angleCoefficients} degrees\n" +
            $"Угол между линиями 3 (Каноническое уравнение): {angleCanonical} degrees";
        }

        private double CalculateSharpAngleGeneral(double a1, double b1, double c1, double a2, double b2, double c2)
        {
            double result = a1 + b1 + c1 + a2 + b2 + c2;
            return result;
        }

        private double CalculateSharpAngleWithCoefficients(double a1, double b1, double a2, double b2)
        {
            double result = a1 + b1 + a2 + b2;
            return result;
        }

        private double CalculateSharpAngleCanonical(double a1, double b1, double c1, double a2, double b2, double c2)
        {
            double result = a1 + b1 + c1 + a2 + b2 + c2;
            return result;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            double a1, b1, c1, a2, b2, c2;

            if (!double.TryParse(textBox1.Text, out a1) ||
            !double.TryParse(textBox2.Text, out b1) ||
            !double.TryParse(textBox3.Text, out c1) ||
            !double.TryParse(textBox4.Text, out a2) ||
            !double.TryParse(textBox5.Text, out b2) ||
            !double.TryParse(textBox6.Text, out c2))
            {
                MessageBox.Show("Please enter valid coefficients for the lines.");
                return;
            }

            double angleGeneral = CalculateSharpAngleGeneral(a1, b1, c1, a2, b2, c2);
            double angleCoefficients = CalculateSharpAngleWithCoefficients(a1, b1, a2, b2);
            double angleCanonical = CalculateSharpAngleCanonical(a1, b1, c1, a2, b2, c2);
            label7.Text = $"Результат:\nУгол между линиями 1 = {angleGeneral}\nУгол между линиями 2 = {angleCoefficients}\nУгол между линиями 3 = {angleCanonical}";

            string details = $"Линия коэффициентов 1 (a1, b1, c1): {a1}, {b1}, {c1}\n" +
            $"Линия коэффициентов 2 (a2, b2, c2): {a2}, {b2}, {c2}\n" +
            $"Угол между линиями 1 (Общее уравнение): {angleGeneral} degrees\n" +
            $"Угол между линиями 2 (Уравнение с коэффициентами): {angleCoefficients} degrees\n" +
            $"Угол между линиями 3 (Каноническое уравнение): {angleCanonical} degrees";
            Word.Application wordApp = new Word.Application();
            wordApp.Visible = true;
            Word.Document doc = wordApp.Documents.Add();
            Word.Range range = doc.Range();
            range.Text = "Angles between lines" + Environment.NewLine + details;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            double a1, b1, c1, a2, b2, c2;

            if (!double.TryParse(textBox1.Text, out a1) ||
            !double.TryParse(textBox2.Text, out b1) ||
            !double.TryParse(textBox3.Text, out c1) ||
            !double.TryParse(textBox4.Text, out a2) ||
            !double.TryParse(textBox5.Text, out b2) ||
            !double.TryParse(textBox6.Text, out c2))
            {
                MessageBox.Show("Please enter valid coefficients for the lines.");
                return;
            }

            double angleGeneral = CalculateSharpAngleGeneral(a1, b1, c1, a2, b2, c2);
            double angleCoefficients = CalculateSharpAngleWithCoefficients(a1, b1, a2, b2);
            double angleCanonical = CalculateSharpAngleCanonical(a1, b1, c1, a2, b2, c2);
            label7.Text = $"Результат:\nУгол между линиями 1 = {angleGeneral}\nУгол между линиями 2 = {angleCoefficients}\nУгол между линиями 3 = {angleCanonical}";

            string details = $"Линия коэффициентов 1 (a1, b1, c1): {a1}, {b1}, {c1}\n" +
            $"Линия коэффициентов 2 (a2, b2, c2): {a2}, {b2}, {c2}\n" +
            $"Угол между линиями 1 (Общее уравнение): {angleGeneral} degrees\n" +
            $"Угол между линиями 2 (Уравнение с коэффициентами): {angleCoefficients} degrees\n" +
            $"Угол между линиями 3 (Каноническое уравнение): {angleCanonical} degrees";
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;
            Excel.Workbook workbook = excelApp.Workbooks.Add();
            Excel.Worksheet sheet = workbook.ActiveSheet;
            sheet.Cells[1, 1] = "Angles between lines";
            sheet.Cells[2, 1] = "Coefficients line 1";
            sheet.Cells[2, 2] = $"{a1}, {b1}, {c1}";
            sheet.Cells[3, 1] = "Coefficients line 2";
            sheet.Cells[3, 2] = $"{a2}, {b2}, {c2}";
            sheet.Cells[5, 1] = "Details";
            sheet.Cells[6, 1] = details;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            double a1, b1, c1, a2, b2, c2;

            if (!double.TryParse(textBox1.Text, out a1) ||
            !double.TryParse(textBox2.Text, out b1) ||
            !double.TryParse(textBox3.Text, out c1) ||
            !double.TryParse(textBox4.Text, out a2) ||
            !double.TryParse(textBox5.Text, out b2) ||
            !double.TryParse(textBox6.Text, out c2))
            {
                MessageBox.Show("Please enter valid coefficients for the lines.");
                return;
            }

            double angleGeneral = CalculateSharpAngleGeneral(a1, b1, c1, a2, b2, c2);
            double angleCoefficients = CalculateSharpAngleWithCoefficients(a1, b1, a2, b2);
            double angleCanonical = CalculateSharpAngleCanonical(a1, b1, c1, a2, b2, c2);
            label7.Text = $"Результат:\nУгол между линиями 1 = {angleGeneral}\nУгол между линиями 2 = {angleCoefficients}\nУгол между линиями 3 = {angleCanonical}";

            string details = $"Линия коэффициентов 1 (a1, b1, c1): {a1}, {b1}, {c1}\n" +
            $"Линия коэффициентов 2 (a2, b2, c2): {a2}, {b2}, {c2}\n" +
            $"Угол между линиями 1 (Общее уравнение): {angleGeneral} degrees\n" +
            $"Угол между линиями 2 (Уравнение с коэффициентами): {angleCoefficients} degrees\n" +
            $"Угол между линиями 3 (Каноническое уравнение): {angleCanonical} degrees";
            string filePathPdf = "angles_calculation_results.pdf";
            using (FileStream fs = new FileStream(filePathPdf, FileMode.Create, FileAccess.Write, FileShare.None))
            {
                iTextSharp.text.Document pdfDoc = new iTextSharp.text.Document();
                PdfWriter.GetInstance(pdfDoc, fs);
                pdfDoc.Open();

                pdfDoc.Add(new Paragraph("Angles between lines"));
                pdfDoc.Add(new Paragraph(details));

                pdfDoc.Close();
            }
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo()
            {
                FileName = filePathPdf,
                UseShellExecute = true
            });
            MessageBox.Show("Calculation results saved in Excel, Word and PDF.");
        }
    }
}