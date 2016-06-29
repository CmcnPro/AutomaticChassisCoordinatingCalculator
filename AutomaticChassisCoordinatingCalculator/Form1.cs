using System;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;

namespace AutomaticChassisCoordinatingCalculator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void cpuURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(cpuURLTextBox.Text);
            core.getProductName();
            core.getPrice();
            cpuNameTextBox.Text = core.ProductName;
            cpuPrice.Text = core.Price;
        }

        private void mbURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(mbURLTextBox.Text);
            core.getProductName();
            core.getPrice();
            mbNameTextBox.Text = core.ProductName;
            mbPrice.Text = core.Price;
        }

        private void ramURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(ramURLTextBox.Text);
            core.getProductName();
            core.getPrice();
            ramNameTextBox.Text = core.ProductName;
            ramPrice.Text = core.Price;
        }

        private void ssdURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(ssdURLTextBox.Text);
            core.getProductName();
            core.getPrice();
            ssdNameTextBox.Text = core.ProductName;
            ssdPrice.Text = core.Price;
        }

        private void hddURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(hddURLTextBox.Text);
            core.getProductName();
            core.getPrice();
            hddNameTextBox.Text = core.ProductName;
            hddPrice.Text = core.Price;
        }

        private void gpuURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(gpuURLTextBox.Text);
            core.getProductName();
            core.getPrice();
            gpuNameTextBox.Text = core.ProductName;
            gpuPrice.Text = core.Price;
        }

        private void psuURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(psuURLTextBox.Text);
            core.getProductName();
            core.getPrice();
            psuNameTextBox.Text = core.ProductName;
            psuPrice.Text = core.Price;
        }

        private void cpufanURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(cpufanURLTextBox.Text);
            core.getProductName();
            core.getPrice();
            cpufanNameTextBox.Text = core.ProductName;
            cpufanPrice.Text = core.Price;
        }

        private void caseURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(caseURLTextBox.Text);
            core.getProductName();
            core.getPrice();
            caseNameTextBox.Text = core.ProductName;
            casePrice.Text = core.Price;
        }

        private void sysfanURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(sysfanURLTextBox.Text);
            core.getProductName();
            core.getPrice();
            sysfanNameTextBox.AppendText(core.ProductName);
            sysfanPrice.Text = core.Price;
        }

        private void totalButton_Click(object sender, System.EventArgs e)
        {
            float temp = float.Parse(cpuPrice.Text) + float.Parse(mbPrice.Text) + float.Parse(ramPrice.Text) + float.Parse(ssdPrice.Text) + float.Parse(hddPrice.Text) + float.Parse(gpuPrice.Text) + float.Parse(psuPrice.Text) + float.Parse(cpufanPrice.Text) + float.Parse(casePrice.Text) + float.Parse(sysfanPrice.Text);
            total.Text= temp.ToString();
        }

        private void ecxelbutton_Click(object sender, EventArgs e)
        {
            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Workbook wbook = app.Workbooks.Add(Type.Missing);
            Worksheet worksheet = (Worksheet)wbook.Worksheets[1];
            Range range = null;
            worksheet.Cells[1, 2] = "Name";
            worksheet.Cells[1, 3] = "Price";
            worksheet.Cells[1, 4] = "URL";
            worksheet.Cells[2, 1] = "CPU";
            worksheet.Cells[2, 2] = cpuNameTextBox.Text;
            worksheet.Cells[2, 3] = cpuPrice.Text;
            worksheet.Cells[2, 4] = cpuURLTextBox.Text;
            worksheet.Cells[3, 1] = "MB";
            worksheet.Cells[3, 2] = mbNameTextBox.Text;
            worksheet.Cells[3, 3] = mbPrice.Text;
            worksheet.Cells[3, 4] = mbURLTextBox.Text;
            worksheet.Cells[4, 1] = "RAM";
            worksheet.Cells[4, 2] = ramNameTextBox.Text;
            worksheet.Cells[4, 3] = ramPrice.Text;
            worksheet.Cells[4, 4] = ramURLTextBox.Text;
            worksheet.Cells[5, 1] = "SSD";
            worksheet.Cells[5, 2] = ssdNameTextBox.Text;
            worksheet.Cells[5, 3] = ssdPrice.Text;
            worksheet.Cells[5, 4] = ssdURLTextBox.Text;
            worksheet.Cells[6, 1] = "HDD";
            worksheet.Cells[6, 2] = hddNameTextBox.Text;
            worksheet.Cells[6, 3] = hddPrice.Text;
            worksheet.Cells[6, 4] = hddURLTextBox.Text;
            worksheet.Cells[7, 1] = "GPU";
            worksheet.Cells[7, 2] = gpuNameTextBox.Text;
            worksheet.Cells[7, 3] = gpuPrice.Text;
            worksheet.Cells[7, 4] = gpuURLTextBox.Text;
            worksheet.Cells[8, 1] = "PSU";
            worksheet.Cells[8, 2] = psuNameTextBox.Text;
            worksheet.Cells[8, 3] = psuPrice.Text;
            worksheet.Cells[8, 4] = psuURLTextBox.Text;
            worksheet.Cells[9, 1] = "CPUFan";
            worksheet.Cells[9, 2] = cpufanNameTextBox.Text;
            worksheet.Cells[9, 3] = cpufanPrice.Text;
            worksheet.Cells[9, 4] = cpufanURLTextBox.Text;
            worksheet.Cells[10, 1] = "Case";
            worksheet.Cells[10, 2] = caseNameTextBox.Text;
            worksheet.Cells[10, 3] = casePrice.Text;
            worksheet.Cells[10, 4] = caseURLTextBox.Text;
            worksheet.Cells[11, 1] = "SysFan";
            worksheet.Cells[11, 2] = sysfanNameTextBox.Text;
            worksheet.Cells[11, 3] = sysfanPrice.Text;
            worksheet.Cells[11, 4] = sysfanURLTextBox.Text;
            worksheet.Cells[12, 3] = "=SUM(C2:C11)";
            range = worksheet.get_Range("A1", "D20");
            range.Columns.AutoFit();
            range.Font.Name = "Microsoft YaHei";
            worksheet.SaveAs("accc.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wbook.Close();
            app.Quit();
        }
    }
}
