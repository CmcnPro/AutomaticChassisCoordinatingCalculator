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
            core.getPrice();
            core.getBrandName(cpuURLTextBox.Text);
            cpuNameTextBox.Text = core.BrandName;
            core.getModel(cpuURLTextBox.Text);
            cpuModelTextBox.Text = core.Model;
            cpuPrice.Text = core.Price;
        }

        private void mbURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(mbURLTextBox.Text);
            core.getPrice();
            core.getBrandName(mbURLTextBox.Text);
            mbNameTextBox.Text = core.BrandName;
            core.getModel(mbModelTextBox.Text);
            mbModelTextBox.Text = core.Model;
            mbPrice.Text = core.Price;
        }

        private void ramURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(ramURLTextBox.Text);
            core.getPrice();
            core.getBrandName(ramURLTextBox.Text);
            ramNameTextBox.Text = core.BrandName;
            core.getModel(ramModelTextBox.Text);
            ramModelTextBox.Text = core.Model;
            ramPrice.Text = core.Price;
        }

        private void ssdURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(ssdURLTextBox.Text);
            core.getPrice();
            core.getBrandName(ssdURLTextBox.Text);
            ssdNameTextBox.Text = core.BrandName;
            core.getModel(ssdModelTextBox.Text);
            ssdModelTextBox.Text = core.Model;
            ssdPrice.Text = core.Price;
        }

        private void hddURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(hddURLTextBox.Text);
            core.getPrice();
            core.getBrandName(hddURLTextBox.Text);
            hddNameTextBox.Text = core.BrandName;
            core.getModel(hddModelTextBox.Text);
            hddModelTextBox.Text = core.Model;
            hddPrice.Text = core.Price;
        }

        private void gpuURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(gpuURLTextBox.Text);
            core.getPrice();
            core.getBrandName(gpuURLTextBox.Text);
            gpuNameTextBox.Text = core.BrandName;
            core.getModel(gpuModelTextBox.Text);
            gpuModelTextBox.Text = core.Model;
            gpuPrice.Text = core.Price;
        }

        private void psuURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(psuURLTextBox.Text);
            core.getPrice();
            core.getBrandName(psuURLTextBox.Text);
            psuNameTextBox.Text = core.BrandName;
            core.getModel(psuModelTextBox.Text);
            psuModelTextBox.Text = core.Model;
            psuPrice.Text = core.Price;
        }

        private void cpufanURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(cpufanURLTextBox.Text);
            core.getPrice();
            core.getBrandName(cpufanURLTextBox.Text);
            cpufanNameTextBox.Text = core.BrandName;
            core.getModel(cpufanModelTextBox.Text);
            cpufanModelTextBox.Text = core.Model;
            cpufanPrice.Text = core.Price;
        }

        private void caseURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(caseURLTextBox.Text);
            core.getPrice();
            core.getBrandName(caseURLTextBox.Text);
            caseNameTextBox.Text = core.BrandName;
            core.getModel(caseModelTextBox.Text);
            caseModelTextBox.Text = core.Model;
            casePrice.Text = core.Price;
        }

        private void sysfanURLTextBox_TextChanged(object sender, System.EventArgs e)
        {
            Core core = new Core();
            core.getResponse(sysfanURLTextBox.Text);
            core.getPrice();
            core.getBrandName(sysfanURLTextBox.Text);
            sysfanNameTextBox.Text = core.BrandName;
            core.getModel(sysfanModelTextBox.Text);
            sysfanModelTextBox.Text = core.Model;
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
            worksheet.Cells[1, 2] = "Brand";
            worksheet.Cells[1, 3] = "Model";
            worksheet.Cells[1, 4] = "Price";
            worksheet.Cells[1, 5] = "URL";
            worksheet.Cells[2, 1] = "CPU";
            worksheet.Cells[2, 2] = cpuNameTextBox.Text;
            worksheet.Cells[2, 3] = cpuModelTextBox.Text;
            worksheet.Cells[2, 4] = cpuPrice.Text;
            worksheet.Cells[2, 5] = cpuURLTextBox.Text;
            worksheet.Cells[3, 1] = "MB";
            worksheet.Cells[3, 2] = mbNameTextBox.Text;
            worksheet.Cells[3, 3] = mbModelTextBox.Text;
            worksheet.Cells[3, 4] = mbPrice.Text;
            worksheet.Cells[3, 5] = mbURLTextBox.Text;
            worksheet.Cells[4, 1] = "RAM";
            worksheet.Cells[4, 2] = ramNameTextBox.Text;
            worksheet.Cells[4, 3] = ramModelTextBox.Text;
            worksheet.Cells[4, 4] = ramPrice.Text;
            worksheet.Cells[4, 5] = ramURLTextBox.Text;
            worksheet.Cells[5, 1] = "SSD";
            worksheet.Cells[5, 2] = ssdNameTextBox.Text;
            worksheet.Cells[5, 3] = ssdModelTextBox.Text;
            worksheet.Cells[5, 4] = ssdPrice.Text;
            worksheet.Cells[5, 5] = ssdURLTextBox.Text;
            worksheet.Cells[6, 1] = "HDD";
            worksheet.Cells[6, 2] = hddNameTextBox.Text;
            worksheet.Cells[6, 3] = hddModelTextBox.Text;
            worksheet.Cells[6, 4] = hddPrice.Text;
            worksheet.Cells[6, 5] = hddURLTextBox.Text;
            worksheet.Cells[7, 1] = "GPU";
            worksheet.Cells[7, 2] = gpuNameTextBox.Text;
            worksheet.Cells[7, 3] = gpuModelTextBox.Text;
            worksheet.Cells[7, 4] = gpuPrice.Text;
            worksheet.Cells[7, 5] = gpuURLTextBox.Text;
            worksheet.Cells[8, 1] = "PSU";
            worksheet.Cells[8, 2] = psuNameTextBox.Text;
            worksheet.Cells[8, 3] = psuModelTextBox.Text;
            worksheet.Cells[8, 4] = psuPrice.Text;
            worksheet.Cells[8, 5] = psuURLTextBox.Text;
            worksheet.Cells[9, 1] = "CPUFan";
            worksheet.Cells[9, 2] = cpufanNameTextBox.Text;
            worksheet.Cells[9, 3] = cpufanModelTextBox.Text;
            worksheet.Cells[9, 4] = cpufanPrice.Text;
            worksheet.Cells[9, 5] = cpufanURLTextBox.Text;
            worksheet.Cells[10, 1] = "Case";
            worksheet.Cells[10, 2] = caseNameTextBox.Text;
            worksheet.Cells[10, 3] = caseModelTextBox.Text;
            worksheet.Cells[10, 4] = casePrice.Text;
            worksheet.Cells[10, 5] = caseURLTextBox.Text;
            worksheet.Cells[11, 1] = "SysFan";
            worksheet.Cells[11, 2] = sysfanNameTextBox.Text;
            worksheet.Cells[11, 3] = sysfanModelTextBox.Text;
            worksheet.Cells[11, 4] = sysfanPrice.Text;
            worksheet.Cells[11, 5] = sysfanURLTextBox.Text;
            worksheet.Cells[12, 4] = "=SUM(D2:D11)";
            range = worksheet.get_Range("A1", "E20");
            range.Columns.AutoFit();
            range.Font.Name = "Microsoft YaHei";
            worksheet.SaveAs("accc.xlsx", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            wbook.Close();
            app.Quit();
        }
    }
}
