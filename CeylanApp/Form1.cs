using CeylanApp.Model.InputModel;
using CeylanApp.Model.OutputModel;
using ExcelDataReader;
using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Windows.Forms;
namespace CeylanApp
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private string JsonXLS;


        private void btnOpen_Click(object sender, EventArgs e)
        {
            OpenFileDialog file = new OpenFileDialog();
            if (file.ShowDialog() == DialogResult.OK)
            {
                string fileExt = Path.GetExtension(file.FileName);
                if (fileExt.CompareTo(".xls") == 0 || fileExt.CompareTo(".xlsx") == 0)
                {
                    try
                    {
                        ReadExcel(file.FileName);
                        MessageBox.Show("Файл Успешно считан");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message.ToString());
                    }
                }
                else
                {
                    MessageBox.Show("Please choose .xls or .xlsx file only.", "Warning", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        private string ReadExcel(string path)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
            {
                using (IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream))
                {
                    DataSet result = reader.AsDataSet(new ExcelDataSetConfiguration()
                    {
                        ConfigureDataTable = (_) => new ExcelDataTableConfiguration() { UseHeaderRow = true }
                    });
                    DataTableCollection tableCollection = result.Tables;

                    DataTable dt = tableCollection[0];
                    var json = JsonConvert.SerializeObject(dt);
                    JsonXLS = json;
                    return json;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<InputModel> inputModels = JsonConvert.DeserializeObject<List<InputModel>>(JsonXLS);
            var a = GetUsers(inputModels);
            string json = JsonConvert.SerializeObject(a);

            // richTextBox1.Text = json;

            dataGridView1.DataSource = a;
        }

        private string ReturnSuperVisorName(string Name)
        {
            if (Name == "7012 Ismayilova Gulxanim" ||
                  Name == "7032 Talibova TUBU" ||
                  Name == "7021 Muxtarli Turqut" ||
                  Name == "7033 Gulmammedova GUNAY " ||
                  Name == "7029 Babayeva XUMAR" ||
                  Name == "7034 Shixaliyeva SITARE" ||
                  Name == "7030 Mikayilov ANVER")
            {
                return "Mehdi Isayev";
            }
            else if (Name == "7006 Nuriyeva Ziba" ||
               Name == "7016 Salimova Rahime" ||
               Name == "7036 Eyvazli AYNUR" ||
               Name == "7005 Seferov Ramin" ||
               Name == "7026 Pashayeva Saltanet")
            {
                return "Həsənov Sərxan";
            }
            else if (Name == "7501 Магомеднабиев Расим (Ln)" ||
               Name == "7503 Абдуллаев Эмин (Ln)" ||
               Name == "7504 Habibov Farid (Ln)" ||
               Name == "7505 Bunyatov ELNUR (Lnd)")
            {
                return "Акшин Кязимов";
            }
            else if (
               Name == "1110 Bashirov Imran" ||
               Name == "1115 Mammedov Rasim" ||
               Name == "1119 Aliyev Ramil" ||
               Name == "1128 Agakerimov UMUD" ||
               Name == "1130 Mammadov ANAR" ||
               Name == "1134 Tagiyev RASIM" ||
               Name == "1135 Aliyev RAHIB" ||
               Name == "1131 Mikayilov MIKAYIL")
            {
                return "Əhmədov Tural";
            }
            else if (
                 Name == "RAHAT Topdan 20 (Гянджа)" ||
                 Name == "20001 Hasanov Eshqin")
            {
                return "Təvakkül Hüseynov";
            }
            else if (Name == "Итог")
            {
                return "Итог";
            }
            else
            {
                return "DİGER";
            }

            /*else if (Name == "Итог")
            {
                return "";
            }
            else
            {
                return "DİGER";
            }*/
        }



        const string TargetX1 = "10000";

        private List<OutputModel> GetUsers(List<InputModel> inputModels)
        {
            List<OutputModel> outputModels = new List<OutputModel>();

            foreach (var item in inputModels)
            {
                var outputModel = new OutputModel
                {
                    SupervisorName = ReturnSuperVisorName(item.AgentName),
                    Name = item.AgentName,
                    Basmati = new FoodModels
                    {
                        Target = TargetX1,
                        Real = item.BasmatiRice,
                        RealProcent = GetRealProcent(TargetX1, item.BasmatiRice),
                        Fkms = "FKMS",
                        Trend = "Trend",
                        TrendProcent = "Trend%"
                    },
                    Bic = new FoodModels
                    {
                        Target = TargetX1,
                        Real = item.Bic,
                        RealProcent = GetRealProcent(TargetX1, item.Bic),
                        Fkms = "FKMS",
                        Trend = "Trend",
                        TrendProcent = "Trend%"
                    },
                    Garofalo_Pasta = new FoodModels
                    {
                        Target = TargetX1,
                        Real = item.GarofaloPasta,
                        RealProcent = GetRealProcent(TargetX1, item.GarofaloPasta),
                        Fkms = "FKMS",
                        Trend = "Trend",
                        TrendProcent = "Trend%"
                    },
                    Isbre = new FoodModels
                    {
                        Target = TargetX1,
                        Real = item.Isbre,
                        RealProcent = GetRealProcent(TargetX1, item.Isbre),
                        Fkms = "FKMS",
                        Trend = "Trend",
                        TrendProcent = "Trend%"
                    },
                    Stmichele = new FoodModels
                    {
                        Target = TargetX1,
                        Real = item.StmichelePecniye,
                        RealProcent = GetRealProcent(TargetX1, item.StmichelePecniye),
                        Fkms = "FKMS",
                        Trend = "Trend",
                        TrendProcent = "Trend%"
                    },
                    Inkitchen = new FoodModels
                    {
                        Target = TargetX1,
                        Real = item.Inkitchen,
                        RealProcent = GetRealProcent(TargetX1, item.Inkitchen),
                        Fkms = "FKMS",
                        Trend = "Trend",
                        TrendProcent = "Trend%"
                    },
                    Oyuncaqlar = new FoodModels
                    {
                        Target = TargetX1,
                        Real = item.Oyuncaqlar,
                        RealProcent = GetRealProcent(TargetX1, item.Oyuncaqlar),
                        Fkms = "FKMS",
                        Trend = "Trend",
                        TrendProcent = "Trend%"
                    },
                    DrKorner = new FoodModels
                    {
                        Target = TargetX1,
                        Real = item.DrKorner,
                        RealProcent = GetRealProcent(TargetX1, item.DrKorner),
                        Fkms = "FKMS",
                        Trend = "Trend",
                        TrendProcent = "Trend%"
                    },
                    GoodDay = new FoodModels
                    {
                        Target = TargetX1,
                        Real = item.GoodDayCoffe,
                        RealProcent = GetRealProcent(TargetX1, item.GoodDayCoffe),
                        Fkms = "FKMS",
                        Trend = "Trend",
                        TrendProcent = "Trend%"
                    },
                    Bahlsen = new FoodModels
                    {
                        Target = TargetX1,
                        Real = item.Bahlsen,
                        RealProcent = GetRealProcent(TargetX1, item.Bahlsen),
                        Fkms = "FKMS",
                        Trend = "Trend",
                        TrendProcent = "Trend%"
                    },
                    Lindt = new FoodModels
                    {
                        Target = TargetX1,
                        Real = item.Lindt,
                        RealProcent = GetRealProcent(TargetX1, item.Lindt),
                        Fkms = "FKMS",
                        Trend = "Trend",
                        TrendProcent = "Trend%"
                    },
                    Zeni = new FoodModels
                    {
                        Target = TargetX1,
                        Real = item.Zeni,
                        RealProcent = GetRealProcent(TargetX1, item.Zeni),
                        Fkms = "FKMS",
                        Trend = "Trend",
                        TrendProcent = "Trend%"
                    },
                    Total = new FoodModels
                    {
                        Target = "777777",
                        Real = "7777777",
                        RealProcent = "7777777",
                        Fkms = "77777777",
                        Trend = "7777777",
                        TrendProcent = "777777"
                    }
                };

                outputModels.Add(outputModel);
            }

            return outputModels.OrderBy(x=>x.SupervisorName).ToList();
        }

        private string GetRealProcent(string target, string real)
        {
            if(real == "0" || string.IsNullOrEmpty(real))
            {
                return "0";
            }

            var targetDecimal = decimal.Parse(target, CultureInfo.InvariantCulture);
            var realDecimal = decimal.Parse(real, CultureInfo.InvariantCulture);

            var realProcent = realDecimal * 100 / targetDecimal;

            return realProcent.ToString();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            using(SaveFileDialog saveFileDialog = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx"})
            {
                if(saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    SaveExcell(saveFileDialog.FileName);
                }
            }   
        }


        private void SaveExcell(string fileName)
        {
            List<InputModel> inputModels = JsonConvert.DeserializeObject<List<InputModel>>(JsonXLS);
            var a = GetUsers(inputModels);

            var fileInfo = new FileInfo(fileName);
            using (var package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet excelWorksheet = package.Workbook.Worksheets.Add("Test");
                excelWorksheet.Cells.LoadFromCollection<OutputModel>(a);
                package.Save();
            }
        }
    }
}