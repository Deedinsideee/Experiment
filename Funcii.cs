
using ExcelDataReader;
using Experiments;

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace DIPLOM
{
    public static class Funcii
    {

        public static void ImportEx(Microsoft.Win32.OpenFileDialog dialog, ExperimentsContext context)
        {
            if (dialog.ShowDialog() == true)
            {

                FileStream stream = File.Open(dialog.FileName, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                DataSet result = excelReader.AsDataSet();
                DataTable dt = result.Tables[0];

                var rows = dt.Rows;
                if (rows.Count >= 2)
                {
                    var header1 = rows[0];

                    if (header1[0].ToString() == "Time" &&
                        header1[1].ToString() == "Fz")
                    {
                        var header2 = rows[1];
                        if (header2[0].ToString() == "s" &&
                            header2[1].ToString() == "N")
                        {
                            var experiment = new Experiment
                            {
                                Name = Path.GetFileNameWithoutExtension(dialog.FileName),
                                Time = DateTime.Now
                            };

                            context?.Experiments?.Add(experiment);
                            context?.SaveChanges();

                            var data = new Data[rows.Count - 2];
                            for (int i = 2; i < rows.Count; i++)
                            {
                                if (double.TryParse(rows[i][0].ToString(), out double time) &&
                                    double.TryParse(rows[i][1].ToString(), out double fz))
                                {
                                    data[i - 2] = new Data()
                                    {
                                        Time = time,
                                        Fz = fz,
                                        ExperementId = experiment.Id
                                    };
                                }
                            }

                            context?.Data?.AddRange(data);
                            context?.SaveChanges();

                        }
                    }
                    else
                        MessageBox.Show("Некорректный файл");
                }

            }
        }


        public static void ImportEx2(string s, ExperimentsContext context)
        {

            FileStream stream = File.Open(s, FileMode.Open, FileAccess.Read);
            IExcelDataReader excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
            DataSet result = excelReader.AsDataSet();
            DataTable dt = result.Tables[0];

            var rows = dt.Rows;
            if (rows.Count >= 2)
            {
                var header1 = rows[0];

                if (header1[0].ToString() == "Time" &&
                    header1[1].ToString() == "Fz")
                {
                    var header2 = rows[1];
                    if (header2[0].ToString() == "s" &&
                        header2[1].ToString() == "N")
                    {
                        var experiment = new Experiment
                        {
                            Name = Path.GetFileNameWithoutExtension(s),
                            Time = DateTime.Now
                        };

                        context?.Experiments?.Add(experiment);
                        context?.SaveChanges();

                        var data = new Data[rows.Count - 2];
                        for (int i = 2; i < rows.Count; i++)
                        {
                            if (double.TryParse(rows[i][0].ToString(), out double time) &&
                                double.TryParse(rows[i][1].ToString(), out double fz))
                            {
                                data[i - 2] = new Data()
                                {
                                    Time = time,
                                    Fz = fz,
                                    ExperementId = experiment.Id
                                };
                            }
                        }

                        context?.Data?.AddRange(data);
                        context?.SaveChanges();

                    }

                }
            }
        } 
            
    }
}
