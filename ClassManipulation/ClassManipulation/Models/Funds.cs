using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Vml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ClassManipulation.Models
{
    public class Fund
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }

        public Fund() { }

    }

    public class Value
    {
        public int Id { get; set; }

        public int FundId { get; set; }
        public Fund Fund { get; set; }
        public DateTime ValueDate { get; set; }
        public double ValueDouble { get; set; }

        public Value() { }
    }

    public class DataHelper
    {
        public static int IdCounter = 1;
        static Random random = new Random();

        public static List<Value> CreateFundValues(List<Fund> funds, int start, int finish)
        {
            List<Value> result = new List<Value>();
            if(funds.Count > 0)
            {
                DataHelper dataHelper = new DataHelper();
                var list = Enumerable.Range(start, finish).Select((v, index) => new Value
                {
                    Id = dataHelper.IncrementId(),
                    FundId = funds.Where(f => f.Id == index + 1).First().Id,
                    ValueDate = DateTime.Now,
                    ValueDouble = Math.Round(random.NextDouble() * random.Next(100000), 2),
                    Fund = funds.Where(f => f.Id == index + 1).First()
                }).ToList();

                return list;
            }

            return null;
        }
        public static List<Fund> CreateFunds(int start, int finish)
        {
            List<Fund> result = new List<Fund>();
            DataHelper helper = new DataHelper();

            var list = Enumerable.Range(start, finish).Select(f => new Fund
            {
                Id = helper.IncrementId(),
                Name = helper.CreateRandomName(),
                Description = helper.CreateRandomDescription()
            }).ToList();

            return list;
        }

        internal int IncrementId()
        {
            return IdCounter++;
        }

        internal string CreateRandomName()
        {
            string[] names = new string[]
            {
                "Paulo", "Pedro", "Alexandre", "Filipe", "Gabriel", "Andre", "Maria", "Carlos", "Mayara", "Flavio", "Diogo", "Janice", "Eva", "Isabelle", "Aoife", 
                "Thiago", "Gabriela", "Caoimhe", "Niamh", "Luiz", "Leandro", "Sandra", "Ana", "Raissa", "Beatriz", "Julia", "Vinicius", "Larissa"
            };

            int size = names.Length;

            int position = random.Next(size);

            return names[position];
        }

        internal string CreateRandomDescription()
        {
            string[] description = new string[]
            {
                "Mutual Fund", "Exchange-Rate Fund", "Bond", "Money Market Fund", "Stock Fund", "Bond Fund", "Close-end Fund", "Index Fund",
                "Unit trust Fund", "Fixed Income Fund" 
            };

            int size = description.Length;  
            int position = random.Next(size);

            return description[position];
        }

        public static void ExportToExcel(List<Value> fundValues)
        {
            XLWorkbook workbook = new XLWorkbook();
            DataSet ds = new DataSet("FundValue_DataSet");
            DataTable dt = new DataTable("FundValue_DataTable");

            string[] columns = new string[]
            {
                "fund_id", "fund_name", "fund_description", "value_date", "value_value"
            };

            dt.Columns.AddRange(columns.Select(c => new DataColumn(c)).ToArray());

            for (int i = 0; i < fundValues.Count; i++)
            {
                var testRow = new object[] 
                {
                    fundValues[i].Fund.Id, 
                    fundValues[i].Fund.Name, 
                    fundValues[i].Fund.Description, 
                    fundValues[i].ValueDate.ToString(), 
                    fundValues[i].ValueDouble.ToString() 
                };

                dt.Rows.Add(testRow);
            }


            ds.Tables.Add(dt);

            workbook.Worksheets.Add(ds);

            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string savePath = System.IO.Path.Combine(desktop + "test.xlsx");
            workbook.SaveAs(savePath, false);

            Console.WriteLine($"Printing in the following directory: {savePath}");

        }
    }
}
