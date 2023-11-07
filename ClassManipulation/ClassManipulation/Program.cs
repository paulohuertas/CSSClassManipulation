using ClassManipulation.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassManipulation
{
    internal class Program
    {
        static void Main(string[] args)
        {
            List<Fund> funds = DataHelper.CreateFunds(1, 100000);
            List<Value> fundValues = DataHelper.CreateFundValues(funds, 1, 100000);

            DateTime starttime = DateTime.Now;
            Console.WriteLine($"Started at {starttime}");

            Console.WriteLine($"A total of {funds.Count()} has been created");

            foreach (var value in fundValues)
            {
                Console.WriteLine($"{value.Id}, {value.FundId}, {value.ValueDate}, {value.ValueDouble}, {value.Fund.Id}, {value.Fund.Name}, {value.Fund.Description}");
            }

            DateTime finishTime = DateTime.Now;

            DataHelper.ExportToExcel(fundValues);

            Console.WriteLine($"Finished at {finishTime}");

            Console.ReadLine();
        }
    }
}
