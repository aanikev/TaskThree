using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskThree
{
    internal class GoldCustomer
    {
        string _date = null;
        string _excelPath = null;
        public GoldCustomer(string date, string excelPath)
        {
            _date = date;
            _excelPath = excelPath;
        }
        public string GetGoldCustomer()
        {
            var excelDocument = new GetExcelDocument(_excelPath);
            if (_date.Length == 7)
            {
                string month = null;
                if (_date.Substring(0, 1) == "0")
                     month = _date.Substring(1, 1);
                else
                    month = _date.Substring(0, 2);

                string year = _date.Substring(3, 4);
                var PostingDateMonth = excelDocument.GetIncidents()
                .Select(x => x.PostingDate.Month.ToString());

                var goldCustomerCode = excelDocument.GetIncidents()
                .Where(x => x.PostingDate.Month.ToString() == month && x.PostingDate.Year.ToString() == year).GroupBy(x => x.CustomerCode)
                .OrderByDescending(x => x.Count()).First();
                var goldCustomer = excelDocument.GetCustomers()
                    .Where(x => x.CustomerCode == goldCustomerCode.Key).FirstOrDefault();
                return $"Организация золотого криента: {goldCustomer.OrganizationName}\t " +
                   $"контактное лицо: {goldCustomer.ContactPerson}";
            }
            else
            {
                var goldCustomerCode = excelDocument.GetIncidents()
                .Where(x => x.PostingDate.Year.ToString() == _date).GroupBy(x => x.CustomerCode)
               .OrderByDescending(x => x.Count()).First();
                var goldCustomer = excelDocument.GetCustomers()
                    .Where(x => x.CustomerCode == goldCustomerCode.Key).FirstOrDefault();
                return $"Организация золотого криента: {goldCustomer.OrganizationName}\t " +
                   $"контактное лицо: {goldCustomer.ContactPerson}";
            }
        }
    }
}
