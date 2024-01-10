using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office.CustomUI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TaskThree
{
    internal class EditContactPerson
    {
        string _organizationName = null;
        string _excelpath = null;
        string _newContactPerson = null;
        public EditContactPerson(string organizationName,string newContactPerson, string path)
        {
            _organizationName = organizationName;
            _newContactPerson = newContactPerson;
            _excelpath = path;
        }
        public string EditContactInTable()
        {
            var excelDocument = new GetExcelDocument(_excelpath);

            var customer = excelDocument.GetCustomers()
                .Where(x => x.OrganizationName == _organizationName).FirstOrDefault();

            var workbook = new XLWorkbook(_excelpath);
            workbook.Worksheet(2).CellsUsed()
            .FirstOrDefault(x => x.Value.ToString() == customer.ContactPerson).Value = _newContactPerson;

            workbook.SaveAs(_excelpath);
            return "Контакт изменён";
        }
    }
}
