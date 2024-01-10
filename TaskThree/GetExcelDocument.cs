using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TaskThree.Entities;

namespace TaskThree
{
    internal class GetExcelDocument
    {
        string _path = null;
        public GetExcelDocument(string path)
        {
           _path = path; 
        }
        public List<Product> GetProducts()
        {
            var workbook = new XLWorkbook(_path);
            var products = workbook.Worksheet(1);
            var rows = products.Range($"A2:D{products.RowCount() - 1}").RowsUsed();
            var productsList = new List<Product>();

            foreach (var row in rows)
            {
                var product = new Product();
                product.ProductCode = (int)row.Cell(1).Value;
                product.Name = row.Cell(2).Value.ToString();
                product.Unit = row.Cell(3).Value.ToString();
                product.ProductPricePerUnit = (double)row.Cell(4).Value;
                productsList.Add(product);
            }
            return productsList;
        }
        public List<Customer> GetCustomers()
        {
            var workbook = new XLWorkbook(_path);
            var custoomers = workbook.Worksheet(2);
            var rows = custoomers.Range($"A2:D{custoomers.RowCount() - 1}").RowsUsed();
            var custoomersList = new List<Customer>();
            foreach (var row in rows)
            {
                var customer = new Customer();
                customer.CustomerCode = (int)row.Cell(1).Value;
                customer.OrganizationName = row.Cell(2).Value.ToString();
                customer.OrganizationAddress = row.Cell(3).Value.ToString();
                customer.ContactPerson = row.Cell(4).Value.ToString();
                custoomersList.Add(customer);
            }
            return custoomersList;
        }
        public List<Incident> GetIncidents()
        {
            var workbook = new XLWorkbook(_path);
            var incidents = workbook.Worksheet(3);
            var rows = incidents.Range($"A2:F{incidents.RowCount() - 1}").RowsUsed();
            var incidentsList = new List<Incident>();
            foreach (var row in rows)
            {
                var incident = new Incident();
                incident.IncidentCode = (int)row.Cell(1).Value;
                incident.ProductCode = (int)row.Cell(2).Value;
                incident.CustomerCode = (int)row.Cell(3).Value;
                incident.NumberIncident = (int)row.Cell(4).Value;
                incident.RequiredQuantity = (int)row.Cell(5).Value;
                incident.PostingDate = DateTime.ParseExact(row.Cell(6).Value.ToString(), "G", null);
                incidentsList.Add(incident);
            }
            return incidentsList;
        }
    }
}
