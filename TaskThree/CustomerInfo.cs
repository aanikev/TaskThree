﻿using DocumentFormat.OpenXml.Bibliography;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using TaskThree.Entities;

namespace TaskThree
{
    internal class CustomerInfo
    {
        string _productName = null;
        string _excelPath = null;
        public CustomerInfo(string productName, string excelPath)
        {
            _productName = productName;
            _excelPath = excelPath;
        }
        public string GetCustomerInfo()
        {
            var excelDocument = new GetExcelDocument(_excelPath);

            var productList = excelDocument.GetProducts();
            var customerList = excelDocument.GetCustomers();
            var incidentList = excelDocument.GetIncidents();

            var product = productList.Where(x => x.Name == _productName).FirstOrDefault();
            var incidents = incidentList
                .Where(x => x.ProductCode == product.ProductCode)
                .Cast<Incident>().ToList();
            var IncidentCustomerList = incidentList
                .Where(x => x.ProductCode == product.ProductCode)
                .Join(customerList,
                i => i.CustomerCode,
                c => c.CustomerCode,
                (i, c) => new { c.OrganizationName, c.ContactPerson, i.RequiredQuantity, i.PostingDate });
            string result = "";
            foreach (var customerInfo in IncidentCustomerList)
            {
                result += $"Название организации:\t {customerInfo.OrganizationName}, \n" +
                          $"Контактное лицо:\t {customerInfo.ContactPerson}, \n" +
                          $"Количество товара:\t {customerInfo.RequiredQuantity} \n" +
                          $"Цена:\t\t\t {customerInfo.RequiredQuantity * product.ProductPricePerUnit}\n" +
                          $"Дата заказа:\t\t {customerInfo.PostingDate} \n\n";
            }
            return result;
        }
    }
}
