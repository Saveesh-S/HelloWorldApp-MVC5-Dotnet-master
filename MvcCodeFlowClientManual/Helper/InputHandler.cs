using DocumentFormat.OpenXml.Office2010.Excel;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;
using MvcCodeFlowClientManual.Model;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Linq;
using System.Web;

namespace MvcCodeFlowClientManual.Helper
{
    public static class InputHandler
    {
        public static string fileName = ConfigurationManager.AppSettings["filelocation"];//invoice record
        public static string FIELD_ServiceType_Microsoft365BusinessVoice = "Microsoft 365 Business Voice (without calling plan) Adoption Promo";
        public static string FIELD_ServiceType_Microsoft365E5 = "Microsoft 365 E5";
        public static string FIELD_VendorType_Cyclefee = "Cycle fee";
        public static decimal MarkUp = 1.22m;
        public static List<InvoiceRecord> GetCustomerProducts()
        {
            List<InvoiceRecord> customerStructure = new List<InvoiceRecord>();
            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

                customerStructure = GetCellData(GetSheetData("Sheet1", spreadsheetDocument), spreadsheetDocument);
                spreadsheetDocument.Close();
            }

            return customerStructure;
        }

        #region getTaxinfo
        //Added by Saveesh
        public static List<MarkUp> getMarkupDetails()
        {
            string markuptaxFilename = ConfigurationManager.AppSettings["tax_mark_file_locn"];
            //List<InvoiceRecord> customerStructure = new List<InvoiceRecord>();

            List<MarkUp> lstMkup = new List<MarkUp>();

            using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(markuptaxFilename, false))
            {
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                lstMkup = getMarkupCellValue(GetSheetData("mkpSheet", spreadsheetDocument), spreadsheetDocument);
                spreadsheetDocument.Close();
            }
            return lstMkup;
        }

        public static List<MarkUp> getMarkupCellValue(Worksheet ws_mrkup, SpreadsheetDocument document)
        {
            List<MarkUp> all_lstMkup = new List<MarkUp>();

            SheetData sheetData = ws_mrkup.GetFirstChild<SheetData>();
            IEnumerable<Row> rows = sheetData.Descendants<Row>();
            //int counter = 0;
            
            int counter = 2;

            foreach (Row row in rows)
            {
                MarkUp markup = new MarkUp();

                markup.subscriptionModel = GetStringCellValue(document, sheetData, "A" + counter.ToString(), null);
                markup.customerName = GetStringCellValue(document, sheetData, "B" + counter.ToString(), null);
                if (string.IsNullOrEmpty( markup.customerName))
                {
                    counter++;
                    continue;
                }
                
                markup.skuName = GetStringCellValue(document, sheetData, "C" + counter.ToString(), null);
                markup.skuID = GetStringCellValue(document, sheetData, "D" + counter.ToString(), null);
                markup.taxModel = GetStringCellValue(document, sheetData, "E" + counter.ToString(), null);
                if (markup.skuName != null)
                {
                    markup.markup = int.Parse(GetStringCellValue(document, sheetData, "F" + counter.ToString(), null));
                }

                markup.address = GetStringCellValue(document, sheetData, "G" + counter.ToString(), null);
                all_lstMkup.Add(markup);
                counter++;
            }


            return all_lstMkup;
        }

        #endregion getTaxinfo

        #region ProcessingParentInvoice

        public static Worksheet GetSheetData(string sheetName, SpreadsheetDocument document)
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
            Sheet sheet1 = sheets.Where(x => x.Name == sheetName).FirstOrDefault();
            string relationshipId = sheet1.Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);

            return worksheetPart.Worksheet;
        }


        public static List<InvoiceRecord> GetCellData(Worksheet workSheet, SpreadsheetDocument document)
        {
            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
            IEnumerable<Row> rows = sheetData.Descendants<Row>();
            //int counter = 0;
            List<InvoiceRecord> allCustomerInvoices = new List<InvoiceRecord>();
            //Get all Markup details
            List<MarkUp> all_lstMkup = InputHandler.getMarkupDetails();
            int counter = 2;
            MarkUp item = new MarkUp();

            foreach (Row row in rows)
            {
                InvoiceRecord invoice = new InvoiceRecord();

                string customerName = string.Empty;
                string productType = string.Empty;
                List<CellData> items = new List<CellData>();
                decimal numberOfDates = 0;
                decimal invoiceId = 0;
                
                string Cin = GetStringCellValue(document, sheetData, "AN" + counter.ToString(), null); //Old Value - AM
                Decimal.TryParse(GetStringCellValue(document, sheetData, "W" + counter.ToString(), null), out decimal daysOfUsage); //Old Value - V

                if (Cin != "Y" || daysOfUsage == 0)
                {
                    counter++;
                    continue;
                }
                //added by Saveesh
                string SynnexSKU = GetStringCellValue(document, sheetData, "R" + counter.ToString(), null);
                string serviceType = GetStringCellValue(document, sheetData, "K" + counter.ToString(), null);

                item = all_lstMkup.Find(c => (c.skuID == SynnexSKU) && (c.skuName == serviceType));
                //item = all_lstMkup.Find(c => (c.skuID == "6825916") && (c.skuName == "O365 F3 1Y COMMIT PAID/MO."));


                string subscriptionStartDate = GetStringCellValue(document, sheetData, "U" + counter.ToString(), null);//Old value - T
                string subscriptionEndDate = GetStringCellValue(document, sheetData, "V" + counter.ToString(), null);//Old value - V
                string lineDescription = GetStringCellValue(document, sheetData, "L" + counter.ToString(), null) + ((!String.IsNullOrEmpty(subscriptionEndDate) && !String.IsNullOrEmpty(subscriptionStartDate)) ? " (" + subscriptionStartDate + "-" + subscriptionEndDate + ")" : "");
                string vendorChargeType = GetStringCellValue(document, sheetData, "Y" + counter.ToString(), null);//Old value - X
                

                Decimal.TryParse(GetStringCellValue(document, sheetData, "W" + counter.ToString(), null), out numberOfDates);//Old value - V
                Decimal.TryParse(GetStringCellValue(document, sheetData, "S" + counter.ToString(), null), out decimal quantity);//Old value - R


                SalesItemLineDetailRecord salesItemLineDetailRecord = new SalesItemLineDetailRecord();
                salesItemLineDetailRecord.Qty = quantity;

                ItemRecord itemRecord = new ItemRecord();
                                
                itemRecord.ServiceType = serviceType;
                itemRecord.VendorChargeType = vendorChargeType;
                itemRecord.synnexskuID = SynnexSKU;


                customerName = GetStringCellValue(document, sheetData, "G" + counter.ToString(), null);
                productType = GetStringCellValue(document, sheetData, "I" + counter.ToString(), null);

                itemRecord.Description = lineDescription;
                salesItemLineDetailRecord.ItemRecord = itemRecord;

                Decimal.TryParse(GetStringCellValue(document, sheetData, "AH" + counter.ToString(), null), out decimal price);//Old value - AG
                Decimal.TryParse(GetStringCellValue(document, sheetData, "Z" + counter.ToString(), null), out decimal resellerPrice);//Old value - Y
                price = decimal.Round(price, 2, MidpointRounding.AwayFromZero);
                resellerPrice = decimal.Round(resellerPrice, 2, MidpointRounding.AwayFromZero);

                if (item!=null)
                {
                    price = decimal.Round(resellerPrice * (item.markup) / 100, 2, MidpointRounding.AwayFromZero);
                }

                /*if (resellerPrice > price )
                {
                    //This section needs to be modified - Saveesh
                    //price = resellerPrice * MarkUp; //considering rate = 18%
                    if (item.markup > 0)
                    {
                        price = resellerPrice * (item.markup)/100;
                    }

                }*/

                itemRecord.Name = serviceType + "_" + price.ToString();
                itemRecord.UnitPrice = price;
                decimal amount = quantity * price;

                LineRecord lineRecord = new LineRecord();

                lineRecord.Description = lineDescription;
                lineRecord.Amount = amount;
                lineRecord.AmountSpecified = true;
                lineRecord.Name = customerName + "-" + productType + "_" + lineDescription;

                lineRecord.SalesItemLineDetailRecord = salesItemLineDetailRecord;
                lineRecord.AnyIntuitObject = salesItemLineDetailRecord;

                Decimal.TryParse(GetStringCellValue(document, sheetData, "E" + counter.ToString(), null), out invoiceId);

                if (counter == 1 || (!String.IsNullOrEmpty(productType) && !String.IsNullOrEmpty(customerName)))
                {
                    int index = allCustomerInvoices.FindIndex(t => t.ProductType.ToLower() == productType.ToLower() && t.CustomerName.ToLower() == customerName.ToLower());
                    if (index > -1)
                    {
                        if(allCustomerInvoices[index].InvoiceId == "0" && invoiceId != 0)
                        {
                            allCustomerInvoices[index].InvoiceId = Math.Truncate(invoiceId).ToString();
                        }
                        //check if there is an item with same description (Azure Subscription)
                        int similarLine = allCustomerInvoices[index].LineItems.FindIndex(l => l.Description == lineRecord.Description && l.Description.ToString().ToLower().StartsWith("azure subscription"));
                        bool duplicateLines = isDuplicated(allCustomerInvoices[index].LineItems, serviceType, vendorChargeType);
                        if (similarLine > -1)
                        {
                            decimal cummulativeUnitPrice = allCustomerInvoices[index].LineItems[similarLine].SalesItemLineDetailRecord.ItemRecord.UnitPrice;
                            decimal cummulativeAmount = allCustomerInvoices[index].LineItems[similarLine].Amount;
                            allCustomerInvoices[index].LineItems[similarLine].SalesItemLineDetailRecord.ItemRecord.UnitPrice = cummulativeUnitPrice + lineRecord.SalesItemLineDetailRecord.ItemRecord.UnitPrice;
                            allCustomerInvoices[index].LineItems[similarLine].Amount = cummulativeAmount + lineRecord.Amount;
                        }
                        else if (duplicateLines)
                        {
                            //skip
                        }
                        else
                        {
                            allCustomerInvoices[index].LineItems.Add(lineRecord);
                        }
                    }
                    else
                    {
                        InvoiceRecord clientInvoiceItem = new InvoiceRecord();
                        clientInvoiceItem.ProductType = productType;
                        clientInvoiceItem.CustomerName = customerName;
                        clientInvoiceItem.InvoiceId = Math.Truncate(invoiceId).ToString();
                        clientInvoiceItem.PONumber = GetStringCellValue(document, sheetData, "AQ" + counter.ToString(), null);//Old value - AP
                        clientInvoiceItem.TotalAmt = amount;
                        clientInvoiceItem.domain = GetStringCellValue(document, sheetData, "O" + counter.ToString(), null);
                        DateTime dateTime = DateTime.Now;
                        DateTime.TryParse(subscriptionEndDate, out dateTime);

                        if (dateTime != null)
                        {
                            clientInvoiceItem.DueDate = dateTime;

                        }
                        DateTime.TryParse(subscriptionStartDate, out dateTime);
                        if (dateTime != null)
                        {
                            clientInvoiceItem.TxnDate = dateTime;

                        }
                        clientInvoiceItem.CustomerMemo = "Invoice for Microsoft licenses for the month of " + getMonth(dateTime.Month, dateTime.Year) + " " + dateTime.Year.ToString();

                        //added by Saveesh - 12/18
                        clientInvoiceItem.synnexskuID = SynnexSKU;

                        CustomerRecord customerRecord = new CustomerRecord();
                        customerRecord.CustomerName = customerName;

                        string addressStr = GetStringCellValue(document, sheetData, "H" + counter.ToString(), null);
                        if (addressStr != null && !String.IsNullOrEmpty(addressStr))
                        {
                            Intuit.Ipp.Data.PhysicalAddress address = new Intuit.Ipp.Data.PhysicalAddress();
                            address.City = addressStr.Split(',')[0];
                            address.PostalCode = addressStr.Split(',')[2];
                            address.CountrySubDivisionCode = addressStr.Split(',')[1];
                            customerRecord.ServiceLocation = address;
                        }

                        AccountRecord accountRecord = new AccountRecord();
                        accountRecord.AccountName = GetStringCellValue(document, sheetData, "A" + counter.ToString(), null);
                        accountRecord.AccountNumber = GetStringCellValue(document, sheetData, "B" + counter.ToString(), null);

                        clientInvoiceItem.AccountRecord = accountRecord;
                        clientInvoiceItem.CustomerRecord = customerRecord;
                        clientInvoiceItem.LineItems = new List<LineRecord>();
                        clientInvoiceItem.LineItems.Add(lineRecord);

                        allCustomerInvoices.Add(clientInvoiceItem);
                    }
                }
                counter++;
            }

            return allCustomerInvoices;
        }

        public static bool isDuplicated(List<LineRecord> customerLines, string serviceType, string vendorChargeType)
        {
            if (serviceType != FIELD_ServiceType_Microsoft365E5 && serviceType != FIELD_ServiceType_Microsoft365BusinessVoice)
            {
                return false;
            } 
            else if (serviceType == FIELD_ServiceType_Microsoft365BusinessVoice && vendorChargeType != FIELD_VendorType_Cyclefee) 
            {
                return false;
            }
            else if(serviceType == FIELD_ServiceType_Microsoft365BusinessVoice)
            {
                var row = customerLines.Where(x => x.SalesItemLineDetailRecord.ItemRecord.ServiceType == serviceType && x.SalesItemLineDetailRecord.ItemRecord.VendorChargeType == vendorChargeType);
                return row.Any();
            }
            else
            {
                var row = customerLines.Where(x => x.SalesItemLineDetailRecord.ItemRecord.ServiceType == serviceType);
                return row.Any();
            }
        }

        public static string GetStringCellValue(SpreadsheetDocument document, SheetData sheetData, string cellReference1, Cell excelCell)
        {
            WorkbookPart workbookPart = document.WorkbookPart;
            string cellValue = string.Empty;
            if (cellReference1 != null || excelCell != null)
            {
                var t = sheetData.Descendants<Cell>();
                Cell cell = new Cell();
                if (cellReference1 != null)
                {
                    cell = sheetData.Descendants<Cell>().FirstOrDefault(p => p.CellReference == cellReference1);
                }
                else
                {
                    cell = excelCell;
                }

                if (cell != null && cell.InnerText != null && cell.InnerText.Length > 0)
                {
                    cellValue = cell.InnerText;
                    if (cell.DataType != null)
                    {
                        switch (cell.DataType.Value)
                        {
                            case CellValues.SharedString:
                                var stringTable = workbookPart.GetPartsOfType<SharedStringTablePart>()
                                    .FirstOrDefault();

                                if (stringTable != null)
                                {
                                    cellValue = stringTable.SharedStringTable
                                        .ElementAt(int.Parse(cellValue)).InnerText;
                                }
                                break;
                            case CellValues.InlineString:
                                if (cell.InlineString != null)
                                {
                                    cellValue = cell.InlineString.InnerText;
                                }
                                break;
                            case CellValues.Boolean:
                                switch (cellValue)
                                {
                                    case "0":
                                        cellValue = "FALSE";
                                        break;
                                    default:
                                        cellValue = "TRUE";
                                        break;
                                }
                                break;
                        }
                    }
                    else if (cell.CellValue != null)
                    {
                        cellValue = cell.CellValue.InnerText;
                    }
                }
            }
            return cellValue;
        }


        public static string getMonth(int month, int year)
        {
            DateTime date = new DateTime(year, month, 1);

            return date.ToString("MMMM");
        }

        #endregion

    }
}