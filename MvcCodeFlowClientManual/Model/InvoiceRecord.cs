using Intuit.Ipp.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;

namespace MvcCodeFlowClientManual.Model
{
    public class InvoiceRecord
    {
        public string CustomerName { get; set; }
        public string ProductType { get; set; }
        public string InvoiceId { get; set; }
        public string domain { get; set; }
        public string PONumber { get; set; }
        public string CustomerMemo { get; set; }
        public DateTime DueDate { get; set; }
        public DateTime TxnDate { get; set; }
        public decimal TotalAmt { get; set; }
        public decimal Balance { get; set; }
        public bool EmailStatus { get; set; }
        public PhysicalAddress ShipFromAddr { get; set; }

        public CustomerRecord CustomerRecord { get; set; }
        public AccountRecord AccountRecord { get; set; }

        public List<LineRecord> LineItems { get; set; }

        //synnexskuID - added by Saveesh
        public string synnexskuID { get; set; }

        public string address { get; set; }

    }
    public class CustomerRecord
    {
        public string CustomerName { get; set; }
        public string CompanyName { get; set; }
        public PhysicalAddress ServiceLocation { get; set; }

    }
    public class AccountRecord
    {
        public string AccountName { get; set; }
        public string AccountNumber { get; set; }

    }

    public class LineRecord
    {
        public string Description { get; set; }
        public string Name { get; set; }
        public decimal Amount { get; set; }
        public bool AmountSpecified { get; set; }
        public SalesItemLineDetailRecord SalesItemLineDetailRecord { get; set; }
        public SalesItemLineDetailRecord AnyIntuitObject { get; set; }


    }
    public class SalesItemLineDetailRecord
    {
        public decimal Qty { get; set; }
        public string CompanyName { get; set; }
        public PhysicalAddress ServiceLocation { get; set; }

        public ItemRecord ItemRecord { get; set; }

    }
    public class ItemRecord
    {
        public decimal UnitPrice { get; set; }
        public string Name { get; set; }
        public string Description { get; set; }
        public string ServiceType { get; set; }
        public string VendorChargeType { get; set; }
        public bool Type { get; set; }
        public bool TypeSpecified { get; set; }
        public bool UnitPriceSpecified { get; set; }
        public AccountRecord IncomeAccountRef { get; set; }
        public AccountRecord ExpenseAccountRef { get; set; }
        //synnexskuID - added by Saveesh
        public string synnexskuID { get; set; }


    }



    public class CellData
    {
        public string Name { get; set; }
        public string Type { get; set; }
    }
    //Added by Saveesh
    public class MarkUp
    {
        public string subscriptionModel { get; set; }
        public string skuName { get; set; }
        public string skuID { get; set; }
        public string taxModel { get; set; }
        public string customerName { get; set; }
        public int markup { get; set; }  
        
        public string address { get; set;}
    }

}