using Intuit.Ipp.OAuth2PlatformClient;
using System;
using System.Security.Claims;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Configuration;
using System.Net;
using Intuit.Ipp.Core;
using Intuit.Ipp.Data;
using Intuit.Ipp.QueryFilter;
using Intuit.Ipp.Security;
using System.Linq;
using Newtonsoft.Json;
using System.Collections.Generic;
using Intuit.Ipp.DataService;
using Item = Intuit.Ipp.Data.Item;
using MvcCodeFlowClientManual.Model;
using MvcCodeFlowClientManual.Helper;

namespace MvcCodeFlowClientManual.Controllers
{
    public class AppController : Controller
    {
        public static string clientid = ConfigurationManager.AppSettings["clientid"];
        public static string clientsecret = ConfigurationManager.AppSettings["clientsecret"];
        public static string redirectUrl = ConfigurationManager.AppSettings["redirectUrl"];
        public static string environment = ConfigurationManager.AppSettings["appEnvironment"];

        public static OAuth2Client auth2Client = new OAuth2Client(clientid, clientsecret, redirectUrl, environment);

        /// <summary>
        /// Use the Index page of App controller to get all endpoints from discovery url
        /// </summary>
        public ActionResult Index()
        {
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            Session.Clear();
            Session.Abandon();
            Request.GetOwinContext().Authentication.SignOut("Cookies");
            return View();
        }

        /// <summary>
        /// Start Auth flow
        /// </summary>
        public ActionResult InitiateAuth(string submitButton)
        {
            switch (submitButton)
            {
                case "Connect to QuickBooks":
                    List<OidcScopes> scopes = new List<OidcScopes>();
                    scopes.Add(OidcScopes.Accounting);
                    string authorizeUrl = auth2Client.GetAuthorizationURL(scopes);
                    return Redirect(authorizeUrl);
                default:
                    return (View());
            }
        }

        /// <summary>
        /// QBO API Request
        /// </summary>
        public ActionResult ApiCallService()
        {
            List<InvoiceRecord> customerStructure = InputHandler.GetCustomerProducts();
            List<MarkUp> all_lstMkup = InputHandler.getMarkupDetails();

            MarkUp mrkp_item = new MarkUp();


            if (Session["realmId"] != null)
            {
                string realmId = Session["realmId"].ToString();
                try
                {
                    var principal = User as ClaimsPrincipal;
                    OAuth2RequestValidator oauthValidator = new OAuth2RequestValidator(principal.FindFirst("access_token").Value);

                    // Create a ServiceContext with Auth tokens and realmId
                    ServiceContext serviceContext = new ServiceContext(realmId, IntuitServicesType.QBO, oauthValidator);
                    serviceContext.IppConfiguration.MinorVersion.Qbo = "65";

                    DataService dataService = new DataService(serviceContext);
                    
                    List<CellData> headerData = new List<CellData>();

                    if (customerStructure.Count > 0)
                    {
                        QueryService<Preferences> prefService = new QueryService<Preferences>(serviceContext);
                        Preferences pref = prefService.ExecuteIdsQuery("SELECT * FROM Preferences").First();
                        List<CustomFieldDefinition> customFieldDefinitions = pref.SalesFormsPrefs.CustomField.ToList();

                        Account accountAdded = null;
                        //check if account is already created
                        QueryService<Account> querySvc = new QueryService<Account>(serviceContext);
                        Account accountInfo = querySvc.ExecuteIdsQuery("SELECT * FROM Account WHERE Name = '" + customerStructure[0].AccountRecord.AccountName + "'").FirstOrDefault();

                        if (accountInfo == null)
                        {
                            //Create Account 
                            Account account = CreateAccount(customerStructure[0].AccountRecord);
                            accountAdded = dataService.Add<Account>(account);
                        }
                        else
                        {
                            accountAdded = dataService.Add<Account>(accountInfo);
                        }
                        List<CellData> items = new List<CellData>();

                        foreach (InvoiceRecord custProdItem in customerStructure)
                        {
                            Customer customerCreated = null;

                            //item = all_lstMkup.Find(c => (c.skuID == SynnexSKU) && (c.skuName == serviceType));
                            mrkp_item = all_lstMkup.Find(c => (c.skuID == custProdItem.synnexskuID) && (c.customerName == custProdItem.CustomerName));
                            //mrkp_item = all_lstMkup.Find(c => (c.skuID == custProdItem.synnexskuID)) ;//&& (c.skuName == custProdItem.CustomerName));
                            //mrkp_item = all_lstMkup.Find(c => (c.customerName == custProdItem.CustomerName));

                            //added by Saveesh -- 12/19
                            string physicalAddress = "";

                            if (mrkp_item != null) { 

                                physicalAddress = mrkp_item.address;
                            }

                            //check if customer is already created
                            QueryService<Customer> customerQuery = new QueryService<Customer>(serviceContext);
                            Customer customerInfo = customerQuery.ExecuteIdsQuery("SELECT * FROM Customer where DisplayName = '" + custProdItem.CustomerName + "'").FirstOrDefault();

                            Intuit.Ipp.Data.PhysicalAddress _address = new Intuit.Ipp.Data.PhysicalAddress();
                            _address.City = physicalAddress;/*"New West";
                            _address.CountrySubDivisionCode = "BC"; ;
                            _address.PostalCode = "V3M 3A8";
                            _address.Id = "331";*/

                            customerInfo.BillAddr = _address;

                            if (customerInfo != null)
                            {
                                customerCreated = dataService.Add<Customer>(customerInfo);
                            }
                            else
                            {
                                //Add customer
                                Customer customer = CreateCustomer(custProdItem.CustomerRecord);
                                customerCreated = dataService.Add<Customer>(customer);
                            }

                            QueryService<Invoice> invoiceQuery = new QueryService<Invoice>(serviceContext);
                            
                            


                            Invoice objInvoice = CreateInvoice(customerCreated, accountAdded, serviceContext, dataService, custProdItem);
                            Invoice addedInvoice = dataService.Add<Invoice>(objInvoice);
                            //Needs to be modified - Saveesh 12/18
                            //invoice.ShipFromAddr =   invoiceRecord.ShipFromAddr;

                            
                            
                            
                            addedInvoice.BillAddr = _address;
                            addedInvoice.ShipAddr = _address;

                            // Email invoice
                            // sending invoice 
                            dataService.SendEmail<Invoice>(addedInvoice, "saveesh.pillai@logicv.com");
                            //dataService.SendEmail<Invoice>(addedInvoice, "hemang@logicv.com");
                        }

                    }

                    return View("ApiCallService", (object)("QBO API call Successful!! Response: "));
                }
                catch (Exception ex)
                {
                    return View("ApiCallService", (object)("QBO API call Failed!" + " Error message: " + ex.Message));
                }
            }
            else
                return View("ApiCallService", (object)"QBO API call Failed!");
        }

        #region create account
        private Account CreateAccount(AccountRecord accountRecord)
        {

            Account account = new Account();

            account.Name = accountRecord.AccountName;

            account.FullyQualifiedName = accountRecord.AccountName;

            account.Classification = AccountClassificationEnum.Revenue;
            account.ClassificationSpecified = true;
            account.AccountType = AccountTypeEnum.Income;
            account.AccountTypeSpecified = true;
            account.AcctNum = accountRecord.AccountNumber;

            account.CurrencyRef = new ReferenceType()
            {
                name = "Canadian Dollar",
                Value = "CAD"
            };

            return account;
        }
        #endregion

        #region create item
        /// <summary>
        /// This API creates invoice item 
        /// </summary>
        /// <returns></returns>
        private Item CreateItem(Account incomeAccount, ItemRecord itemRecord, TaxCode taxCodeGST)
        {

            Item item = new Item();
            
            item.Name = itemRecord.Name;
            item.Description = itemRecord.Description;
            item.Type = ItemTypeEnum.NonInventory;
            item.TypeSpecified = true;

            item.Active = true;
            item.ActiveSpecified = true;

            item.Taxable = true;
            item.TaxableSpecified = true;

            item.SalesTaxCodeRef = new ReferenceType()
            {
                Value = taxCodeGST.Id
                //Value = null
            };

            item.UnitPrice = itemRecord.UnitPrice;
            item.UnitPriceSpecified = true;

            //item.RatePercent = itemRecord.UnitPrice;
            //item.RatePercentSpecified = true;

            item.TrackQtyOnHand = false;
            item.TrackQtyOnHandSpecified = false;            

            item.IncomeAccountRef = new ReferenceType()
            {
                name = incomeAccount.Name,
                Value = incomeAccount.Id
            };

            item.ExpenseAccountRef = new ReferenceType()
            {
                name = incomeAccount.Name,
                Value = incomeAccount.Id
            };

            //For inventory item, assetacocunref is required
            return item;

        }
        #endregion

        #region create customer
        /// <summary>
        /// This API creates customer 
        /// </summary>
        /// <returns></returns>
        private Customer CreateCustomer(CustomerRecord customerRecord)
        {
            Customer customer = new Customer();

            customer.DisplayName = customerRecord.CustomerName;
            customer.CompanyName = customerRecord.CompanyName;
            customer.BillAddr = customerRecord.ServiceLocation;

            return customer;
        }


        /// <summary>
        /// This API creates an Invoice
        /// </summary>
        
        private Invoice CreateInvoice(Customer customer, Account accountAdded, ServiceContext serviceContext, DataService dataService, InvoiceRecord invoiceRecord)
        {

            Invoice invoice = new Invoice();          
            
            invoice.CustomerRef = new ReferenceType()
            {
                Value = customer.Id
            };

            List<Line> lineList = new List<Line>();
            invoiceRecord.LineItems.Sort((x, y) => x.Description.CompareTo(y.Description));

            foreach (LineRecord lineRecord in invoiceRecord.LineItems)
            {
                Line line = new Line();
                line.Description = lineRecord.Description;
                line.Amount = lineRecord.Amount;
                line.AmountSpecified = lineRecord.AmountSpecified;                

                QueryService<Item> querySvcItem = new QueryService<Item>(serviceContext);
                Item itemAdded = querySvcItem.ExecuteIdsQuery("SELECT * FROM Item WHERE Name = '" + lineRecord.SalesItemLineDetailRecord.ItemRecord.Name + "'").FirstOrDefault();
                QueryService<TaxCode> querySvcTaxGSTCode = new QueryService<TaxCode>(serviceContext);
                TaxCode taxCodeGSTAdded = querySvcTaxGSTCode.ExecuteIdsQuery("SELECT * FROM TaxCode WHERE Name = '" + "GST/PST BC" + "'").FirstOrDefault();
                //TaxCode taxCodeGSTAdded = querySvcTaxGSTCode.ExecuteIdsQuery("SELECT * FROM TaxCode").FirstOrDefault();

                if (itemAdded == null)
                {
                    Item item = CreateItem(accountAdded, lineRecord.SalesItemLineDetailRecord.ItemRecord, taxCodeGSTAdded);
                    itemAdded = dataService.Add<Item>(item);
                }
                SalesItemLineDetail salesItemLineDetail = new SalesItemLineDetail();
                salesItemLineDetail.Qty = lineRecord.SalesItemLineDetailRecord.Qty;
                salesItemLineDetail.QtySpecified = true;
                salesItemLineDetail.TaxCodeRef = new ReferenceType()
                {
                    Value = taxCodeGSTAdded.Id
                }; 

                salesItemLineDetail.ItemRef = new ReferenceType()
                {
                    Value = itemAdded.Id
                };
                salesItemLineDetail.AnyIntuitObject = itemAdded.UnitPrice;
                salesItemLineDetail.ItemElementName = ItemChoiceType.UnitPrice;

                line.AnyIntuitObject = salesItemLineDetail;                
                line.DetailType = LineDetailTypeEnum.SalesItemLineDetail;                                                 
                line.DetailTypeSpecified = true;
                lineList.Add(line);
            }

            invoice.Line = lineList.ToArray();
            invoice.PONumber = invoiceRecord.PONumber;

            Random rnd = new Random();
            invoice.DocNumber = invoiceRecord.InvoiceId != "0" ? invoiceRecord.InvoiceId : rnd.Next().ToString();
            //invoice.DocNumber = rnd.Next().ToString();

            invoice.DueDate = invoiceRecord.DueDate;
            invoice.DueDateSpecified = true;
            invoice.domain = invoiceRecord.domain;
            invoice.TotalAmt = invoiceRecord.TotalAmt;
            invoice.TotalAmtSpecified = true;

            //QueryService<Term> querySvcTerm = new QueryService<Term>(serviceContext);
            //Term termAdded = querySvcTerm.ExecuteIdsQuery("SELECT * FROM Term WHERE Name = 'NET 30'").FirstOrDefault();

            //if (termAdded != null)
            //{
            //    invoice.SalesTermRef = new ReferenceType()
            //    {
            //        name = termAdded.Name,
            //        Value = termAdded.Id
            //    };
            //}            

            invoice.EmailStatus = EmailStatusEnum.NotSet;
            invoice.EmailStatusSpecified = true;
            
            invoice.TxnDate = invoiceRecord.TxnDate;
            invoice.TxnDateSpecified = true;

            invoice.CustomerMemo = new MemoRef() {                 
                Value = invoiceRecord.CustomerMemo 
            };
            
            return invoice;
        }

        public TaxCode CreateTaxCode()
        {
            TaxCode taxCode = new TaxCode();

            taxCode.Taxable = true;
            taxCode.Name = "HAETax";
            taxCode.Description = "HAETax";
            taxCode.TaxGroup = false;

            return taxCode;

        }
        #endregion       

        /// <summary>
        /// Use the Index page of App controller to get all endpoints from discovery url
        /// </summary>
        public ActionResult Error()
        {
            return View("Error");
        }

        /// <summary>
        /// Action that takes redirection from Callback URL
        /// </summary>
        public ActionResult Tokens()
        {
            return View("Tokens");
        }
    }
}