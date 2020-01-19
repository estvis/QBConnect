using QBXMLRP2Lib;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace QBTest
{
    public class QBConnect
    {
        /*
         Author: Timur Nizamov
         Versio: 1.0.0.1
         req:[
         COM dll: QBXMLRP2Lib (https://static.developer.intuit.com/resources/qbsdk130.exe)
            ]
        */
        string workCurrency = "";
        string homeCurrency = "";
        decimal exchangeRate = 1.0m;
        private string ticket;
        private RequestProcessor2 rp;
        private string maxVersion;
        private string companyFile = "";
        private QBFileMode mode = QBFileMode.qbFileOpenDoNotCare;
        public ErrorInfoData ErrorList = new ErrorInfoData();
        public string appID = "EST13";
        public string appName = "EstVis";

        public QBConnect(string companyFile)
        {
            this.companyFile = companyFile;
        }
        public QBConnect(string companyFile,string appid,string appname)
        {
            this.companyFile = companyFile;
            this.appID = appid;
            this.appName = appname;
        }
        public QBConnect()
        {
            
        }
        private async Task<bool> connectToQB()
        {
            try
            {
                rp = new RequestProcessor2(); //RequestProcessor2(); //RequestProcessor2Class();
                await Task.Run(()=>rp.OpenConnection(appID, appName));
                ticket = rp.BeginSession(companyFile, mode);
                string[] versions = (string[])rp.get_QBXMLVersionsForSession(ticket);//Array.ConvertAll<System.Array, string>(rp.get_QBXMLVersionsForSession(ticket),Convert.ToString);
                maxVersion = versions[versions.Length - 1];
                return true;
            }
            catch (Exception ex)
            {
                ErrorList.Add(DateTime.Now,ex);
                return false;
            }
        }
        private void disconnectFromQB()
        {
            if (ticket != null)
            {
                try
                {
                    rp.EndSession(ticket);
                    ticket = null;
                    rp.CloseConnection();
                    rp = null;
                }
                catch (Exception e)
                {
                    ErrorList.Add(DateTime.Now, e);
                }
            }
        }
        #region Loads
        public async Task<string[]> loadCustomers()
        {
            string[] customerList = new string[0];
            string request = "CustomerQueryRq";
           await connectToQB();
            try
            {
                int count =await getCount(request);
                string response = await processRequestFromQB(buildCustomerQueryRqXML(new string[] { "FullName" }, null));
                customerList = parseCustomerQueryRs(response, count);
            }
            catch(Exception ex)
            {
                ErrorList.Add(DateTime.Now, ex);
            }
            disconnectFromQB();

            return customerList;
        }
        public async Task<string[]> loadItems()
        {
            string[] itemList = new string[0];
            string request = "ItemQueryRq";
            await connectToQB();
            try
            {
                int count =await getCount(request);
                string response = await processRequestFromQB(buildItemQueryRqXML(new string[] { "FullName" }, null));
                itemList = parseItemQueryRs(response, count);
            }
            catch (Exception ex)
            {
                ErrorList.Add(DateTime.Now.AddMilliseconds(1), ex);
            }
            disconnectFromQB();

            return itemList;          
        }
        public async Task<string[]> loadTerms()
        {
            string[] termsList= new string[0];
            string request = "TermsQueryRq";
            await connectToQB();
            try
            {
                int count =await getCount(request);
                string response = await processRequestFromQB(buildTermsQueryRqXML());
                termsList = parseTermsQueryRs(response, count);
            }
            catch (Exception ex)
            {
                ErrorList.Add(DateTime.Now, ex);
            }
            disconnectFromQB();
            return termsList;
        }
        public async Task<string[]> loadSalesTaxCodes()
        {
            string[] salesTaxCodeList = new string[0];
            string request = "SalesTaxCodeQueryRq";
            await connectToQB();
            try
            {
                int count =await getCount(request);
                string response =await processRequestFromQB(buildSalesTaxCodeQueryRqXML());
                salesTaxCodeList = parseSalesTaxCodeQueryRs(response, count);
            }
            catch (Exception ex)
            {
                ErrorList.Add(DateTime.Now, ex);
            }
            disconnectFromQB();
            return salesTaxCodeList;
        }

        public async Task<string> getBillShipTo(string customerName, string billOrShip)
        {
            await connectToQB();
            string response =await processRequestFromQB(buildCustomerQueryRqXML(new string[] { billOrShip }, customerName));
            string[] billShipTo = parseCustomerQueryRs(response, 1);
            if (billShipTo[0] == null) billShipTo[0] = "";
            disconnectFromQB();
            return billShipTo[0];
        }
        public async Task<string> getCurrencyCode(string customerName)
        {
            await connectToQB();
            string response =await processRequestFromQB(buildCustomerQueryRqXML(new string[] { "CurrencyRef" }, customerName));
            string[] currencyCode = parseCustomerQueryRs(response, 1);
            disconnectFromQB();
            return currencyCode[0];
        }
        #endregion
        public async Task<int> getCount(string request)
        {
            string response =await processRequestFromQB(buildDataCountQuery(request));
            int count = parseRsForCount(response, request);
            return count;
        }
        #region Parsing
        private string[] parseCustomerQueryRs(string xml, int count)
        {
            /*
             <?xml version="1.0" ?> 
             <QBXML>
             <QBXMLMsgsRs>
             <CustomerQueryRs requestID="1" statusCode="0" statusSeverity="Info" statusMessage="Status OK">
                 <CustomerRet>
                     <FullName>Abercrombie, Kristy</FullName> 
                 </CustomerRet>
             </CustomerQueryRs>
             </QBXMLMsgsRs>
             </QBXML>    
            */

            string[] retVal = new string[count];
            System.IO.StringReader rdr = new System.IO.StringReader(xml);
            System.Xml.XPath.XPathDocument doc = new System.Xml.XPath.XPathDocument(rdr);
            System.Xml.XPath.XPathNavigator nav = doc.CreateNavigator();

            if (nav != null)
            {
                nav.MoveToFirstChild();
            }
            bool more = true;
            int x = 0;
            while (more)
            {
                switch (nav.LocalName)
                {
                    case "QBXML":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "QBXMLMsgsRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "CustomerQueryRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "CustomerRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "FullName":
                        retVal[x] = nav.Value.Trim();
                        x++;
                        more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    case "BillAddress":
                    case "ShipAddress":
                    case "CurrencyRef":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "Addr1":
                    case "Addr2":
                    case "Addr3":
                    case "Addr4":
                    case "Addr5":
                    case "City":
                    case "State":
                    case "PostalCode":
                        retVal[x] = retVal[x] + "\r\n" + nav.Value.Trim();
                        more = nav.MoveToNext();
                        continue;
                    default:
                        more = nav.MoveToNext();
                        continue;
                }
            }
            return retVal;
        }
        public virtual int parseRsForCount(string xml, string request)
        {
            int ret = -1;
            try
            {
                XmlNodeList RsNodeList = null;
                XmlDocument Doc = new XmlDocument();
                Doc.LoadXml(xml);
                string tagname = request.Replace("Rq", "Rs");
                RsNodeList = Doc.GetElementsByTagName(tagname);
                System.Text.StringBuilder popupMessage = new System.Text.StringBuilder();
                XmlAttributeCollection rsAttributes = RsNodeList.Item(0).Attributes;
                XmlNode retCount = rsAttributes.GetNamedItem("retCount");
                ret = Convert.ToInt32(retCount.Value);
            }
            catch (Exception e)
            {
                ErrorList.Add(DateTime.Now, e);
                ret = -1;
            }
            return ret;
        }
        private string[] parseItemQueryRs(string xml, int count)
        {
            /*
              <?xml version="1.0" ?> 
            - <QBXML>
            - <QBXMLMsgsRs>
            - <ItemQueryRs requestID="2" statusCode="0" statusSeverity="Info" statusMessage="Status OK">
            - <ItemServiceRet>
  	            <ListID>20000-933272655</ListID> 
  	            <TimeCreated>1999-07-29T11:24:15-08:00</TimeCreated> 
  	            <TimeModified>2007-12-15T11:32:53-08:00</TimeModified> 
  	            <EditSequence>1197747173</EditSequence> 
  	            <Name>Installation</Name> 
  	            <FullName>Installation</FullName> 
  	            <IsActive>true</IsActive> 
  	            <Sublevel>0</Sublevel> 
            - 	<SalesTaxCodeRef>
  		            <ListID>20000-999022286</ListID> 
  		            <FullName>Non</FullName> 
  	            </SalesTaxCodeRef>
            - 	<SalesOrPurchase>
  		            <Desc>Installation labor</Desc> 
  		            <Price>35.00</Price> 
            - 		<AccountRef>
  			            <ListID>190000-933270541</ListID> 
  			            <FullName>Construction Income:Labor Income</FullName> 
  		            </AccountRef>
  	            </SalesOrPurchase>
              </ItemServiceRet>
              </ItemQueryRs>
              </QBXMLMsgsRs>
              </QBXML>
            */

            string[] retVal = new string[count];
            System.IO.StringReader rdr = new System.IO.StringReader(xml);
            System.Xml.XPath.XPathDocument doc = new System.Xml.XPath.XPathDocument(rdr);
            System.Xml.XPath.XPathNavigator nav = doc.CreateNavigator();

            if (nav != null)
            {
                nav.MoveToFirstChild();
            }
            bool more = true;
            int x = 0;
            while (more)
            {
                switch (nav.LocalName)
                {
                    case "QBXML":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "QBXMLMsgsRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "ItemQueryRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "ItemServiceRet":
                    case "ItemNonInventoryRet":
                    case "ItemOtherChargeRet":
                    case "ItemInventoryRet":
                    case "ItemInventoryAssemblyRet":
                    case "ItemFixedAssetRet":
                    case "ItemSubtotalRet":
                    case "ItemDiscountRet":
                    case "ItemPaymentRet":
                    case "ItemSalesTaxRet":
                    case "ItemSalesTaxGroupRet":
                    case "ItemGroupRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "FullName":
                        retVal[x] = nav.Value.Trim().Length>0? nav.Value.Trim():"No name";
                        x++;
                        more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    case "SalesOrPurchase":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "Desc":
                    case "Price":
                        string val = nav.Value.Trim();
                        decimal price = 0.0m;
                        if (IsDecimal(val))
                        {
                            price = Decimal.TryParse(val, out decimal v) ? v : 0;
                            try
                            {
                                if (exchangeRate != 1.0m) { price = price / exchangeRate; }
                            }
                            catch { }
                            retVal[1] = price.ToString("N2");
                        }
                        else
                        {
                            retVal[0] = val;
                        }

                        more = nav.MoveToNext();
                        continue;
                    default:
                        more = nav.MoveToNext();
                        continue;
                }
            }
            return retVal;
        }
        private string[] parseInvoiceAddRs(string xml)
        {
            string[] retVal = new string[3];
            try
            {
                XmlNodeList RsNodeList = null;
                XmlDocument Doc = new XmlDocument();
                Doc.LoadXml(xml);
                RsNodeList = Doc.GetElementsByTagName("InvoiceAddRs");
                XmlAttributeCollection rsAttributes = RsNodeList.Item(0).Attributes;
                XmlNode statusCode = rsAttributes.GetNamedItem("statusCode");
                retVal[0] = Convert.ToString(statusCode.Value);
                XmlNode statusSeverity = rsAttributes.GetNamedItem("statusSeverity");
                retVal[1] = Convert.ToString(statusSeverity.Value);
                XmlNode statusMessage = rsAttributes.GetNamedItem("statusMessage");
                retVal[2] = Convert.ToString(statusMessage.Value);
            }
            catch (Exception e)
            {
                ErrorList.Add(DateTime.Now, e);
                retVal = null;
            }
            return retVal;
        }
        private string[] parseSalesTaxCodeQueryRs(string xml, int count)
        {
            /*
            <?xml version="1.0" ?> 
            <QBXML>
            <QBXMLMsgsRs>
            <SalesTaxCodeQueryRs requestID="3" statusCode="0" statusSeverity="Info" statusMessage="Status OK">
                <SalesTaxCodeRet>
                    <FullName>Tax</FullName> 
                </SalesTaxCodeRet>
            </SalesTaxCodeQueryRs>
            </QBXMLMsgsRs>
            </QBXML>            
            */

            string[] retVal = new string[count];
            System.IO.StringReader rdr = new System.IO.StringReader(xml);
            System.Xml.XPath.XPathDocument doc = new System.Xml.XPath.XPathDocument(rdr);
            System.Xml.XPath.XPathNavigator nav = doc.CreateNavigator();

            if (nav != null)
            {
                nav.MoveToFirstChild();
            }
            bool more = true;
            int x = 0;
            while (more)
            {
                switch (nav.LocalName)
                {
                    case "QBXML":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "QBXMLMsgsRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "SalesTaxCodeQueryRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "SalesTaxCodeRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "Name":
                        retVal[x] = nav.Value.Trim();
                        x++;
                        more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    default:
                        more = nav.MoveToNext();
                        continue;
                }
            }
            return retVal;
        }
        private string[] parseTermsQueryRs(string xml, int count)
        {
            /*
            <?xml version="1.0" ?> 
            <QBXML>
            <QBXMLMsgsRs>
            <TermsQueryRs requestID="3" statusCode="0" statusSeverity="Info" statusMessage="Status OK">
                <StandardTermsRet>
                    <Name>1% 10 Net 30</Name> 
                </StandardTermsRet>
            </TermsQueryRs>
            </QBXMLMsgsRs>
            </QBXML>            
            */

            string[] retVal = new string[count];
            System.IO.StringReader rdr = new System.IO.StringReader(xml);
            System.Xml.XPath.XPathDocument doc = new System.Xml.XPath.XPathDocument(rdr);
            System.Xml.XPath.XPathNavigator nav = doc.CreateNavigator();

            if (nav != null)
            {
                nav.MoveToFirstChild();
            }
            bool more = true;
            int x = 0;
            while (more)
            {
                switch (nav.LocalName)
                {
                    case "QBXML":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "QBXMLMsgsRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "TermsQueryRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "StandardTermsRet":
                    case "DateDrivenTermsRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "Name":
                        retVal[x] = nav.Value.Trim();
                        x++;
                        more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    default:
                        more = nav.MoveToNext();
                        continue;
                }
            }
            return retVal;
        }
        private string[] parsePreferencesQueryRs(string xml, int count)
        {
            string[] retVal = new string[count];
            System.IO.StringReader rdr = new System.IO.StringReader(xml);
            System.Xml.XPath.XPathDocument doc = new System.Xml.XPath.XPathDocument(rdr);
            System.Xml.XPath.XPathNavigator nav = doc.CreateNavigator();

            if (nav != null)
            {
                nav.MoveToFirstChild();
            }
            bool more = true;
            int x = 0;
            while (more)
            {
                switch (nav.LocalName)
                {
                    case "QBXML":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "QBXMLMsgsRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "PreferencesQueryRs":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "PreferencesRet":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "MultiCurrencyPreferences":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "HomeCurrencyRef":
                        more = nav.MoveToFirstChild();
                        continue;
                    case "FullName":
                        retVal[x] = nav.Value.Trim();
                        x++;
                        more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    case "IsMultiCurrencyOn":
                        retVal[x] = nav.Value.Trim();
                        x++;
                        //more = nav.MoveToParent();
                        more = nav.MoveToNext();
                        continue;
                    default:
                        more = nav.MoveToNext();
                        continue;
                }
            }
            return retVal;
        }

        #endregion
        private async Task<string> processRequestFromQB(string request)
        {
            try
            {               
                return await Task.Run(()=>rp.ProcessRequest(ticket, request));
            }
            catch (Exception e)
            {
                ErrorList.Add(DateTime.Now, e);
                return null;
            }
        }

        #region QueryBlders
        private string buildCustomerQueryRqXML(string[] includeRetElement, string fullName)
        {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement CustomerQueryRq = xmlDoc.CreateElement("CustomerQueryRq");
            qbXMLMsgsRq.AppendChild(CustomerQueryRq);
            if (fullName != null)
            {
                XmlElement fullNameElement = xmlDoc.CreateElement("FullName");
                CustomerQueryRq.AppendChild(fullNameElement).InnerText = fullName;
            }
            for (int x = 0; x < includeRetElement.Length; x++)
            {
                XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
                CustomerQueryRq.AppendChild(includeRet).InnerText = includeRetElement[x];
            }
            CustomerQueryRq.SetAttribute("requestID", "1");
            xml = xmlDoc.OuterXml;
            return xml;
        }

        public virtual string buildDataCountQuery(string request)
        {
            string input = "";
            XmlDocument inputXMLDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(inputXMLDoc, maxVersion);
            XmlElement queryRq = inputXMLDoc.CreateElement(request);
            queryRq.SetAttribute("metaData", "MetaDataOnly");
            qbXMLMsgsRq.AppendChild(queryRq);
            input = inputXMLDoc.OuterXml;
            return input;
        }

        private XmlElement buildRqEnvelope(XmlDocument doc, string maxVer)
        {
            doc.AppendChild(doc.CreateXmlDeclaration("1.0", null, null));
            doc.AppendChild(doc.CreateProcessingInstruction("qbxml", "version=\"" + maxVer + "\""));
            XmlElement qbXML = doc.CreateElement("QBXML");
            doc.AppendChild(qbXML);
            XmlElement qbXMLMsgsRq = doc.CreateElement("QBXMLMsgsRq");
            qbXML.AppendChild(qbXMLMsgsRq);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            return qbXMLMsgsRq;
        }
        private string buildItemQueryRqXML(string[] includeRetElement, string fullName)
        {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement ItemQueryRq = xmlDoc.CreateElement("ItemQueryRq");
            qbXMLMsgsRq.AppendChild(ItemQueryRq);
            if (fullName != null)
            {
                XmlElement fullNameElement = xmlDoc.CreateElement("FullName");
                ItemQueryRq.AppendChild(fullNameElement).InnerText = fullName;
            }
            for (int x = 0; x < includeRetElement.Length; x++)
            {
                XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
                ItemQueryRq.AppendChild(includeRet).InnerText = includeRetElement[x];
            }
            ItemQueryRq.SetAttribute("requestID", "2");
            xml = xmlDoc.OuterXml;
            return xml;
        }
        private string buildSalesTaxCodeQueryRqXML()
        {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement salesTaxCodeQueryRq = xmlDoc.CreateElement("SalesTaxCodeQueryRq");
            qbXMLMsgsRq.AppendChild(salesTaxCodeQueryRq);
            XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
            salesTaxCodeQueryRq.AppendChild(includeRet).InnerText = "Name";
            salesTaxCodeQueryRq.SetAttribute("requestID", "4");
            xml = xmlDoc.OuterXml;
            return xml;
        }
        private string buildTermsQueryRqXML()
        {
            string xml = "";
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement termsQueryRq = xmlDoc.CreateElement("TermsQueryRq");
            qbXMLMsgsRq.AppendChild(termsQueryRq);
            XmlElement includeRet = xmlDoc.CreateElement("IncludeRetElement");
            termsQueryRq.AppendChild(includeRet).InnerText = "Name";
            termsQueryRq.SetAttribute("requestID", "3");
            xml = xmlDoc.OuterXml;
            return xml;
        }
        private string buildInvoiceAddRqXML(InvoiceData inv)
        {
            string requestXML = "";

            //if (!validateInput()) return null;

            //GET ALL INPUT INTO XML
            XmlDocument xmlDoc = new XmlDocument();
            XmlElement qbXMLMsgsRq = buildRqEnvelope(xmlDoc, maxVersion);
            qbXMLMsgsRq.SetAttribute("onError", "stopOnError");
            XmlElement InvoiceAddRq = xmlDoc.CreateElement("InvoiceAddRq");
            qbXMLMsgsRq.AppendChild(InvoiceAddRq);
            XmlElement InvoiceAdd = xmlDoc.CreateElement("InvoiceAdd");
            InvoiceAddRq.AppendChild(InvoiceAdd);

            // CustomerRef -> FullName
            if (inv.CustomerFullName!= "")
            {
                XmlElement Element_CustomerRef = xmlDoc.CreateElement("CustomerRef");
                InvoiceAdd.AppendChild(Element_CustomerRef);
                XmlElement Element_CustomerRef_FullName = xmlDoc.CreateElement("FullName");
                Element_CustomerRef.AppendChild(Element_CustomerRef_FullName).InnerText = inv.CustomerFullName;
            }

            // TxnDate 
            DateTime DT_TxnDate = System.DateTime.Today;
            if (inv.InvDate!= "")
            {               
                XmlElement Element_TxnDate = xmlDoc.CreateElement("TxnDate");
                InvoiceAdd.AppendChild(Element_TxnDate).InnerText = inv.InvDate;
            }

            // RefNumber 
            if (inv.InvNo != "")
            {
                XmlElement Element_RefNumber = xmlDoc.CreateElement("RefNumber");
                InvoiceAdd.AppendChild(Element_RefNumber).InnerText = inv.InvNo;
            }

            // BillAddress
            if (inv.BillTo != "")
            {
                string[] BillAddress = inv.BillTo.Split('\n');
                XmlElement Element_BillAddress = xmlDoc.CreateElement("BillAddress");
                InvoiceAdd.AppendChild(Element_BillAddress);
                for (int i = 0; i < BillAddress.Length; i++)
                {
                    if (BillAddress[i] != "" || BillAddress[i] != null)
                    {
                        XmlElement Element_Addr = xmlDoc.CreateElement("Addr" + (i + 1));
                        Element_BillAddress.AppendChild(Element_Addr).InnerText = BillAddress[i];
                    }
                }
            }

            // TermsRef -> FullName 
            bool termsAvailable = false;
            if (inv.Term!= "")
            {
                termsAvailable = true;
                XmlElement Element_TermsRef = xmlDoc.CreateElement("TermsRef");
                InvoiceAdd.AppendChild(Element_TermsRef);
                XmlElement Element_TermsRef_FullName = xmlDoc.CreateElement("FullName");
                Element_TermsRef.AppendChild(Element_TermsRef_FullName).InnerText = inv.Term;
            }

            // DueDate 
            if (termsAvailable)
            {
                DateTime DT_DueDate = System.DateTime.Today;
                double dueInDays = getDueInDays(inv.InvDate);
                DT_DueDate = DT_TxnDate.AddDays(dueInDays);
                string DueDate = getDateString(DT_DueDate);
                XmlElement Element_DueDate = xmlDoc.CreateElement("DueDate");
                InvoiceAdd.AppendChild(Element_DueDate).InnerText = DueDate;
            }

            // CustomerMsgRef -> FullName 
            if (inv.CustomerMessage != "")
            {
                XmlElement Element_CustomerMsgRef = xmlDoc.CreateElement("CustomerMsgRef");
                InvoiceAdd.AppendChild(Element_CustomerMsgRef);
                XmlElement Element_CustomerMsgRef_FullName = xmlDoc.CreateElement("FullName");
                Element_CustomerMsgRef.AppendChild(Element_CustomerMsgRef_FullName).InnerText = inv.CustomerMessage;
            }

            // ExchangeRate 
            if (inv.ExchangeRate!= "")
            {
                XmlElement Element_ExchangeRate = xmlDoc.CreateElement("ExchangeRate");
                InvoiceAdd.AppendChild(Element_ExchangeRate).InnerText = inv.ExchangeRate;
            }

            //Line Items
            XmlElement Element_InvoiceLineAdd;

            foreach(InvoiceRow r in inv.Rows)

            for (int x = 1; x < 6; x++)
            {
                Element_InvoiceLineAdd = xmlDoc.CreateElement("InvoiceLineAdd");
                InvoiceAdd.AppendChild(Element_InvoiceLineAdd);                
                if (r.ItemRef != "")
                {
                    XmlElement Element_InvoiceLineAdd_ItemRef = xmlDoc.CreateElement("ItemRef");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_ItemRef);
                    XmlElement Element_InvoiceLineAdd_ItemRef_FullName = xmlDoc.CreateElement("FullName");
                    Element_InvoiceLineAdd_ItemRef.AppendChild(Element_InvoiceLineAdd_ItemRef_FullName).InnerText = r.ItemRef;
                }
                if (r.Desc != "")
                {
                    XmlElement Element_InvoiceLineAdd_Desc = xmlDoc.CreateElement("Desc");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Desc).InnerText = r.Desc;
                }
                if (r.Quantity != "")
                {
                    XmlElement Element_InvoiceLineAdd_Quantity = xmlDoc.CreateElement("Quantity");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Quantity).InnerText = r.Quantity;
                }
                if (r.Rate != "")
                {
                    XmlElement Element_InvoiceLineAdd_Rate = xmlDoc.CreateElement("Rate");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Rate).InnerText = r.Rate;
                }
                if (r.Amount != "")
                {
                    XmlElement Element_InvoiceLineAdd_Amount = xmlDoc.CreateElement("Amount");
                    Element_InvoiceLineAdd.AppendChild(Element_InvoiceLineAdd_Amount).InnerText = r.Amount;
                }
            }


            InvoiceAddRq.SetAttribute("requestID", "99");

            requestXML = xmlDoc.OuterXml;

            return requestXML;
        }
        #endregion

        public double getDueInDays(string val)
        {
            double dueInDays = 0;
            switch (val)
            {
                case "Due on receipt":
                    dueInDays = 0;
                    break;
                case "Net 15":
                    dueInDays = 15;
                    break;
                case "Net 30":
                    dueInDays = 30;
                    break;
                case "Net 60":
                    dueInDays = 60;
                    break;
                default:
                    dueInDays = 0;
                    break;
            }
            return dueInDays;
        }
        public string getDateString(DateTime dt)
        {
            string year = dt.Year.ToString();
            string month = dt.Month.ToString();
            if (month.Length < 2) month = "0" + month;
            string day = dt.Day.ToString();
            if (day.Length < 2) day = "0" + day;
            return year + "-" + month + "-" + day;
        }
        public static bool IsDecimal(string theValue)
        {
            bool returnVal = false;
            try
            {
                Convert.ToDouble(theValue, System.Globalization.CultureInfo.CurrentCulture);
                returnVal = true;
            }
            catch
            {
                returnVal = false;
            }
            finally
            {
            }

            return returnVal;

        }

        public decimal ParseDecimal(object val)
        {
            decimal res = 0m;
            if(decimal.TryParse(val.ToString().Replace(",","."), out decimal d))
            {
                res = d;
            }
            if (decimal.TryParse(val.ToString().Replace(".", ","), out decimal d1))
            {
                res = d1;
            }

            return res;
        }

        #region Addmethods
        public async Task<string> AddInvoice(InvoiceData inv)
        {
            string msg = "";
            string requestXML = buildInvoiceAddRqXML(inv);
            if (requestXML == null)
            {
                msg = "One of the input is missing. Double-check your entries and then click Save again. Error saving invoice";
                return msg;
            }
            await connectToQB();
            string response = await processRequestFromQB(requestXML);
            disconnectFromQB();
            string[] status = new string[3];
            if (response != null) status = parseInvoiceAddRs(response);
           

            if (response != null & status[0] == "0")
            {
                msg = "Invoice was added successfully!";
            }
            else
            {
                msg = "Could not add invoice.";
            }

            msg = msg + "\n\n";
            msg = msg + "Status Code = " + status[0] + "\n";
            msg = msg + "Status Severity = " + status[1] + "\n";
            msg = msg + "Status Message = " + status[2] + "\n";
            //MessageBox.Show(msg);
            return msg;
        }
        #endregion

    }

    public abstract class InvoiceRow
    {
        public string ItemRef;
        public string Desc;
        public string Quantity;
        public string Rate;
        public string Amount;

        protected InvoiceRow()
        {
            ItemRef = "";
            Desc = "";
            Quantity = "";
            Rate = "";
            Amount = "";
        }
    }
    public abstract class InvoiceData
    {
        public string CustomerFullName;
        public string InvDate;
        public string InvNo;
        public string BillTo;
        public string Term;
        public string ExchangeRate;
        public string CustomerMessage;
        public List<InvoiceRow> Rows;

        protected InvoiceData()
        {
            CustomerFullName = "";
            InvDate = "";
            InvNo = "";
            BillTo = "";
            Term = "";
            ExchangeRate = "";
            CustomerMessage = "";
            Rows = new List<InvoiceRow>();
        }

        public void SetDate(DateTime dt)
        {
            string year = dt.Year.ToString();
            string month = dt.Month.ToString();
            if (month.Length < 2) month = "0" + month;
            string day = dt.Day.ToString();
            if (day.Length < 2) day = "0" + day;
            InvDate = year + "-" + month + "-" + day;            
        }
    }

    public class ErrorInfoData
    {
        public List<Dictionary<DateTime,Exception>> ErrorList = new List<Dictionary<DateTime, Exception>>();
        public void Add(DateTime dt,Exception ex)
        {
            Dictionary<DateTime, Exception> it = new Dictionary<DateTime, Exception>();
            it.Add(dt, ex);
            ErrorList.Add(it);
        }
    }
}
