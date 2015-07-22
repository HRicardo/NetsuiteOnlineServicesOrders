using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using NetsuiteOnlineServicesOrders.wsNetsuite;

namespace NetsuiteOnlineServicesOrders
{
    class Util
    {
        public static NetSuiteService _service;

        public static string Login(string strUser, string strPassword, string strAccount, string strRole)
        {
            string strResult = null;
            try
            {
                Passport passport = new Passport();
                RecordRef role = new RecordRef();
                passport.email = strUser;
                passport.password = strPassword;
                role.internalId = strRole;
                passport.role = role;
                passport.account = strAccount;

                _service = new NetSuiteService();
                _service.Timeout = 1000 * 60 * 60 * 2;
                _service.passport = passport;
                SessionResponse objResponse = _service.login(passport);
                if (!objResponse.status.isSuccess)
                {
                    strResult = objResponse.status.statusDetail[0].message;
                }
                Console.WriteLine(_service.Url);
            }
            catch (Exception objExc)
            {
                strResult = objExc.Message;
            }
            return strResult;
        }

        public static ArrayList GetUsers(string strRole)
        {
            ArrayList arrList = new ArrayList();
            try
            {
                EmployeeSearch objEmployeeSearch = new EmployeeSearch();
                objEmployeeSearch.basic = new EmployeeSearchBasic();
                objEmployeeSearch.basic.role = new SearchMultiSelectField();
                objEmployeeSearch.basic.role.@operator = SearchMultiSelectFieldOperator.anyOf;
                objEmployeeSearch.basic.role.operatorSpecified = true;
                objEmployeeSearch.basic.role.searchValue = new RecordRef[1];
                objEmployeeSearch.basic.role.searchValue[0] = new RecordRef();
                objEmployeeSearch.basic.role.searchValue[0].internalId = strRole;
                objEmployeeSearch.basic.role.searchValue[0].type = RecordType.salesRole;
                objEmployeeSearch.basic.role.searchValue[0].typeSpecified = true;

                SearchResult objSearchResult = _service.search(objEmployeeSearch);
                if (objSearchResult.status.isSuccess && objSearchResult.recordList != null)
                {
                    foreach (Employee objEmployee in objSearchResult.recordList)
                    {
                        arrList.Add(objEmployee);
                    }
                }
            }
            catch (Exception objExc)
            {
            }
            return arrList;
        }

        public static Customer GetCustomer(string strCustomerName)
        {
            RecordRef reseller = new RecordRef();
            Customer objCustomer = new Customer();
            try
            {
                CustomerSearch custSearch = new CustomerSearch();
                custSearch.basic = new CustomerSearchBasic();
                custSearch.basic.entityId = new SearchStringField();
                custSearch.basic.entityId.@operator = SearchStringFieldOperator.@is;
                custSearch.basic.entityId.operatorSpecified = true;
                custSearch.basic.entityId.searchValue = strCustomerName;

                SearchResult res = _service.search(custSearch);
                if (res.status.isSuccess)
                {
                    if (res.recordList != null && res.recordList.Length == 1)
                    {
                        reseller.type = RecordType.customer;
                        reseller.typeSpecified = true;
                        System.String entID = ((Customer)(res.recordList[0])).entityId;
                        reseller.name = entID;
                        reseller.internalId = ((Customer)(res.recordList[0])).internalId;

                        ReadResponse objReadResponse = _service.get(reseller);
                        if (objReadResponse.status.isSuccess)
                        {
                            objCustomer = (Customer)objReadResponse.record;
                        }
                    }
                }
            }
            catch (Exception objExc)
            {
            }

            return objCustomer;
        }

        public static RecordRef GetCustomer(string strParent, string strPhone)
        {
            RecordRef objCustomer = null;
            CustomerSearch custSearch = new CustomerSearch();
            SearchStringField customerPhone = new SearchStringField();
            customerPhone.@operator = SearchStringFieldOperator.@is;
            customerPhone.operatorSpecified = true;
            customerPhone.searchValue = strPhone;

            CustomerSearchBasic custBasic = new CustomerSearchBasic();
            custBasic.phone = customerPhone;
            custBasic.parent = new SearchMultiSelectField();
            custBasic.parent.@operator = SearchMultiSelectFieldOperator.anyOf;
            custBasic.parent.operatorSpecified = true;
            custBasic.parent.searchValue = new RecordRef[1];
            custBasic.parent.searchValue[0] = new RecordRef();
            custBasic.parent.searchValue[0].internalId = strParent;
            custBasic.parent.searchValue[0].type = RecordType.customer;
            custBasic.parent.searchValue[0].typeSpecified = true;

            custSearch.basic = custBasic;

            SearchResult res = _service.search(custSearch);
            if (res.status.isSuccess)
            {
                if (res.recordList != null && res.recordList.Length == 1)
                {
                    objCustomer = new RecordRef();
                    objCustomer.type = RecordType.customer;
                    objCustomer.typeSpecified = true;
                    System.String entID = ((Customer)(res.recordList[0])).entityId;
                    objCustomer.name = entID;
                    objCustomer.internalId = ((Customer)(res.recordList[0])).internalId;
                }
            }

            return objCustomer;
        }

        public static string GetAccountId(string strAccountNumber)
        {
            string strId = "";
            AccountSearch objAccountSearch = new AccountSearch();
            objAccountSearch.basic = new AccountSearchBasic();
            objAccountSearch.basic.number = new SearchStringField();
            objAccountSearch.basic.number.@operator = SearchStringFieldOperator.@is;
            objAccountSearch.basic.number.operatorSpecified = true;
            objAccountSearch.basic.number.searchValue = strAccountNumber;

            SearchResult objSearchResult = _service.search(objAccountSearch);
            if (objSearchResult.recordList != null && objSearchResult.recordList.Length == 1)
            {
                if (objSearchResult.recordList[0] is Account)
                {
                    strId = ((Account)objSearchResult.recordList[0]).internalId;
                }
            }
            return strId;
        }

        public static string GetPackedBunch(string strQuoteName)
        {
            string strPackedBunchId = null;
            try
            {

                string strQuoteID = int.Parse(strQuoteName.Replace("QB", "")).ToString();
                CustomRecordSearch objSearch = new CustomRecordSearch();
                objSearch.basic = new CustomRecordSearchBasic();
                objSearch.basic.recType = new RecordRef();
                objSearch.basic.recType.internalId = "524";

                objSearch.basic.isInactive = new SearchBooleanField();
                objSearch.basic.isInactive.searchValue = false;
                objSearch.basic.isInactive.searchValueSpecified = true;

                SearchMultiSelectCustomField objQuoteId = new SearchMultiSelectCustomField();
                objQuoteId.internalId = "custrecord_vf_qd_quote";
                objQuoteId.@operator = SearchMultiSelectFieldOperator.anyOf;
                objQuoteId.operatorSpecified = true;
                objQuoteId.searchValue = new ListOrRecordRef[1];
                objQuoteId.searchValue[0] = new ListOrRecordRef();
                objQuoteId.searchValue[0].internalId = strQuoteID;

                /*SearchStringCustomField objBatchName = new SearchStringCustomField();
                objBatchName.internalId = "custrecordfb_batchname";
                objBatchName.@operator = SearchStringFieldOperator.@is;
                objBatchName.operatorSpecified = true;
                objBatchName.searchValue = "CM1309MP007";*/

                objSearch.basic.customFieldList = new SearchCustomField[] { objQuoteId };
                CustomRecordSearchAdvanced objSearchAdv = new CustomRecordSearchAdvanced();
                objSearchAdv.savedSearchScriptId = "customsearch_spm_quotedetail_export";
                objSearchAdv.criteria = objSearch;

                SearchResult objSearchResult = _service.search(objSearch);
                if (objSearchResult.recordList != null && objSearchResult.recordList.Length > 0)
                {
                    for (int j = 0; j < objSearchResult.recordList.Length; j++)
                    {
                        if (objSearchResult.recordList[j] is CustomRecord)
                        {
                            CustomRecord objRecord = (CustomRecord)objSearchResult.recordList[j];
                            strPackedBunchId = GetValue("custrecord_vf_qd_packedbunch", objRecord.customFieldList).ToString();
                            strPackedBunchId = "PB" + strPackedBunchId.PadLeft(8, '0');
                        }
                    }
                }
            }
            catch (Exception)
            {
            }

            return strPackedBunchId;
        }

        public static ArrayList GetItemId(string strProductId)
        {
            ArrayList arrItem = new ArrayList();
            try
            {
                strProductId = strProductId.Trim();
                ItemSearchBasic objItemSearchBasic = new ItemSearchBasic();
                objItemSearchBasic.vendorName = new SearchStringField();
                if (strProductId.StartsWith("QB0"))
                {
                    strProductId = GetPackedBunch(strProductId);
                    objItemSearchBasic.vendorName.@operator = SearchStringFieldOperator.@is;
                    if (string.IsNullOrEmpty(strProductId))
                    {
                        return arrItem;
                    }
                }
                else
                {
                    objItemSearchBasic.vendorName.@operator = SearchStringFieldOperator.startsWith;
                }
                objItemSearchBasic.vendorName.searchValue = strProductId;// "CST34434A_B";
                objItemSearchBasic.vendorName.operatorSpecified = true;

                objItemSearchBasic.isInactive = new SearchBooleanField();
                objItemSearchBasic.isInactive.searchValue = false;
                objItemSearchBasic.isInactive.searchValueSpecified = true;

                ItemSearch objItemSearch = new ItemSearch();
                objItemSearch.basic = objItemSearchBasic;
                SearchResult objSearchResult = _service.search(objItemSearch);
                if (objSearchResult.recordList != null && objSearchResult.recordList.Length > 0)
                {
                    for (int i = 0; i < objSearchResult.recordList.Length; i++)
                    {
                        if (objSearchResult.recordList[i] is NonInventoryResaleItem)
                        {
                            NonInventoryResaleItem objItem = (NonInventoryResaleItem)objSearchResult.recordList[i];
                            arrItem.Add(new string[] { objItem.internalId, objItem.internalId + " " + objItem.vendorName + ": " + objItem.salesDescription });
                        }
                        if (objSearchResult.recordList[i] is LotNumberedAssemblyItem)
                        {
                            LotNumberedAssemblyItem objItem = (LotNumberedAssemblyItem)objSearchResult.recordList[i];
                            arrItem.Add(new string[] { objItem.internalId, objItem.internalId + " " + objItem.vendorName + ": " + objItem.displayName });
                        }
                    }
                }
            }
            catch (Exception)
            {
            }
            return arrItem;
        }

        public static ArrayList GetItemByItemNumber(string strProductId)
        {
            ArrayList arrItem = new ArrayList();
            try
            {
                strProductId = strProductId.Trim();

                SearchStringCustomField objItemNumber = new SearchStringCustomField();
                objItemNumber.@operator = SearchStringFieldOperator.@is;
                objItemNumber.operatorSpecified = true;
                objItemNumber.searchValue = strProductId;
                objItemNumber.internalId = "custitemcst_itemnumber";

                ItemSearchBasic objItemSearchBasic = new ItemSearchBasic();
                objItemSearchBasic.customFieldList = new SearchCustomField[] { objItemNumber };

                objItemSearchBasic.isInactive = new SearchBooleanField();
                objItemSearchBasic.isInactive.searchValue = false;
                objItemSearchBasic.isInactive.searchValueSpecified = true;

                ItemSearch objItemSearch = new ItemSearch();
                objItemSearch.basic = objItemSearchBasic;
                SearchResult objSearchResult = _service.search(objItemSearch);
                if (objSearchResult.recordList != null && objSearchResult.recordList.Length > 0)
                {
                    for (int i = 0; i < objSearchResult.recordList.Length; i++)
                    {
                        if (objSearchResult.recordList[i] is NonInventoryResaleItem)
                        {
                            NonInventoryResaleItem objItem = (NonInventoryResaleItem)objSearchResult.recordList[i];
                            arrItem.Add(new string[] { objItem.internalId, objItem.internalId + " " + objItem.vendorName + ": " + objItem.salesDescription });
                        }
                        if (objSearchResult.recordList[i] is LotNumberedAssemblyItem)
                        {
                            LotNumberedAssemblyItem objItem = (LotNumberedAssemblyItem)objSearchResult.recordList[i];
                            arrItem.Add(new string[] { objItem.internalId, objItem.internalId + " " + objItem.vendorName + ": " + objItem.displayName });
                        }
                    }
                }
            }
            catch (Exception)
            {
            }
            return arrItem;
        }

        public static NonInventoryResaleItem GetItem(string strProductId)
        {
            NonInventoryResaleItem objItem = new NonInventoryResaleItem();
            try
            {
                ItemSearchBasic objItemSearchBasic = new ItemSearchBasic();
                objItemSearchBasic.vendorCode = new SearchStringField();
                objItemSearchBasic.vendorCode.@operator = SearchStringFieldOperator.startsWith;
                objItemSearchBasic.vendorCode.searchValue = strProductId;// "CST34434A_B";
                objItemSearchBasic.vendorCode.operatorSpecified = true;

                objItemSearchBasic.isInactive = new SearchBooleanField();
                objItemSearchBasic.isInactive.searchValue = false;
                objItemSearchBasic.isInactive.searchValueSpecified = false;

                ItemSearch objItemSearch = new ItemSearch();
                objItemSearch.basic = objItemSearchBasic;
                SearchResult objSearchResult = _service.search(objItemSearch);
                if (objSearchResult.recordList != null && objSearchResult.recordList.Length == 1)
                {
                    if (objSearchResult.recordList[0] is NonInventoryResaleItem)
                    {
                        objItem = (NonInventoryResaleItem)objSearchResult.recordList[0];
                        //RecordRef objRecord = new RecordRef();
                        //objRecord.internalId = objItem.internalId;
                        //objRecord.type = RecordType.nonInventoryResaleItem;
                        //objRecord.typeSpecified = true;
                        //ReadResponse objResponse = _service.get(objRecord);
                        //objItem = (NonInventoryResaleItem)objResponse.record;
                    }
                }
            }
            catch (Exception)
            {
            }
            return objItem;
        }

        public static string GetOrderId(string strPONumber, string strReseller)
        {
            string strOrderId = "";
            try
            {
                TransactionSearch xactionSearch = new TransactionSearch();

                SearchEnumMultiSelectField searchSalesOrderField = new SearchEnumMultiSelectField();
                searchSalesOrderField.@operator = SearchEnumMultiSelectFieldOperator.anyOf;
                searchSalesOrderField.operatorSpecified = true;
                System.String[] soStringArray = new System.String[1];
                soStringArray[0] = "_salesOrder";
                searchSalesOrderField.searchValue = soStringArray;

                TransactionSearchBasic xactionBasic = new TransactionSearchBasic();
                xactionBasic.type = searchSalesOrderField;
                xactionBasic.otherRefNum = new SearchTextNumberField();
                xactionBasic.otherRefNum.@operator = SearchTextNumberFieldOperator.equalTo;
                xactionBasic.otherRefNum.operatorSpecified = true;
                xactionBasic.otherRefNum.searchValue = strPONumber;

                SearchMultiSelectCustomField objReseller = new SearchMultiSelectCustomField();
                objReseller.internalId = "custbodytbfcustomerid";
                objReseller.@operator = SearchMultiSelectFieldOperator.anyOf;
                objReseller.operatorSpecified = true;
                objReseller.searchValue = new ListOrRecordRef[1];
                objReseller.searchValue[0] = new ListOrRecordRef();
                objReseller.searchValue[0].internalId = strReseller;
                xactionBasic.customFieldList = new SearchCustomField[] { objReseller };

                xactionSearch.basic = xactionBasic;

                SearchResult objSearchResult = _service.search(xactionSearch);
                if (objSearchResult.recordList != null)
                {
                    Record[] recordList;
                    for (int i = 1; i <= objSearchResult.totalPages; i++)
                    {
                        recordList = objSearchResult.recordList;

                        for (int j = 0; j < recordList.Length; j++)
                        {
                            if (recordList[j] is SalesOrder)
                            {
                                SalesOrder so = (SalesOrder)(recordList[j]);
                                strOrderId = so.internalId;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
            }
            return strOrderId;
        }

        public static Record GetEntity(string strEntityId, RecordType objType)
        {
            Record objRecord = null;
            try
            {
                RecordRef objRef = new RecordRef();
                objRef.internalId = strEntityId;
                objRef.type = objType;
                objRef.typeSpecified = true;
                ReadResponse objReadResponse = _service.get(objRef);
                objRecord = objReadResponse.record;
            }
            catch (Exception objExc)
            {
            }
            return objRecord;
        }

        public static SalesOrder GetOrder(string strOrderId)
        {
            SalesOrder objSalesOrder = new SalesOrder();
            try
            {
                RecordRef objRef = new RecordRef();
                objRef.internalId = strOrderId;
                objRef.type = RecordType.salesOrder;
                objRef.typeSpecified = true;
                ReadResponse objReadResponse = _service.get(objRef);
                objSalesOrder = (SalesOrder)(objReadResponse.record);
            }
            catch (Exception)
            {
            }
            return objSalesOrder;
        }

        public static SalesOrder GetOrderByExternalID(string strId)
        {
            SalesOrder objSalesOrder = new SalesOrder();
            try
            {
                RecordRef objRef = new RecordRef();
                objRef.externalId = strId;
                objRef.type = RecordType.salesOrder;
                objRef.typeSpecified = true;
                ReadResponse objReadResponse = _service.get(objRef);
                objSalesOrder = (SalesOrder)(objReadResponse.record);
            }
            catch (Exception)
            {
            }
            return objSalesOrder;
        }

        public static string CloseSalesOrder(SalesOrder objSalesOrder)
        {
            string booResult = null;
            try
            {
                SalesOrder objNewSalesOrder = new SalesOrder();
                objNewSalesOrder.internalId = objSalesOrder.internalId;
                objNewSalesOrder.itemList = new SalesOrderItemList();
                objNewSalesOrder.itemList.item = new SalesOrderItem[objSalesOrder.itemList.item.Length];
                for (int i = 0; i < objSalesOrder.itemList.item.Length; i++)
                {
                    objNewSalesOrder.itemList.item[i] = new SalesOrderItem();
                    objNewSalesOrder.itemList.item[i].item = objSalesOrder.itemList.item[i].item;
                    objNewSalesOrder.itemList.item[i].line = objSalesOrder.itemList.item[i].line;
                    objNewSalesOrder.itemList.item[i].lineSpecified = true;
                    objNewSalesOrder.itemList.item[i].isClosed = true;
                    objNewSalesOrder.itemList.item[i].isClosedSpecified = true;
                }

                WriteResponse objWriteResponse = _service.update(objNewSalesOrder);
                if (!objWriteResponse.status.isSuccess)
                {
                    booResult = objWriteResponse.status.statusDetail[0].code+" : "+objWriteResponse.status.statusDetail[0].message;
                }
            }
            catch (Exception)
            {
            }
            return booResult;
        }

        public static string GetOrderId(string strCustomerId, string strPONumber, string strOrderNumber)
        {
            string strItem = "";
            try
            {
                TransactionSearch xactionSearch = new TransactionSearch();
                //SearchMultiSelectField entity = new SearchMultiSelectField();
                //entity.@operator = SearchMultiSelectFieldOperator.anyOf;
                //entity.operatorSpecified = true;
                //RecordRef custRecordRef = new RecordRef();
                //custRecordRef.internalId = strCustomerId;
                //custRecordRef.type = RecordType.customer;
                //custRecordRef.typeSpecified = true;
                //RecordRef[] custRecordRefArray = new RecordRef[1];
                //custRecordRefArray[0] = custRecordRef;
                //entity.searchValue = custRecordRefArray;

                SearchEnumMultiSelectField searchSalesOrderField = new SearchEnumMultiSelectField();
                searchSalesOrderField.@operator = SearchEnumMultiSelectFieldOperator.anyOf;
                searchSalesOrderField.operatorSpecified = true;
                System.String[] soStringArray = new System.String[1];
                soStringArray[0] = "_salesOrder";
                searchSalesOrderField.searchValue = soStringArray;

                TransactionSearchBasic xactionBasic = new TransactionSearchBasic();
                xactionBasic.type = searchSalesOrderField;
                //xactionBasic.entity = entity;
                xactionBasic.otherRefNum = new SearchTextNumberField();
                xactionBasic.otherRefNum.@operator = SearchTextNumberFieldOperator.equalTo;
                xactionBasic.otherRefNum.operatorSpecified = true;
                xactionBasic.otherRefNum.searchValue = strPONumber;

                SearchStringCustomField objOrderNumber = new SearchStringCustomField();
                objOrderNumber.internalId = "custbodycst_ordernumber";
                objOrderNumber.@operator = SearchStringFieldOperator.contains;
                objOrderNumber.operatorSpecified = true;
                objOrderNumber.searchValue = strOrderNumber;
                xactionBasic.customFieldList = new SearchCustomField[] { objOrderNumber };

                xactionSearch.basic = xactionBasic;

                SearchResult objSearchResult = _service.search(xactionSearch);
                if (objSearchResult.status != null && objSearchResult.status.isSuccess)
                {
                    if (objSearchResult.recordList != null)
                    {
                        Record[] recordList;
                        for (int i = 1; i <= objSearchResult.totalPages; i++)
                        {
                            recordList = objSearchResult.recordList;

                            for (int j = 0; j < recordList.Length; j++)
                            {
                                if (recordList[j] is SalesOrder)
                                {
                                    SalesOrder so = (SalesOrder)(recordList[j]);
                                    strItem = so.internalId;
                                    return strItem;
                                }
                            }
                        }
                    }
                }
                else
                {
                    strItem = "error";
                }
            }
            catch (Exception)
            {
            }
            return strItem;
        }

        public static void DeleteSalesOrder()
        {
            string strItem = "";
            try
            {
                TransactionSearch xactionSearch = new TransactionSearch();
                //xactionSearch.basic
                //SearchMultiSelectField entity = new SearchMultiSelectField();
                //entity.@operator = SearchMultiSelectFieldOperator.anyOf;
                //entity.operatorSpecified = true;
                //RecordRef custRecordRef = new RecordRef();
                //custRecordRef.internalId = strCustomerId;
                //custRecordRef.type = RecordType.customer;
                //custRecordRef.typeSpecified = true;
                //RecordRef[] custRecordRefArray = new RecordRef[1];
                //custRecordRefArray[0] = custRecordRef;
                //entity.searchValue = custRecordRefArray;

                SearchEnumMultiSelectField searchSalesOrderField = new SearchEnumMultiSelectField();
                searchSalesOrderField.@operator = SearchEnumMultiSelectFieldOperator.anyOf;
                searchSalesOrderField.operatorSpecified = true;
                System.String[] soStringArray = new System.String[1];
                soStringArray[0] = "_salesOrder";
                searchSalesOrderField.searchValue = soStringArray;

                TransactionSearchBasic xactionBasic = new TransactionSearchBasic();
                xactionBasic.type = searchSalesOrderField;
                xactionBasic.customFieldList = new SearchCustomField[1];

                SearchStringCustomField objNumber = new SearchStringCustomField();
                objNumber.internalId = "custbodycst_ordernumber";
                objNumber.@operator = SearchStringFieldOperator.notEmpty;
                objNumber.operatorSpecified = true;
                xactionBasic.customFieldList[0] = objNumber;
                xactionSearch.basic = xactionBasic;

                SearchResult objSearchResult = _service.search(xactionSearch);
                if (objSearchResult.status != null && objSearchResult.status.isSuccess)
                {
                    if (objSearchResult.recordList != null)
                    {
                        Record[] recordList;
                        for (int i = 1; i <= objSearchResult.totalPages; i++)
                        {
                            recordList = objSearchResult.recordList;

                            for (int j = 0; j < recordList.Length; j++)
                            {
                                if (recordList[j] is SalesOrder)
                                {
                                    try
                                    {
                                        SalesOrder objOrder = GetSalesOrder(((SalesOrder)(recordList[j])).internalId);
                                        object objCustomReseller = GetValue("custbodytbfcustomerid", objOrder.customFieldList);
                                        if (objCustomReseller.ToString().Equals("CST100WA"))
                                        {
                                            RecordRef objRef = new RecordRef();
                                            objRef.internalId = objOrder.internalId;
                                            objRef.type = RecordType.salesOrder;
                                            objRef.typeSpecified = true;
                                            WriteResponse writeRes = _service.delete(objRef);
                                            if (writeRes.status.isSuccess)
                                            {
                                                Console.WriteLine("delete: " + objOrder.tranId);
                                            }
                                            else
                                            {
                                                Console.WriteLine(getStatusDetails(writeRes.status));
                                            }

                                        }
                                    }
                                    catch (Exception objExc)
                                    {
                                        Console.WriteLine(objExc.ToString());
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    strItem = "error";
                }
            }
            catch (Exception)
            {
            }
        }

        public static void DeletePurchaseOrder()
        {
            string strItem = "";
            try
            {
                TransactionSearch xactionSearch = new TransactionSearch();
                //xactionSearch.basic
                //SearchMultiSelectField entity = new SearchMultiSelectField();
                //entity.@operator = SearchMultiSelectFieldOperator.anyOf;
                //entity.operatorSpecified = true;
                //RecordRef custRecordRef = new RecordRef();
                //custRecordRef.internalId = strCustomerId;
                //custRecordRef.type = RecordType.customer;
                //custRecordRef.typeSpecified = true;
                //RecordRef[] custRecordRefArray = new RecordRef[1];
                //custRecordRefArray[0] = custRecordRef;
                //entity.searchValue = custRecordRefArray;

                SearchEnumMultiSelectField searchSalesOrderField = new SearchEnumMultiSelectField();
                searchSalesOrderField.@operator = SearchEnumMultiSelectFieldOperator.anyOf;
                searchSalesOrderField.operatorSpecified = true;
                System.String[] soStringArray = new System.String[1];
                soStringArray[0] = "_purchaseOrder";
                searchSalesOrderField.searchValue = soStringArray;

                TransactionSearchBasic xactionBasic = new TransactionSearchBasic();
                xactionBasic.type = searchSalesOrderField;
                xactionBasic.customFieldList = new SearchCustomField[1];

                SearchStringCustomField objNumber = new SearchStringCustomField();
                objNumber.internalId = "custbodycst_ordernumber";
                objNumber.@operator = SearchStringFieldOperator.notEmpty;
                objNumber.operatorSpecified = true;
                xactionBasic.customFieldList[0] = objNumber;
                xactionSearch.basic = xactionBasic;

                SearchResult objSearchResult = _service.search(xactionSearch);
                if (objSearchResult.status != null && objSearchResult.status.isSuccess)
                {
                    if (objSearchResult.recordList != null)
                    {
                        Record[] recordList;
                        for (int i = 1; i <= objSearchResult.totalPages; i++)
                        {
                            recordList = objSearchResult.recordList;

                            for (int j = 0; j < recordList.Length; j++)
                            {
                                if (recordList[j] is PurchaseOrder)
                                {
                                    try
                                    {
                                        PurchaseOrder objOrder = GetPurchaseOrder(((PurchaseOrder)(recordList[j])).internalId);
                                        object objCustomReseller = GetValue("custbodytbfcustomerid", objOrder.customFieldList);
                                        if (objCustomReseller.ToString().Equals("CST100WA"))
                                        {
                                            RecordRef objRef = new RecordRef();
                                            objRef.internalId = objOrder.internalId;
                                            objRef.type = RecordType.purchaseOrder;
                                            objRef.typeSpecified = true;
                                            WriteResponse writeRes = _service.delete(objRef);
                                            if (writeRes.status.isSuccess)
                                            {
                                                Console.WriteLine("delete: " + objOrder.tranId);
                                            }
                                            else
                                            {
                                                Console.WriteLine(getStatusDetails(writeRes.status));
                                            }

                                        }
                                    }
                                    catch (Exception objExc)
                                    {
                                        Console.WriteLine(objExc.ToString());
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    strItem = "error";
                }
            }
            catch (Exception objExc)
            {
                Console.WriteLine(objExc.ToString());
            }
        }

        public static void DeleteInvoices()
        {
            string strItem = "";
            try
            {
                TransactionSearch xactionSearch = new TransactionSearch();

                SearchEnumMultiSelectField searchSalesOrderField = new SearchEnumMultiSelectField();
                searchSalesOrderField.@operator = SearchEnumMultiSelectFieldOperator.anyOf;
                searchSalesOrderField.operatorSpecified = true;
                System.String[] soStringArray = new System.String[1];
                soStringArray[0] = "_invoice";
                searchSalesOrderField.searchValue = soStringArray;

                TransactionSearchBasic xactionBasic = new TransactionSearchBasic();
                xactionBasic.type = searchSalesOrderField;
                xactionBasic.dateCreated = new SearchDateField();
                xactionBasic.dateCreated.@operator = SearchDateFieldOperator.within;
                xactionBasic.dateCreated.operatorSpecified = true;
                xactionBasic.dateCreated.searchValue = new DateTime(2011, 12, 30);
                xactionBasic.dateCreated.searchValueSpecified = true;
                xactionBasic.dateCreated.searchValue2 = new DateTime(2011, 12, 31);
                xactionBasic.dateCreated.searchValue2Specified = true;

                //xactionBasic.customFieldList = new SearchCustomField[1];
                //SearchStringCustomField objNumber = new SearchStringCustomField();
                //objNumber.internalId = "custbodycst_ordernumber";
                //objNumber.@operator = SearchStringFieldOperator.notEmpty;
                //objNumber.operatorSpecified = true;
                //xactionBasic.customFieldList[0] = objNumber;
                xactionSearch.basic = xactionBasic;

                SearchResult objSearchResult = _service.search(xactionSearch);
                if (objSearchResult.status != null && objSearchResult.status.isSuccess)
                {
                    if (objSearchResult.recordList != null)
                    {
                        Record[] recordList;
                        for (int i = 1; i <= objSearchResult.totalPages; i++)
                        {
                            recordList = objSearchResult.recordList;

                            for (int j = 0; j < recordList.Length; j++)
                            {
                                if (recordList[j] is Invoice)
                                {
                                    try
                                    {
                                        //Invoice objOrder = (Invoice)GetEntity(((Invoice)(recordList[j])).internalId, RecordType.invoice);
                                        Invoice objOrder = (Invoice)(recordList[j]);
                                        object objCustomReseller = GetValue("custbodytbfcustomerid", objOrder.customFieldList);

                                        if (objCustomReseller != null && (objCustomReseller.ToString().Equals("CST100WA") || objCustomReseller.ToString().Equals("1")))
                                        {
                                            RecordRef objRef = new RecordRef();
                                            objRef.internalId = objOrder.internalId;
                                            objRef.type = RecordType.invoice;
                                            objRef.typeSpecified = true;
                                            WriteResponse writeRes = _service.delete(objRef);
                                            if (writeRes.status.isSuccess)
                                            {
                                                Console.WriteLine("delete: " + objOrder.tranId);
                                            }
                                            else
                                            {
                                                Console.WriteLine(getStatusDetails(writeRes.status));
                                            }

                                        }
                                    }
                                    catch (Exception objExc)
                                    {
                                        Console.WriteLine(objExc.ToString());
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    strItem = "error";
                }
            }
            catch (Exception objExc)
            {
                Console.WriteLine(objExc.ToString());
            }
        }

        public static void DeleteItemFullfitment()
        {
            string strItem = "";
            try
            {
                TransactionSearch xactionSearch = new TransactionSearch();
                //xactionSearch.basic
                //SearchMultiSelectField entity = new SearchMultiSelectField();
                //entity.@operator = SearchMultiSelectFieldOperator.anyOf;
                //entity.operatorSpecified = true;
                //RecordRef custRecordRef = new RecordRef();
                //custRecordRef.internalId = strCustomerId;
                //custRecordRef.type = RecordType.customer;
                //custRecordRef.typeSpecified = true;
                //RecordRef[] custRecordRefArray = new RecordRef[1];
                //custRecordRefArray[0] = custRecordRef;
                //entity.searchValue = custRecordRefArray;

                SearchEnumMultiSelectField searchSalesOrderField = new SearchEnumMultiSelectField();
                searchSalesOrderField.@operator = SearchEnumMultiSelectFieldOperator.anyOf;
                searchSalesOrderField.operatorSpecified = true;
                System.String[] soStringArray = new System.String[1];
                soStringArray[0] = "_itemFulfillment";
                searchSalesOrderField.searchValue = soStringArray;


                TransactionSearchBasic xactionBasic = new TransactionSearchBasic();
                xactionBasic.type = searchSalesOrderField;
                xactionBasic.dateCreated = new SearchDateField();
                xactionBasic.dateCreated.@operator = SearchDateFieldOperator.within;
                xactionBasic.dateCreated.operatorSpecified = true;
                xactionBasic.dateCreated.searchValue = new DateTime(2011, 12, 30);
                xactionBasic.dateCreated.searchValueSpecified = true;
                xactionBasic.dateCreated.searchValue2 = new DateTime(2011, 12, 31);
                xactionBasic.dateCreated.searchValue2Specified = true;

                xactionSearch.basic = xactionBasic;

                SearchResult objSearchResult = _service.search(xactionSearch);
                if (objSearchResult.status != null && objSearchResult.status.isSuccess)
                {
                    if (objSearchResult.recordList != null)
                    {
                        Record[] recordList;
                        for (int i = 1; i <= objSearchResult.totalPages; i++)
                        {
                            recordList = objSearchResult.recordList;

                            for (int j = 0; j < recordList.Length; j++)
                            {
                                if (recordList[j] is ItemFulfillment)
                                {
                                    try
                                    {
                                        ItemFulfillment objOrder = (ItemFulfillment)(recordList[j]);
                                        object objCustomReseller = GetValue("custbodytbfcustomerid", objOrder.customFieldList);
                                        if (objCustomReseller != null && (objCustomReseller.ToString().Equals("CST100WA") || objCustomReseller.ToString().Equals("1")))
                                        {

                                            RecordRef objRef = new RecordRef();
                                            objRef.internalId = ((ItemFulfillment)recordList[j]).internalId;
                                            objRef.type = RecordType.itemFulfillment;
                                            objRef.typeSpecified = true;
                                            WriteResponse writeRes = _service.delete(objRef);
                                            if (writeRes.status.isSuccess)
                                            {
                                                Console.WriteLine("delete: " + objRef.internalId);
                                            }
                                            else
                                            {
                                                Console.WriteLine(getStatusDetails(writeRes.status));
                                            }
                                        }
                                        
                                    }
                                    catch (Exception objExc)
                                    {
                                        Console.WriteLine(objExc.ToString());
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    strItem = "error";
                }
            }
            catch (Exception objExc)
            {
                Console.WriteLine(objExc.ToString());
            }
        }

        private static PurchaseOrder GetPurchaseOrder(string strOrderId)
        {
            PurchaseOrder objOrder = new PurchaseOrder();
            try
            {
                RecordRef objRef = new RecordRef();
                objRef.internalId = strOrderId;
                objRef.type = RecordType.purchaseOrder;
                objRef.typeSpecified = true;
                ReadResponse objReadResponse = _service.get(objRef);
                objOrder = (PurchaseOrder)(objReadResponse.record);
            }
            catch (Exception)
            {
            }
            return objOrder;
        }

        private static SalesOrder GetSalesOrder(string strOrderId)
        {
            SalesOrder objOrder = new SalesOrder();
            try
            {
                RecordRef objRef = new RecordRef();
                objRef.internalId = strOrderId;
                objRef.type = RecordType.salesOrder;
                objRef.typeSpecified = true;
                ReadResponse objReadResponse = _service.get(objRef);
                objOrder = (SalesOrder)(objReadResponse.record);
            }
            catch (Exception)
            {
            }
            return objOrder;
        }

        private static object GetValue(string strId, CustomFieldRef[] arrFields)
        {
            object objValue = null;

            foreach (CustomFieldRef objField in arrFields)
            {
                if (objField is BooleanCustomFieldRef)
                {
                    BooleanCustomFieldRef objCustomFieldRef = (BooleanCustomFieldRef)objField;
                    if (objCustomFieldRef.internalId.Equals(strId))
                    {
                        objValue = objCustomFieldRef.value;
                        break;
                    }
                }
                else if (objField is StringCustomFieldRef)
                {
                    StringCustomFieldRef objCustomFieldRef = (StringCustomFieldRef)objField;
                    if (objCustomFieldRef.internalId.Equals(strId))
                    {
                        objValue = objCustomFieldRef.value;
                        break;
                    }
                }
                else if (objField is SelectCustomFieldRef)
                {
                    SelectCustomFieldRef objCustomFieldRef = (SelectCustomFieldRef)objField;
                    if (objCustomFieldRef.internalId.Equals(strId))
                    {
                        if (objCustomFieldRef.value.name != null)
                        {
                            objValue = objCustomFieldRef.value.name;
                        }
                        else
                        {
                            objValue = objCustomFieldRef.value.internalId;
                        }
                        break;
                    }
                }
                else if (objField is MultiSelectCustomFieldRef)
                {
                    MultiSelectCustomFieldRef objCustomFieldRef = (MultiSelectCustomFieldRef)objField;
                    if (objCustomFieldRef.internalId.Equals(strId))
                    {
                        objValue = objCustomFieldRef.value[0].internalId;
                        break;
                    }
                }
                else if (objField is DateCustomFieldRef)
                {
                    DateCustomFieldRef objCustomFieldRef = (DateCustomFieldRef)objField;
                    if (objCustomFieldRef.internalId.Equals(strId))
                    {
                        objValue = objCustomFieldRef.value;
                        break;
                    }
                }
                else if (objField is DoubleCustomFieldRef)
                {
                    DoubleCustomFieldRef objCustomFieldRef = (DoubleCustomFieldRef)objField;
                    if (objCustomFieldRef.internalId.Equals(strId))
                    {
                        objValue = objCustomFieldRef.value;
                        break;
                    }
                }
                else if (objField is LongCustomFieldRef)
                {
                    LongCustomFieldRef objCustomFieldRef = (LongCustomFieldRef)objField;
                    if (objCustomFieldRef.internalId.Equals(strId))
                    {
                        objValue = objCustomFieldRef.value;
                        break;
                    }
                }
            }
            return objValue;
        }

        public static string GetOrderId(string strCustomerId, double numAmount)
        {
            string strItem = "";
            try
            {
                TransactionSearch xactionSearch = new TransactionSearch();
                SearchMultiSelectField entity = new SearchMultiSelectField();
                entity.@operator = SearchMultiSelectFieldOperator.anyOf;
                entity.operatorSpecified = true;
                RecordRef custRecordRef = new RecordRef();
                custRecordRef.internalId = strCustomerId;
                custRecordRef.type = RecordType.customer;
                custRecordRef.typeSpecified = true;
                RecordRef[] custRecordRefArray = new RecordRef[1];
                custRecordRefArray[0] = custRecordRef;
                entity.searchValue = custRecordRefArray;

                SearchEnumMultiSelectField searchSalesOrderField = new SearchEnumMultiSelectField();
                searchSalesOrderField.@operator = SearchEnumMultiSelectFieldOperator.anyOf;
                searchSalesOrderField.operatorSpecified = true;
                System.String[] soStringArray = new System.String[1];
                soStringArray[0] = "_salesOrder";
                searchSalesOrderField.searchValue = soStringArray;

                TransactionSearchBasic xactionBasic = new TransactionSearchBasic();
                xactionBasic.type = searchSalesOrderField;
                xactionBasic.entity = entity;
                xactionBasic.status = new SearchEnumMultiSelectField();
                xactionBasic.status.@operator = SearchEnumMultiSelectFieldOperator.anyOf;
                xactionBasic.status.operatorSpecified = true;
                xactionBasic.status.searchValue = new string[] { "_salesOrderPendingApproval" };
                xactionBasic.amount = new SearchDoubleField();
                xactionBasic.amount.@operator = SearchDoubleFieldOperator.equalTo;
                xactionBasic.amount.operatorSpecified = true;
                xactionBasic.amount.searchValue = numAmount;
                xactionBasic.amount.searchValueSpecified = true;

                xactionSearch.basic = xactionBasic;

                SearchResult objSearchResult = _service.search(xactionSearch);
                if (objSearchResult.recordList != null)
                {
                    Record[] recordList;
                    for (int i = 1; i <= objSearchResult.totalPages; i++)
                    {
                        recordList = objSearchResult.recordList;

                        for (int j = 0; j < recordList.Length; j++)
                        {
                            if (recordList[j] is SalesOrder)
                            {
                                SalesOrder so = (SalesOrder)(recordList[j]);
                                strItem = so.internalId;
                                return strItem;
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {
            }
            return strItem;
        }

        public static string AddSalesOrder(int numReseller, RecordRef objCustomer, string strPoNumber, string strOrderNumber, DateTime datOrderDate,
            DateTime datInsertDate, DateTime datEventDate, ArrayList arrItems, string strUserId, DateTime datInvoiceDate, string strClass, out string strError)
        {
            strError = "";
            string strId = "";
            SalesOrder objSalesOrder = new SalesOrder();
            objSalesOrder.entity = objCustomer;
            objSalesOrder.tranDate = new System.DateTime();
            objSalesOrder.orderStatus = SalesOrderOrderStatus._pendingFulfillment;
            objSalesOrder.externalId = numReseller+"_"+strOrderNumber;

            objSalesOrder.customForm = new RecordRef();
            objSalesOrder.customForm.internalId = GlobalSettings.Default.nsOrderForm;
            objSalesOrder.customForm.type = RecordType.account;
            objSalesOrder.customForm.typeSpecified = true;

            //objSalesOrder.salesRep = new RecordRef();
            //objSalesOrder.salesRep.internalId = strUserId;
            //objSalesOrder.salesRep.type = RecordType.employee;
            //objSalesOrder.salesRep.typeSpecified = true;

            SelectCustomFieldRef objReseller = new SelectCustomFieldRef();
            objReseller.internalId = "custbodytbfcustomerid";
            objReseller.value = new ListOrRecordRef();
            objReseller.value.internalId = numReseller.ToString();

            StringCustomFieldRef objOrderNumber = new StringCustomFieldRef();
            objOrderNumber.internalId = "custbodycst_ordernumber";
            objOrderNumber.value = strOrderNumber;

            StringCustomFieldRef objPONumber = new StringCustomFieldRef();
            objPONumber.internalId = "custbodycostcoponumber";
            objPONumber.value = strPoNumber;

            DateCustomFieldRef objOrderDate = new DateCustomFieldRef();
            objOrderDate.internalId = "custbodycst_orderdate";
            objOrderDate.value = datOrderDate;

            DateCustomFieldRef objInvoiceDate = new DateCustomFieldRef();
            objInvoiceDate.internalId = "custbody_vf_invoice_date";
            objInvoiceDate.value = datInvoiceDate;

            DateCustomFieldRef objInsertDate = new DateCustomFieldRef();
            objInsertDate.internalId = "custbodycst_insertdate";
            objInsertDate.value = datInsertDate;

            //DateCustomFieldRef objEventDate = new DateCustomFieldRef();
            //objEventDate.internalId = "custbody_wedding_date";
            //objEventDate.value = datEventDate;

            object[] arrFirstItem = (object[])arrItems[0];

            DateCustomFieldRef objDeliveryDate = new DateCustomFieldRef();
            objDeliveryDate.internalId = "custbody_delivery_date";
            objDeliveryDate.value = (DateTime)arrFirstItem[3];

            DateCustomFieldRef objShippingDate = new DateCustomFieldRef();
            objShippingDate.internalId = "custbody_shipping_date";
            objShippingDate.value = (DateTime)arrFirstItem[2];

            StringCustomFieldRef objMessage = new StringCustomFieldRef();
            objMessage.internalId = "custbodycst_giftmessage";
            objMessage.value = "";

            if (objDeliveryDate.value != DateTime.MaxValue)
            {
                objSalesOrder.customFieldList = new CustomFieldRef[] { objReseller, objOrderNumber, objPONumber, objOrderDate, objInsertDate, objMessage, objInvoiceDate };
            }
            else
            {
                objSalesOrder.customFieldList = new CustomFieldRef[] { objReseller, objOrderNumber, objPONumber, objOrderDate, objInsertDate, objMessage, objInvoiceDate };
            }
            objSalesOrder.otherRefNum = strPoNumber;

            //arrItems.Add(new object[] { strProductId, numQty, strShipDate, strDeliDate, arrItems.IndexOf(strShipType) });

            SalesOrderItem[] salesOrderItemArray = new SalesOrderItem[arrItems.Count];
            for (int i = 0; i < salesOrderItemArray.Length; i++)
            {
                object[] arrColumns = (object[])arrItems[i];

                RecordRef item = new RecordRef();
                item.type = RecordType.nonInventoryResaleItem;
                item.typeSpecified = true;
                item.internalId = arrColumns[0].ToString();

                salesOrderItemArray[i] = new SalesOrderItem();
                salesOrderItemArray[i].item = item;

                SelectCustomFieldRef objShippingType = new SelectCustomFieldRef();
                objShippingType.internalId = "custcolcst_shippingtype";
                objShippingType.value = new ListOrRecordRef();
                objShippingType.value.internalId = arrColumns[4].ToString();
                if (objShippingType.value.internalId.Equals("-1") || objShippingType.value.internalId.Equals("0"))
                {
                    objShippingType.value.internalId = "7";
                }

                SelectCustomFieldRef objShippingCourier = new SelectCustomFieldRef();
                objShippingCourier.internalId = "custcolcst_courier";
                objShippingCourier.value = new ListOrRecordRef();
                objShippingCourier.value.internalId = arrColumns[7].ToString();
                if (objShippingCourier.value.internalId.Equals("-1") || objShippingCourier.value.internalId.Equals("0"))
                {
                    objShippingCourier.value.internalId = "1"; 
                    if (objShippingType.value.internalId.Equals("8") || objShippingType.value.internalId.Equals("9"))
                    {
                        objShippingCourier.value.internalId = "2";
                    }
                    if (objShippingType.value.internalId.Equals("5") || objShippingType.value.internalId.Equals("6") || objShippingType.value.internalId.Equals("7"))
                    {
                        objShippingCourier.value.internalId = "3";
                    }
                    
                }

                DateCustomFieldRef objShipDate = new DateCustomFieldRef();
                objShipDate.internalId = "custcolcst_shippingdate";
                objShipDate.value = (DateTime)arrColumns[2];

                DateCustomFieldRef objDelivDate = new DateCustomFieldRef();
                objDelivDate.internalId = "custcolcst_deliverydate";
                objDelivDate.value = (DateTime)arrColumns[3];

                StringCustomFieldRef objPo = new StringCustomFieldRef();
                objPo.internalId = "custcolos_ponumber";
                objPo.value = arrColumns[5].ToString();

                StringCustomFieldRef objColMessage = new StringCustomFieldRef();
                objColMessage.internalId = "custcolcst_message";
                objColMessage.value = "";

                if (arrColumns.Length >= 7)
                {
                    objColMessage.value = arrColumns[6].ToString();
                    objMessage.value = arrColumns[6].ToString();
                }

                if (objDelivDate.value != DateTime.MaxValue)
                {
                    salesOrderItemArray[i].customFieldList = new CustomFieldRef[] { objDelivDate, objShipDate, objShippingType, objPo, objShippingCourier, objColMessage };
                }
                else
                {
                    salesOrderItemArray[i].customFieldList = new CustomFieldRef[] { objShipDate, objShippingType, objPo, objShippingCourier, objColMessage };
                }

                System.Double quantity = System.Double.Parse(arrColumns[1].ToString());
                salesOrderItemArray[i].quantity = quantity;
                salesOrderItemArray[i].quantitySpecified = true;
            }

            SalesOrderItemList salesOrderItemList = new SalesOrderItemList();
            salesOrderItemList.item = salesOrderItemArray;

            objSalesOrder.itemList = salesOrderItemList;
            WriteResponse writeRes = _service.add(objSalesOrder);
            if (writeRes.status.isSuccess)
            {
                strId = ((RecordRef)writeRes.baseRef).internalId;
                //_out.writeLn("\nSales order created successfully\nSales Order Internal ID=" + ((RecordRef)writeRes.baseRef).internalId);
            }
            else
            {
                strError = getStatusDetails(writeRes.status);
            }
            return strId;
        }

        public static string AddSalesOrder1800(int numReseller, Customer objCustomer, string strVendorCode, string strPoNumber, string strOrderNumber, DateTime datOrderDate,
            DateTime datInsertDate, DateTime datEventDate, ArrayList arrItems, string strUserId, out string strError)
        {
            strError = "";
            string strId = "";
            SalesOrder objSalesOrder = new SalesOrder();
            objSalesOrder.entity = new RecordRef();
            objSalesOrder.entity.type = RecordType.customer;
            objSalesOrder.entity.typeSpecified = true;
            objSalesOrder.entity.externalId = objCustomer.externalId;
            objSalesOrder.tranDate = new System.DateTime();
            objSalesOrder.orderStatus = SalesOrderOrderStatus._pendingFulfillment;
            objSalesOrder.externalId = "so1800_"+strOrderNumber;

            objSalesOrder.customForm = new RecordRef();
            objSalesOrder.customForm.internalId = GlobalSettings.Default.nsOrderForm;
            objSalesOrder.customForm.type = RecordType.account;
            objSalesOrder.customForm.typeSpecified = true;

            //objSalesOrder.salesRep = new RecordRef();
            //objSalesOrder.salesRep.internalId = strUserId;
            //objSalesOrder.salesRep.type = RecordType.employee;
            //objSalesOrder.salesRep.typeSpecified = true;

            SelectCustomFieldRef objReseller = new SelectCustomFieldRef();
            objReseller.internalId = "custbodytbfcustomerid";
            objReseller.value = new ListOrRecordRef();
            objReseller.value.internalId = numReseller.ToString();

            StringCustomFieldRef objVendorCode = new StringCustomFieldRef();
            objVendorCode.internalId = "custbody_vfvendorcode";
            objVendorCode.value = strVendorCode;

            StringCustomFieldRef objOrderNumber = new StringCustomFieldRef();
            objOrderNumber.internalId = "custbodycst_ordernumber";
            objOrderNumber.value = strOrderNumber;

            StringCustomFieldRef objPONumber = new StringCustomFieldRef();
            objPONumber.internalId = "custbodycostcoponumber";
            objPONumber.value = strPoNumber;

            DateCustomFieldRef objOrderDate = new DateCustomFieldRef();
            objOrderDate.internalId = "custbodycst_orderdate";
            objOrderDate.value = datOrderDate;

            DateCustomFieldRef objInsertDate = new DateCustomFieldRef();
            objInsertDate.internalId = "custbodycst_insertdate";
            objInsertDate.value = datInsertDate;

            //DateCustomFieldRef objEventDate = new DateCustomFieldRef();
            //objEventDate.internalId = "custbody_wedding_date";
            //objEventDate.value = datEventDate;

            object[] arrFirstItem = (object[])arrItems[0];

            DateCustomFieldRef objDeliveryDate = new DateCustomFieldRef();
            objDeliveryDate.internalId = "custbody_delivery_date";
            objDeliveryDate.value = (DateTime)arrFirstItem[3];

            DateCustomFieldRef objShippingDate = new DateCustomFieldRef();
            objShippingDate.internalId = "custbody_shipping_date";
            objShippingDate.value = (DateTime)arrFirstItem[2];

            StringCustomFieldRef objMessage = new StringCustomFieldRef();
            objMessage.internalId = "custbodycst_giftmessage";
            objMessage.value = "";

            if (objDeliveryDate.value != DateTime.MaxValue)
            {
                objSalesOrder.customFieldList = new CustomFieldRef[] { objReseller, objOrderNumber, objPONumber, objOrderDate, objInsertDate, objMessage, objVendorCode };
            }
            else
            {
                objSalesOrder.customFieldList = new CustomFieldRef[] { objReseller, objOrderNumber, objPONumber, objOrderDate, objInsertDate, objMessage, objVendorCode };
            }
            objSalesOrder.otherRefNum = strPoNumber;

            //arrItems.Add(new object[] { strProductId, numQty, strShipDate, strDeliDate, arrItems.IndexOf(strShipType) });

            SalesOrderItem[] salesOrderItemArray = new SalesOrderItem[arrItems.Count];
            for (int i = 0; i < salesOrderItemArray.Length; i++)
            {
                object[] arrColumns = (object[])arrItems[i];

                RecordRef item = new RecordRef();
                item.type = RecordType.nonInventoryResaleItem;
                item.typeSpecified = true;
                item.internalId = arrColumns[0].ToString();

                salesOrderItemArray[i] = new SalesOrderItem();
                salesOrderItemArray[i].item = item;

                SelectCustomFieldRef objShippingType = new SelectCustomFieldRef();
                objShippingType.internalId = "custcolcst_shippingtype";
                objShippingType.value = new ListOrRecordRef();
                objShippingType.value.internalId = arrColumns[4].ToString();
                if (objShippingType.value.internalId.Equals("-1"))
                {
                    objShippingType.value.internalId = "7";
                }

                SelectCustomFieldRef objShippingCourier = new SelectCustomFieldRef();
                objShippingCourier.internalId = "custcolcst_courier";
                objShippingCourier.value = new ListOrRecordRef();
                objShippingCourier.value.internalId = "1";
                if (objShippingType.value.internalId.Equals("8") || objShippingType.value.internalId.Equals("9"))
                {
                    objShippingCourier.value.internalId = "2";
                }
                if (objShippingType.value.internalId.Equals("5") || objShippingType.value.internalId.Equals("6") || objShippingType.value.internalId.Equals("7"))
                {
                    objShippingCourier.value.internalId = "3";
                }

                DateCustomFieldRef objShipDate = new DateCustomFieldRef();
                objShipDate.internalId = "custcolcst_shippingdate";
                objShipDate.value = (DateTime)arrColumns[2];

                DateCustomFieldRef objDelivDate = new DateCustomFieldRef();
                objDelivDate.internalId = "custcolcst_deliverydate";
                objDelivDate.value = (DateTime)arrColumns[3];

                StringCustomFieldRef objPo = new StringCustomFieldRef();
                objPo.internalId = "custcolos_ponumber";
                objPo.value = arrColumns[5].ToString();

                StringCustomFieldRef objColMessage = new StringCustomFieldRef();
                objColMessage.internalId = "custcolcst_message";
                objColMessage.value = "";

                if (arrColumns.Length == 7)
                {
                    objColMessage.value = arrColumns[6].ToString();
                    objMessage.value = arrColumns[6].ToString();
                }

                if (objDelivDate.value != DateTime.MaxValue)
                {
                    salesOrderItemArray[i].customFieldList = new CustomFieldRef[] { objDelivDate, objShipDate, objShippingType, objPo, objShippingCourier, objColMessage };
                }
                else
                {
                    salesOrderItemArray[i].customFieldList = new CustomFieldRef[] { objShipDate, objShippingType, objPo, objShippingCourier, objColMessage };
                }

                System.Double quantity = System.Double.Parse(arrColumns[1].ToString());
                salesOrderItemArray[i].quantity = quantity;
                salesOrderItemArray[i].quantitySpecified = true;
            }

            SalesOrderItemList salesOrderItemList = new SalesOrderItemList();
            salesOrderItemList.item = salesOrderItemArray;

            objSalesOrder.itemList = salesOrderItemList;
            //WriteResponse writeRes = _service.add(objSalesOrder);
            WriteResponse[] writeRes = _service.addList(new Record[] { objCustomer, objSalesOrder });
            if (writeRes[0].status.isSuccess)
            {
                strId = ((RecordRef)writeRes[0].baseRef).internalId;
                //_out.writeLn("\nSales order created successfully\nSales Order Internal ID=" + ((RecordRef)writeRes.baseRef).internalId);
            }
            else
            {
                strError = getStatusDetails(writeRes[0].status);
            }
            if (writeRes[1].status.isSuccess)
            {
                strId = ((RecordRef)writeRes[1].baseRef).internalId;
                //_out.writeLn("\nSales order created successfully\nSales Order Internal ID=" + ((RecordRef)writeRes.baseRef).internalId);
            }
            else
            {
                strError = getStatusDetails(writeRes[1].status);
            }
            return strId;
        }

        public static string UpdateSalesOrder(SalesOrder objOrder, ArrayList arrItems, out string strError)
        {
            strError = "";
            string strId = "";
            SalesOrder objSalesOrder = new SalesOrder();
            objSalesOrder.internalId = objOrder.internalId;
            objSalesOrder.customForm = new RecordRef();
            objSalesOrder.customForm.internalId = GlobalSettings.Default.nsOrderForm;
            objSalesOrder.customForm.type = RecordType.account;
            objSalesOrder.customForm.typeSpecified = true;

            SalesOrderItem[] salesOrderItemArray = new SalesOrderItem[arrItems.Count + objOrder.itemList.item.Length];
            for (int i = 0; i < salesOrderItemArray.Length; i++)
            {
                if (i < objOrder.itemList.item.Length)
                {
                    salesOrderItemArray[i] = new SalesOrderItem();
                    salesOrderItemArray[i].item = objOrder.itemList.item[i].item;
                    salesOrderItemArray[i].line = objOrder.itemList.item[i].line;
                    salesOrderItemArray[i].lineSpecified = true;

                    SelectCustomFieldRef objShippingType = new SelectCustomFieldRef();
                    objShippingType.internalId = "custcolcst_shippingtype";
                    objShippingType.value = new ListOrRecordRef();
                    objShippingType.value.internalId = "7";

                    //salesOrderItemArray[i].customFieldList = new CustomFieldRef[] { objShippingType };
                }
                else
                {
                    object[] arrColumns = (object[])arrItems[i - objOrder.itemList.item.Length];

                    RecordRef item = new RecordRef();
                    item.type = RecordType.nonInventoryResaleItem;
                    item.typeSpecified = true;
                    item.internalId = arrColumns[0].ToString();
                    salesOrderItemArray[i] = new SalesOrderItem();
                    salesOrderItemArray[i].item = item;

                    SelectCustomFieldRef objShippingType = new SelectCustomFieldRef();
                    objShippingType.internalId = "custcolcst_shippingtype";
                    objShippingType.value = new ListOrRecordRef();
                    objShippingType.value.internalId = arrColumns[4].ToString();
                    if (objShippingType.value.internalId.Equals("-1"))
                    {
                        objShippingType.value.internalId = "7";
                    }

                    DateCustomFieldRef objShipDate = new DateCustomFieldRef();
                    objShipDate.internalId = "custcolcst_shippingdate";
                    objShipDate.value = (DateTime)arrColumns[2];

                    DateCustomFieldRef objDelivDate = new DateCustomFieldRef();
                    objDelivDate.internalId = "custcolcst_deliverydate";
                    objDelivDate.value = (DateTime)arrColumns[3];

                    StringCustomFieldRef objPo = new StringCustomFieldRef();
                    objPo.internalId = "custcolos_ponumber";
                    objPo.value = arrColumns[5].ToString();

                    if (objDelivDate.value != DateTime.MaxValue)
                    {
                        salesOrderItemArray[i].customFieldList = new CustomFieldRef[] { objDelivDate, objShipDate, objShippingType, objPo };
                    }
                    else
                    {
                        salesOrderItemArray[i].customFieldList = new CustomFieldRef[] { objShipDate, objShippingType, objPo };
                    }

                    System.Double quantity = System.Double.Parse(arrColumns[1].ToString());
                    salesOrderItemArray[i].quantity = quantity;
                    salesOrderItemArray[i].quantitySpecified = true;
                }
            }

            SalesOrderItemList salesOrderItemList = new SalesOrderItemList();
            salesOrderItemList.item = salesOrderItemArray;

            objSalesOrder.itemList = salesOrderItemList;
            WriteResponse writeRes = _service.update(objSalesOrder);
            if (writeRes.status.isSuccess)
            {
                strId = ((RecordRef)writeRes.baseRef).internalId;
                //_out.writeLn("\nSales order created successfully\nSales Order Internal ID=" + ((RecordRef)writeRes.baseRef).internalId);
            }
            else
            {
                strError = getStatusDetails(writeRes.status);
            }
            return strId;
        }


        public static string AddSalesOrderMassive(RecordRef objCustomer, string strCategory, string strPoNumber, string strOrderNumber, DateTime datOrderDate,
    DateTime datInsertDate, DateTime datEventDate, bool booConfirmed, bool booCases, string strBride, string strGroom, ArrayList arrItems, string strUserId)
        {
            string strId = "";
            SalesOrder objSalesOrder = new SalesOrder();
            objSalesOrder.entity = objCustomer;
            objSalesOrder.tranDate = new System.DateTime();
            objSalesOrder.orderStatus = SalesOrderOrderStatus._pendingFulfillment;

            //objSalesOrder.salesRep = new RecordRef();
            //objSalesOrder.salesRep.internalId = strUserId;
            //objSalesOrder.salesRep.type = RecordType.employee;
            //objSalesOrder.salesRep.typeSpecified = true;

            SelectCustomFieldRef objReseller = new SelectCustomFieldRef();
            objReseller.internalId = "custbodytbfcustomerid";
            objReseller.value = new ListOrRecordRef();
            objReseller.value.internalId = "1";

            BooleanCustomFieldRef objCategory = new BooleanCustomFieldRef();
            objCategory.internalId = "custbodycst_ordercategory";
            objCategory.value = bool.Parse(strCategory);

            BooleanCustomFieldRef objConfirmed = new BooleanCustomFieldRef();
            objConfirmed.internalId = "custbodycst_orderconfirmed";
            objConfirmed.value = booConfirmed;

            BooleanCustomFieldRef objCases = new BooleanCustomFieldRef();
            objCases.internalId = "custbodycst_cases";
            objCases.value = booCases;

            StringCustomFieldRef objOrderNumber = new StringCustomFieldRef();
            objOrderNumber.internalId = "custbodycst_ordernumber";
            objOrderNumber.value = strOrderNumber;

            StringCustomFieldRef objPONumber = new StringCustomFieldRef();
            objPONumber.internalId = "custbodycostcoponumber";
            objPONumber.value = strPoNumber;

            StringCustomFieldRef objGroom = new StringCustomFieldRef();
            objGroom.internalId = "custbody_groom";
            objGroom.value = strGroom;

            StringCustomFieldRef objBride = new StringCustomFieldRef();
            objBride.internalId = "custbody_bride";
            objBride.value = strBride;

            DateCustomFieldRef objOrderDate = new DateCustomFieldRef();
            objOrderDate.internalId = "custbodycst_orderdate";
            objOrderDate.value = datOrderDate;

            DateCustomFieldRef objInsertDate = new DateCustomFieldRef();
            objInsertDate.internalId = "custbodycst_insertdate";
            objInsertDate.value = datInsertDate;

            DateCustomFieldRef objEventDate = new DateCustomFieldRef();
            objEventDate.internalId = "custbody_wedding_date";
            objEventDate.value = datEventDate;

            object[] arrFirstItem = (object[])arrItems[0];

            DateCustomFieldRef objDeliveryDate = new DateCustomFieldRef();
            objDeliveryDate.internalId = "custbody_delivery_date";
            objDeliveryDate.value = (DateTime)arrFirstItem[2];

            DateCustomFieldRef objShippingDate = new DateCustomFieldRef();
            objShippingDate.internalId = "custbody_shipping_date";
            objShippingDate.value = (DateTime)arrFirstItem[3];

            if (datEventDate != DateTime.MaxValue)
            {
                objSalesOrder.customFieldList = new CustomFieldRef[] { objReseller, objConfirmed, objCases, objCategory, objGroom, objBride, objOrderNumber, objPONumber, objOrderDate, objInsertDate, objEventDate, objDeliveryDate, objShippingDate };
            }
            else
            {
                objSalesOrder.customFieldList = new CustomFieldRef[] { objReseller, objConfirmed, objCases, objCategory, objGroom, objBride, objOrderNumber, objPONumber, objOrderDate, objInsertDate, objDeliveryDate, objShippingDate };
            }
            objSalesOrder.otherRefNum = strPoNumber;

            SalesOrderItem[] salesOrderItemArray = new SalesOrderItem[arrItems.Count];
            for (int i = 0; i < salesOrderItemArray.Length; i++)
            {
                object[] arrColumns = (object[])arrItems[i];

                RecordRef item = new RecordRef();
                item.type = RecordType.nonInventoryResaleItem;
                item.typeSpecified = true;
                item.internalId = arrColumns[0].ToString();

                

                salesOrderItemArray[i] = new SalesOrderItem();
                salesOrderItemArray[i].item = item;

                StringCustomFieldRef objCustomInfo = new StringCustomFieldRef();
                objCustomInfo.internalId = "custcol_additional_info";
                objCustomInfo.value = arrColumns[8].ToString();

                StringCustomFieldRef objOriginal = new StringCustomFieldRef();
                objOriginal.internalId = "custcolcst_originalid";
                objOriginal.value = arrColumns[9].ToString();

                LongCustomFieldRef objOrderLine = new LongCustomFieldRef();
                objOrderLine.internalId = "custcolcst_orderline";
                objOrderLine.value = (int)arrColumns[6];

                SelectCustomFieldRef objShippingType = new SelectCustomFieldRef();
                objShippingType.internalId = "custcolcst_shippingtype";
                objShippingType.value = new ListOrRecordRef();
                switch (arrColumns[7].ToString())
                {
                    case "IPD":
                        objShippingType.value.internalId = "1";
                        break;
                    case "IP":
                        objShippingType.value.internalId = "2";
                        break;
                    case "Domestic":
                        objShippingType.value.internalId = "3";
                        break;
                    case "Fedex Next Day":
                        objShippingType.value.internalId = "4";
                        break;
                }

                if ((DateTime)arrColumns[4] == DateTime.MaxValue)
                {
                    DateCustomFieldRef objDelivDate = new DateCustomFieldRef();
                    objDelivDate.internalId = "custcolcst_deliverydate";
                    objDelivDate.value = (DateTime)arrColumns[2];

                    DateCustomFieldRef objShipDate = new DateCustomFieldRef();
                    objShipDate.internalId = "custcolcst_shippingdate";
                    objShipDate.value = (DateTime)arrColumns[3];

                    salesOrderItemArray[i].customFieldList = new CustomFieldRef[] { objDelivDate, objShipDate, objOrderLine, objShippingType, objCustomInfo, objOriginal };
                }
                else
                {
                    DateCustomFieldRef objDelivDate = new DateCustomFieldRef();
                    objDelivDate.internalId = "custcolcst_deliverydate";
                    objDelivDate.value = (DateTime)arrColumns[2];

                    DateCustomFieldRef objShipDate = new DateCustomFieldRef();
                    objShipDate.internalId = "custcolcst_shippingdate";
                    objShipDate.value = (DateTime)arrColumns[3];

                    DateCustomFieldRef objPrefDate = new DateCustomFieldRef();
                    objPrefDate.internalId = "custcolcst_preferredarrivaldate";
                    objPrefDate.value = (DateTime)arrColumns[4];

                    salesOrderItemArray[i].customFieldList = new CustomFieldRef[] { objDelivDate, objShipDate, objPrefDate, objOrderLine, objShippingType, objCustomInfo, objOriginal };
                }

                System.Double quantity = System.Double.Parse(arrColumns[1].ToString());
                salesOrderItemArray[i].quantity = quantity;
                salesOrderItemArray[i].quantitySpecified = true;
            }

            SalesOrderItemList salesOrderItemList = new SalesOrderItemList();
            salesOrderItemList.item = salesOrderItemArray;

            objSalesOrder.itemList = salesOrderItemList;
            WriteResponse writeRes = _service.add(objSalesOrder);
            if (writeRes.status.isSuccess)
            {
                strId = ((RecordRef)writeRes.baseRef).internalId;
                //_out.writeLn("\nSales order created successfully\nSales Order Internal ID=" + ((RecordRef)writeRes.baseRef).internalId);
            }
            else
            {
                throw new Exception("Orden no ingresada: " + getStatusDetails(writeRes.status));
                //_out.error(getStatusDetails(writeRes.status));
            }
            return strId;
        }


        public static void UpdateItem(NonInventoryResaleItem objItem)
        {
            try
            {
                WriteResponse writeRes = _service.update(objItem);
                if (writeRes.status.isSuccess)
                {
                    //_out.writeLn("\nSales order created successfully\nSales Order Internal ID=" + ((RecordRef)writeRes.baseRef).internalId);
                }
                else
                {
                    //_out.error(getStatusDetails(writeRes.status));
                }
            }
            catch (Exception)
            {
            }
        }

        public static string UpdateSalesOrderStatus(string strOrderId)
        {
            string strId = "";
            try
            {
                PurchaseOrder objPurchaseOrder = new PurchaseOrder();

                SalesOrder objSalesOrder = new SalesOrder();
                objSalesOrder.internalId = strOrderId;
                objSalesOrder.orderStatus = SalesOrderOrderStatus._pendingFulfillment;
                objSalesOrder.orderStatusSpecified = true;

                WriteResponse writeRes = _service.update(objSalesOrder);
                if (writeRes.status.isSuccess)
                {
                    //_out.writeLn("\nSales order created successfully\nSales Order Internal ID=" + ((RecordRef)writeRes.baseRef).internalId);
                }
                else
                {
                    //_out.error(getStatusDetails(writeRes.status));
                }

            }
            catch (Exception objExc)
            {
            }
            return strId;
        }

        public static string UpdateSalesOrder(SalesOrder objSalesOrder, SalesOrderItem objItem, ArrayList arrCustomFields)
        {
            string strId = "";
            try
            {
                SalesOrder objSOUpdated = new SalesOrder();
                objSOUpdated.internalId = objSalesOrder.internalId;
                objSOUpdated.itemList = new SalesOrderItemList();
                objSOUpdated.itemList.replaceAll = false;
                objSOUpdated.itemList.item = new SalesOrderItem[1];
                objSOUpdated.itemList.item[0] = new SalesOrderItem();
                //objSOUpdated.itemList.item[0] = objItem;
                objSOUpdated.itemList.item[0].item = objItem.item;
                objSOUpdated.itemList.item[0].line = objItem.line;
                objSOUpdated.itemList.item[0].lineSpecified = true;
                objSOUpdated.itemList.item[0].customFieldList = new CustomFieldRef[arrCustomFields.Count];
                for (int i = 0; i < arrCustomFields.Count; i++)
                {
                    objSOUpdated.itemList.item[0].customFieldList[i] = (CustomFieldRef)arrCustomFields[i];
                }

                WriteResponse writeRes = _service.update(objSOUpdated);
                if (writeRes.status.isSuccess)
                {
                    //_out.writeLn("\nSales order created successfully\nSales Order Internal ID=" + ((RecordRef)writeRes.baseRef).internalId);
                }
                else
                {
                    //_out.error(getStatusDetails(writeRes.status));
                }

            }
            catch (Exception objExc)
            {
            }
            return strId;
        }

        public static string UpdateSalesOrder(SalesOrder objSalesOrder)
        {
            string strId = "";
            try
            {
                SalesOrder objSOUpdated = new SalesOrder();
                objSOUpdated.internalId = objSalesOrder.internalId;
                objSOUpdated.itemList = new SalesOrderItemList();
                objSOUpdated.itemList.replaceAll = false;
                objSOUpdated.itemList.item = new SalesOrderItem[objSalesOrder.itemList.item.Length];
                for (int k = 0; k < objSOUpdated.itemList.item.Length; k++)
                {
                    objSOUpdated.itemList.item[k] = new SalesOrderItem();
                    objSOUpdated.itemList.item[k].item = objSalesOrder.itemList.item[k].item;
                    objSOUpdated.itemList.item[k].line = objSalesOrder.itemList.item[k].line;
                    objSOUpdated.itemList.item[k].lineSpecified = true;
                    objSOUpdated.itemList.item[k].customFieldList = new CustomFieldRef[objSalesOrder.itemList.item[k].customFieldList.Length];
                    for (int i = 0; i < objSOUpdated.itemList.item[k].customFieldList.Length; i++)
                    {
                        objSOUpdated.itemList.item[k].customFieldList[i] = objSalesOrder.itemList.item[k].customFieldList[i];
                    }
                }

                WriteResponse writeRes = _service.update(objSOUpdated);
                if (writeRes.status.isSuccess)
                {
                    //_out.writeLn("\nSales order created successfully\nSales Order Internal ID=" + ((RecordRef)writeRes.baseRef).internalId);
                }
                else
                {
                    //_out.error(getStatusDetails(writeRes.status));
                }

            }
            catch (Exception objExc)
            {
            }
            return strId;
        }

        public static string UpdateSalesOrderMessage(string strOrderId)
        {
            string strId = "";
            try
            {
                SalesOrder objSalesOrder = new SalesOrder();
                objSalesOrder.internalId = strOrderId;

                WriteResponse writeRes = _service.update(objSalesOrder);
                if (writeRes.status.isSuccess)
                {
                    //_out.writeLn("\nSales order created successfully\nSales Order Internal ID=" + ((RecordRef)writeRes.baseRef).internalId);
                }
                else
                {
                    //_out.error(getStatusDetails(writeRes.status));
                }

            }
            catch (Exception objExc)
            {
            }
            return strId;
        }

        public static RecordRef AddCustomer(string strParent, string strSalesRep, string strFirstName, string strLastName, string strEmail,
            string strBillTo, string strBillPhone, string strBillAddress1, string strBillAddress2, string strBillAddress3, string strBillCity, string strBillState, string strBillZip,
            string strShipTo, string strShipPhone, string strShipAddress1, string strShipAddress2, string strShipAddress3, string strShipCity, string strShipState, string strShipZip,
            out string strError)
        {
            strError = "";
            RecordRef objRecordRef = null;
            Customer objCustomer = new Customer();
            objCustomer.parent = new RecordRef();
            objCustomer.parent.internalId = strParent;
            objCustomer.parent.type = RecordType.customer;
            objCustomer.parent.typeSpecified = true;
            objCustomer.firstName = strFirstName;
            objCustomer.lastName = strLastName;
            objCustomer.email = strEmail;
            objCustomer.phone = strShipPhone;
            objCustomer.homePhone = strBillPhone;


            if(strParent == GlobalSettings.Default.nsParentCustomerTaxable)
            {
                objCustomer.taxable = true;
                objCustomer.taxableSpecified = true;
            }
            
            //objCustomer.altPhone = strShipPhone;

            objCustomer.salesRep = new RecordRef();
            objCustomer.salesRep.internalId = strSalesRep;
            objCustomer.salesRep.type = RecordType.employee;
            objCustomer.salesRep.typeSpecified = true;

            objCustomer.customForm = new RecordRef();
            objCustomer.customForm.internalId = GlobalSettings.Default.nsCustomerForm;
            objCustomer.customForm.type = RecordType.account;
            objCustomer.customForm.typeSpecified = true;

            //objCustomer.terms = new RecordRef();
            //objCustomer.terms.internalId = GlobalSettings.Default.nsCustomerTerms;
            //objCustomer.terms.type = RecordType.account;
            //objCustomer.terms.typeSpecified = true;

            objCustomer.addressbookList = new CustomerAddressbookList();
            objCustomer.addressbookList.addressbook = new CustomerAddressbook[2];
            objCustomer.addressbookList.addressbook[0] = new CustomerAddressbook();
            objCustomer.addressbookList.addressbook[0].defaultBilling = true;
            objCustomer.addressbookList.addressbook[0].defaultBillingSpecified = true;
            objCustomer.addressbookList.addressbook[0].defaultShipping = false;
            objCustomer.addressbookList.addressbook[0].defaultShippingSpecified = true;
            objCustomer.addressbookList.addressbook[0].isResidential = false;
            objCustomer.addressbookList.addressbook[0].isResidentialSpecified = true;
            objCustomer.addressbookList.addressbook[0].addr1 = strBillAddress1;
            objCustomer.addressbookList.addressbook[0].addr2 = strBillAddress2;
            objCustomer.addressbookList.addressbook[0].addr3 = strBillAddress3;
            objCustomer.addressbookList.addressbook[0].city = strBillCity;
            objCustomer.addressbookList.addressbook[0].zip = strBillZip;
            objCustomer.addressbookList.addressbook[0].state = strBillState;
            objCustomer.addressbookList.addressbook[0].attention = strBillTo;
            objCustomer.addressbookList.addressbook[0].label = strBillTo;
            objCustomer.addressbookList.addressbook[0].phone = strBillPhone;
            objCustomer.addressbookList.addressbook[0].country = Country._unitedStates;
            objCustomer.addressbookList.addressbook[0].countrySpecified = true;
            //objCustomer.addressbookList.addressbook[0].addressee = strBillTo + "\n" + strBillAddress1 + "\n" + strBillAddress2 + "\n" + strBillAddress3 + "\n" + strBillCity + " " + strBillState + " " + strBillZip + "\nUS\n" + strBillPhone;
            //objCustomer.addressbookList.addressbook[0].addrText = strBillTo + "\n" + strBillAddress1 + "\n" + strBillAddress2 + "\n" + strBillAddress3 + "\n" + strBillCity + " " + strBillState + " " + strBillZip + "\nUS\n" + strBillPhone;

            //if (objCustomer.addressbookList.addressbook[0].addressee.Length > 100)
            //{
            //    objCustomer.addressbookList.addressbook[0].addressee = strBillTo + "\n" + strBillAddress1 + "\n" + strBillAddress2 + "\n" + strBillCity + " " + strBillState + " " + strBillZip + "\nUS";
            //    while (objCustomer.addressbookList.addressbook[0].addressee.Length > 100)
            //    {
            //        objCustomer.addressbookList.addressbook[0].addressee = objCustomer.addressbookList.addressbook[0].addressee.Substring(1);
            //    }
            //}

            /*if (objCustomer.addressbookList.addressbook[0].addrText.Length > 100)
            {
                //objCustomer.addressbookList.addressbook[0].addrText = strBillTo + "\n" + strBillAddress1 + "\n" + strBillAddress2 + "\n" + strBillCity + " " + strBillState + " " + strBillZip + "\nUS";
                while (objCustomer.addressbookList.addressbook[0].addrText.Length > 100)
                {
                    //objCustomer.addressbookList.addressbook[0].addrText = objCustomer.addressbookList.addressbook[0].addrText.Substring(1);
                }
            }*/

            objCustomer.addressbookList.addressbook[1] = new CustomerAddressbook();
            objCustomer.addressbookList.addressbook[1].defaultBilling = false;
            objCustomer.addressbookList.addressbook[1].defaultBillingSpecified = true;
            objCustomer.addressbookList.addressbook[1].defaultShipping = true;
            objCustomer.addressbookList.addressbook[1].defaultShippingSpecified = true;
            objCustomer.addressbookList.addressbook[1].isResidential = false;
            objCustomer.addressbookList.addressbook[1].isResidentialSpecified = true;
            objCustomer.addressbookList.addressbook[1].addr1 = strShipAddress1;
            objCustomer.addressbookList.addressbook[1].addr2 = strShipAddress2;
            objCustomer.addressbookList.addressbook[1].addr3 = strShipAddress3;
            objCustomer.addressbookList.addressbook[1].city = strShipCity;
            objCustomer.addressbookList.addressbook[1].state = strShipState;
            objCustomer.addressbookList.addressbook[1].zip = strShipZip;
            objCustomer.addressbookList.addressbook[1].attention = strShipTo;
            objCustomer.addressbookList.addressbook[1].label = strShipTo;
            objCustomer.addressbookList.addressbook[1].phone = strShipPhone;
            objCustomer.addressbookList.addressbook[1].country = Country._unitedStates;
            objCustomer.addressbookList.addressbook[1].countrySpecified = true;
            //objCustomer.addressbookList.addressbook[1].addressee = strShipTo + "\n" + strShipAddress1 + "\n" + strShipAddress2 + "\n" + strShipAddress3 + "\n" + strShipCity + " " + strShipState + " " + strShipZip + "\nUS\n" + strShipPhone;
            //objCustomer.addressbookList.addressbook[1].addrText = strShipTo + "\n" + strShipAddress1 + "\n" + strShipAddress2 + "\n" + strShipAddress3 + "\n" + strShipCity + " " + strShipState + " " + strShipZip + "\nUS\n" + strShipPhone;

            //if (objCustomer.addressbookList.addressbook[1].addressee.Length > 100)
            //{
            //    objCustomer.addressbookList.addressbook[1].addressee = strShipTo + "\n" + strShipAddress1 + "\n" + strShipAddress2 + "\n" + strShipCity + " " + strShipState + " " + strShipZip + "\nUS";
            //    while (objCustomer.addressbookList.addressbook[1].addressee.Length > 100)
            //    {
            //        objCustomer.addressbookList.addressbook[1].addressee = objCustomer.addressbookList.addressbook[1].addressee.Substring(1);
            //    }
            //}

            /*if (objCustomer.addressbookList.addressbook[1].addrText.Length > 100)
            {
                //objCustomer.addressbookList.addressbook[1].addrText = strShipTo + "\n" + strShipAddress1 + "\n" + strShipAddress2 + "\n" + strShipCity + " " + strShipState + " " + strShipZip + "\nUS";
                while (objCustomer.addressbookList.addressbook[1].addrText.Length > 100)
                {
                    //objCustomer.addressbookList.addressbook[1].addrText = objCustomer.addressbookList.addressbook[1].addrText.Substring(1);
                }
            }*/

            WriteResponse objResponse = _service.add(objCustomer);
            if (objResponse.status.isSuccess)
            {
                objRecordRef = (RecordRef)objResponse.baseRef;
            }
            else
            {
                strError = objResponse.status.statusDetail[0].message;
            }
            return objRecordRef;
        }

        public static Customer AddCustomer1800(string strExternalID, string strParent, string strSalesRep, string strFirstName, string strLastName, string strEmail,
            string strBillTo, string strBillPhone, string strBillAddress1, string strBillAddress2, string strBillAddress3, string strBillCity, string strBillState, string strBillZip,
            string strShipTo, string strShipPhone, string strShipAddress1, string strShipAddress2, string strShipAddress3, string strShipCity, string strShipState, string strShipZip,
            out string strError)
        {
            strError = "";
            Customer objCustomer = new Customer();
            objCustomer.parent = new RecordRef();
            objCustomer.parent.internalId = strParent;
            objCustomer.parent.type = RecordType.customer;
            objCustomer.parent.typeSpecified = true;
            objCustomer.firstName = strFirstName;
            objCustomer.lastName = strLastName;
            objCustomer.email = strEmail;
            objCustomer.phone = strShipPhone;
            objCustomer.homePhone = strBillPhone;
            objCustomer.externalId = strExternalID;
            //objCustomer.altPhone = strShipPhone;

            objCustomer.salesRep = new RecordRef();
            objCustomer.salesRep.internalId = strSalesRep;
            objCustomer.salesRep.type = RecordType.employee;
            objCustomer.salesRep.typeSpecified = true;

            objCustomer.customForm = new RecordRef();
            objCustomer.customForm.internalId = GlobalSettings.Default.nsCustomerForm;
            objCustomer.customForm.type = RecordType.account;
            objCustomer.customForm.typeSpecified = true;

            //objCustomer.terms = new RecordRef();
            //objCustomer.terms.internalId = GlobalSettings.Default.nsCustomerTerms;
            //objCustomer.terms.type = RecordType.account;
            //objCustomer.terms.typeSpecified = true;

            objCustomer.addressbookList = new CustomerAddressbookList();
            objCustomer.addressbookList.addressbook = new CustomerAddressbook[2];
            objCustomer.addressbookList.addressbook[0] = new CustomerAddressbook();
            objCustomer.addressbookList.addressbook[0].defaultBilling = true;
            objCustomer.addressbookList.addressbook[0].defaultBillingSpecified = true;
            objCustomer.addressbookList.addressbook[0].defaultShipping = false;
            objCustomer.addressbookList.addressbook[0].defaultShippingSpecified = true;
            objCustomer.addressbookList.addressbook[0].isResidential = false;
            objCustomer.addressbookList.addressbook[0].isResidentialSpecified = true;
            objCustomer.addressbookList.addressbook[0].addr1 = strBillAddress1;
            objCustomer.addressbookList.addressbook[0].addr2 = strBillAddress2;
            objCustomer.addressbookList.addressbook[0].addr3 = strBillAddress3;
            objCustomer.addressbookList.addressbook[0].city = strBillCity;
            objCustomer.addressbookList.addressbook[0].zip = strBillZip;
            objCustomer.addressbookList.addressbook[0].state = strBillState;
            objCustomer.addressbookList.addressbook[0].attention = strBillTo;
            objCustomer.addressbookList.addressbook[0].label = strBillTo;
            objCustomer.addressbookList.addressbook[0].phone = strBillPhone;
            objCustomer.addressbookList.addressbook[0].country = Country._unitedStates;
            //objCustomer.addressbookList.addressbook[0].addressee = strBillTo + "\n" + strBillAddress1 + "\n" + strBillAddress2 + "\n" + strBillAddress3 + "\n" + strBillCity + " " + strBillState + " " + strBillZip + "\nUS\n" + strBillPhone;
            //objCustomer.addressbookList.addressbook[0].addrText = strBillTo + "\n" + strBillAddress1 + "\n" + strBillAddress2 + "\n" + strBillAddress3 + "\n" + strBillCity + " " + strBillState + " " + strBillZip + "\nUS\n" + strBillPhone;

            //if (objCustomer.addressbookList.addressbook[0].addressee.Length > 100)
            //{
            //    objCustomer.addressbookList.addressbook[0].addressee = strBillTo + "\n" + strBillAddress1 + "\n" + strBillAddress2 + "\n" + strBillCity + " " + strBillState + " " + strBillZip + "\nUS";
            //    while (objCustomer.addressbookList.addressbook[0].addressee.Length > 100)
            //    {
            //        objCustomer.addressbookList.addressbook[0].addressee = objCustomer.addressbookList.addressbook[0].addressee.Substring(1);
            //    }
            //}

            /*if (objCustomer.addressbookList.addressbook[0].addrText.Length > 100)
            {
                objCustomer.addressbookList.addressbook[0].addrText = strBillTo + "\n" + strBillAddress1 + "\n" + strBillAddress2 + "\n" + strBillCity + " " + strBillState + " " + strBillZip + "\nUS";
                while (objCustomer.addressbookList.addressbook[0].addrText.Length > 100)
                {
                    objCustomer.addressbookList.addressbook[0].addrText = objCustomer.addressbookList.addressbook[0].addrText.Substring(1);
                }
            }*/

            objCustomer.addressbookList.addressbook[1] = new CustomerAddressbook();
            objCustomer.addressbookList.addressbook[1].defaultBilling = false;
            objCustomer.addressbookList.addressbook[1].defaultBillingSpecified = true;
            objCustomer.addressbookList.addressbook[1].defaultShipping = true;
            objCustomer.addressbookList.addressbook[1].defaultShippingSpecified = true;
            objCustomer.addressbookList.addressbook[1].isResidential = false;
            objCustomer.addressbookList.addressbook[1].isResidentialSpecified = true;
            objCustomer.addressbookList.addressbook[1].addr1 = strShipAddress1;
            objCustomer.addressbookList.addressbook[1].addr2 = strShipAddress2;
            objCustomer.addressbookList.addressbook[1].addr3 = strShipAddress3;
            objCustomer.addressbookList.addressbook[1].city = strShipCity;
            objCustomer.addressbookList.addressbook[1].state = strShipState;
            objCustomer.addressbookList.addressbook[1].zip = strShipZip;
            objCustomer.addressbookList.addressbook[1].attention = strShipTo;
            objCustomer.addressbookList.addressbook[1].label = strShipTo;
            objCustomer.addressbookList.addressbook[1].phone = strShipPhone;
            objCustomer.addressbookList.addressbook[1].country = Country._unitedStates;
            //objCustomer.addressbookList.addressbook[1].addressee = strShipTo + "\n" + strShipAddress1 + "\n" + strShipAddress2 + "\n" + strShipAddress3 + "\n" + strShipCity + " " + strShipState + " " + strShipZip + "\nUS\n" + strShipPhone;
            //objCustomer.addressbookList.addressbook[1].addrText = strShipTo + "\n" + strShipAddress1 + "\n" + strShipAddress2 + "\n" + strShipAddress3 + "\n" + strShipCity + " " + strShipState + " " + strShipZip + "\nUS\n" + strShipPhone;

            //if (objCustomer.addressbookList.addressbook[1].addressee.Length > 100)
            //{
            //    objCustomer.addressbookList.addressbook[1].addressee = strShipTo + "\n" + strShipAddress1 + "\n" + strShipAddress2 + "\n" + strShipCity + " " + strShipState + " " + strShipZip + "\nUS";
            //    while (objCustomer.addressbookList.addressbook[1].addressee.Length > 100)
            //    {
            //        objCustomer.addressbookList.addressbook[1].addressee = objCustomer.addressbookList.addressbook[1].addressee.Substring(1);
            //    }
            //}

            /*if (objCustomer.addressbookList.addressbook[1].addrText.Length > 100)
            {
                objCustomer.addressbookList.addressbook[1].addrText = strShipTo + "\n" + strShipAddress1 + "\n" + strShipAddress2 + "\n" + strShipCity + " " + strShipState + " " + strShipZip + "\nUS";
                while (objCustomer.addressbookList.addressbook[1].addrText.Length > 100)
                {
                    objCustomer.addressbookList.addressbook[1].addrText = objCustomer.addressbookList.addressbook[1].addrText.Substring(1);
                }
            }*/

            return objCustomer;
        }

        public static RecordRef UpdateCustomer(string strInternalId, string strSalesRep, string strParent, string strFirstName, string strLastName, string strEmail,
            string strBillTo, string strBillPhone, string strBillAddress1, string strBillAddress2, string strBillAddress3, string strBillCity, string strBillState, string strBillZip,
            string strShipTo, string strShipPhone, string strShipAddress1, string strShipAddress2, string strShipAddress3, string strShipCity, string strShipState, string strShipZip,
            out string strError)
        {
            strError = "";
            RecordRef objRecordRef = null;
            RecordRef objTemp = new RecordRef();
            objTemp.internalId = strInternalId;
            objTemp.type = RecordType.customer;
            objTemp.typeSpecified = true;

            ReadResponse objReadResponse = _service.get(objTemp);

            string strInternalId1 = "";
            string strInternalId2 = "";

            if (objReadResponse.status.isSuccess)
            {
                Customer objCustomerTemp = (Customer)objReadResponse.record;
                CustomerAddressbookList temp = objCustomerTemp.addressbookList;

                if (objCustomerTemp.addressbookList.addressbook != null && objCustomerTemp.addressbookList.addressbook.Length > 1)
                {
                    foreach (CustomerAddressbook objAddress in objCustomerTemp.addressbookList.addressbook)
                    {
                        strInternalId1 = objCustomerTemp.addressbookList.addressbook[0].internalId;
                        strInternalId2 = objCustomerTemp.addressbookList.addressbook[1].internalId;
                    }
                }
            }
            Customer objCustomer = new Customer();
            objCustomer.internalId = strInternalId;
            objCustomer.parent = new RecordRef();
            //objCustomer.parent.internalId = strParent;
            //objCustomer.parent.type = RecordType.customer;
            //objCustomer.parent.typeSpecified = true;
            objCustomer.taxable = false;
            objCustomer.taxableSpecified = true;

            objCustomer.salesRep = new RecordRef();
            objCustomer.salesRep.internalId = strSalesRep;
            objCustomer.salesRep.type = RecordType.employee;
            objCustomer.salesRep.typeSpecified = true;

            //objCustomer.customForm = new RecordRef();
            //objCustomer.customForm.internalId = "12";
            //objCustomer.customForm.type = RecordType.account;
            //objCustomer.customForm.typeSpecified = true;

            //objCustomer.terms = new RecordRef();
            //objCustomer.terms.internalId = "2";
            //objCustomer.terms.type = RecordType.account;
            //objCustomer.terms.typeSpecified = true;

            objCustomer.taxable = false;
            objCustomer.taxableSpecified = true;

            objCustomer.salesRep = new RecordRef();
            objCustomer.salesRep.internalId = strSalesRep;
            objCustomer.salesRep.type = RecordType.employee;
            objCustomer.salesRep.typeSpecified = true;

            objCustomer.defaultAddress = "";
            objCustomer.addressbookList = new CustomerAddressbookList();
            objCustomer.addressbookList.replaceAll = true;
            objCustomer.addressbookList.addressbook = new CustomerAddressbook[2];
            objCustomer.addressbookList.addressbook[0] = new CustomerAddressbook();
            objCustomer.addressbookList.addressbook[0].internalId = strInternalId1;
            objCustomer.addressbookList.addressbook[0].@override = true;
            objCustomer.addressbookList.addressbook[0].overrideSpecified = true;
            objCustomer.addressbookList.addressbook[0].defaultBilling = true;
            objCustomer.addressbookList.addressbook[0].defaultBillingSpecified = true;
            objCustomer.addressbookList.addressbook[0].defaultShipping = false;
            objCustomer.addressbookList.addressbook[0].defaultShippingSpecified = true;
            objCustomer.addressbookList.addressbook[0].isResidential = false;
            objCustomer.addressbookList.addressbook[0].isResidentialSpecified = true;
            objCustomer.addressbookList.addressbook[0].addr1 = strBillAddress1;
            objCustomer.addressbookList.addressbook[0].addr2 = strBillAddress2;
            objCustomer.addressbookList.addressbook[0].addr3 = strBillAddress3;
            objCustomer.addressbookList.addressbook[0].city = strBillCity;
            objCustomer.addressbookList.addressbook[0].zip = strBillZip;
            objCustomer.addressbookList.addressbook[0].state = strBillState;
            objCustomer.addressbookList.addressbook[0].phone = strBillPhone;
            objCustomer.addressbookList.addressbook[0].attention = strBillTo;
            objCustomer.addressbookList.addressbook[0].label = strBillTo;
            objCustomer.addressbookList.addressbook[0].country = Country._unitedStates;
            objCustomer.addressbookList.addressbook[0].addressee = strBillTo + "\n" + strBillAddress1 + "\n" + strBillAddress2 + "\n" + strBillAddress3 + "\n" + strBillCity + " " + strBillState + " " + strBillZip + "\nUS\n" + strBillPhone;
            objCustomer.addressbookList.addressbook[0].addrText = strBillTo + "\n" + strBillAddress1 + "\n" + strBillAddress2 + "\n" + strBillAddress3 + "\n" + strBillCity + " " + strBillState + " " + strBillZip + "\nUS\n" + strBillPhone;

            if (objCustomer.addressbookList.addressbook[0].addressee.Length > 100)
            {
                objCustomer.addressbookList.addressbook[0].addressee = strBillTo + "\n" + strBillAddress1 + "\n" + strBillAddress2 + "\n" + strBillCity + " " + strBillState + " " + strBillZip + "\nUS";
                while (objCustomer.addressbookList.addressbook[0].addressee.Length > 100)
                {
                    objCustomer.addressbookList.addressbook[0].addressee = objCustomer.addressbookList.addressbook[0].addressee.Substring(1);
                }
            }

            if (objCustomer.addressbookList.addressbook[0].addrText.Length > 100)
            {
                objCustomer.addressbookList.addressbook[0].addrText = strBillTo + "\n" + strBillAddress1 + "\n" + strBillAddress2 + "\n" + strBillCity + " " + strBillState + " " + strBillZip + "\nUS";
                while (objCustomer.addressbookList.addressbook[0].addrText.Length > 100)
                {
                    objCustomer.addressbookList.addressbook[0].addrText = objCustomer.addressbookList.addressbook[0].addrText.Substring(1);
                }
            }

            objCustomer.addressbookList.addressbook[1] = new CustomerAddressbook();
            objCustomer.addressbookList.addressbook[1].internalId = strInternalId2;
            objCustomer.addressbookList.addressbook[1].@override = true;
            objCustomer.addressbookList.addressbook[1].overrideSpecified = true;
            objCustomer.addressbookList.addressbook[1].defaultBilling = false;
            objCustomer.addressbookList.addressbook[1].defaultBillingSpecified = true;
            objCustomer.addressbookList.addressbook[1].defaultShipping = true;
            objCustomer.addressbookList.addressbook[1].defaultShippingSpecified = true;
            objCustomer.addressbookList.addressbook[1].isResidential = false;
            objCustomer.addressbookList.addressbook[1].isResidentialSpecified = true;
            objCustomer.addressbookList.addressbook[1].addr1 = strShipAddress1;
            objCustomer.addressbookList.addressbook[1].addr2 = strShipAddress2;
            objCustomer.addressbookList.addressbook[1].addr3 = strShipAddress3;
            objCustomer.addressbookList.addressbook[1].city = strShipCity;
            objCustomer.addressbookList.addressbook[1].state = strShipState;
            objCustomer.addressbookList.addressbook[1].zip = strShipZip;
            objCustomer.addressbookList.addressbook[1].attention = strShipTo;
            objCustomer.addressbookList.addressbook[1].label = strShipTo;
            objCustomer.addressbookList.addressbook[1].phone = strShipPhone;
            objCustomer.addressbookList.addressbook[1].country = Country._unitedStates;
            objCustomer.addressbookList.addressbook[1].addressee = strShipTo + "\n" + strShipAddress1 + "\n" + strShipAddress2 + "\n" + strShipAddress3 + "\n" + strShipCity + " " + strShipState + " " + strShipZip + "\nUS \n" + strShipPhone;
            objCustomer.addressbookList.addressbook[1].addrText = strShipTo + "\n" + strShipAddress1 + "\n" + strShipAddress2 + "\n" + strShipAddress3 + "\n" + strShipCity + " " + strShipState + " " + strShipZip + "\nUS \n" + strShipPhone;

            if (objCustomer.addressbookList.addressbook[1].addressee.Length > 100)
            {
                objCustomer.addressbookList.addressbook[1].addressee = strShipTo + "\n" + strShipAddress1 + "\n" + strShipAddress2 + "\n" + strShipCity + " " + strShipState + " " + strShipZip + "\nUS";
                while (objCustomer.addressbookList.addressbook[1].addressee.Length > 100)
                {
                    objCustomer.addressbookList.addressbook[1].addressee = objCustomer.addressbookList.addressbook[1].addressee.Substring(1);
                }
            }

            if (objCustomer.addressbookList.addressbook[1].addrText.Length > 100)
            {
                objCustomer.addressbookList.addressbook[1].addrText = strShipTo + "\n" + strShipAddress1 + "\n" + strShipAddress2 + "\n" + strShipCity + " " + strShipState + " " + strShipZip + "\nUS";
                while (objCustomer.addressbookList.addressbook[1].addrText.Length > 100)
                {
                    objCustomer.addressbookList.addressbook[1].addrText = objCustomer.addressbookList.addressbook[1].addrText.Substring(1);
                }
            }

            WriteResponse objResponse = _service.update(objCustomer);
            if (objResponse.status.isSuccess)
            {
                objRecordRef = (RecordRef)objResponse.baseRef;
            }
            else
            {
                strError = objResponse.status.statusDetail[0].message;
            }
            return objRecordRef;
        }

        private static System.String getStatusDetails(Status status)
        {
            System.Text.StringBuilder sb = new System.Text.StringBuilder();
            for (int i = 0; i < status.statusDetail.Length; i++)
            {
                sb.Append("[Code=" + status.statusDetail[i].code + "] " + status.statusDetail[i].message + "\n");
            }
            return sb.ToString();
        }

    }
}
