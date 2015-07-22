using System;
using System.Globalization;
using System.Threading;
using System.Data;
using System.Data.OleDb;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using NetsuiteOnlineServicesOrders.wsNetsuite;

namespace NetsuiteOnlineServicesOrders
{
    public class Common
    {
        public static bool RUN;
        public static DataSet dsData = new DataSet();
        public static DataTable dsTable = new DataTable();
        public static ArrayList arrShipType = new ArrayList();
        public static ArrayList arrCourierType = new ArrayList();
        public static ArrayList arrReseller = new ArrayList();
        public static ArrayList arrAcum = new ArrayList();
        public static Form1 objForm;

        public static string strMessage = "";
        public static ProcessResult LastProcessResult;

        public enum ProcessResult
        {
            Ok1,
            Ok2,
            Added,
            Error,
            Processed,
            Closed
        }

        public static void LoadGeneralSettings()
        {
            arrShipType.Clear();
            string[] arrShip = GlobalSettings.Default.nsShipType.Split(new string[] { "," }, StringSplitOptions.None);
            foreach (string strType   in arrShip)
            {
                arrShipType.Add(strType);
            }

            arrCourierType.Clear();
            string[] arrCourier = GlobalSettings.Default.nsCourierType.Split(new string[] { "," }, StringSplitOptions.None);
            foreach (string strType in arrCourier)
            {
                arrCourierType.Add(strType);
            }

            arrAcum.Clear();
            string[] arrTemp = GlobalSettings.Default.customersAcumByShippingDate.Split(new string[] { "," }, StringSplitOptions.None);
            foreach (string strTemp in arrTemp)
            {
                arrAcum.Add(strTemp);
            }

            arrReseller.Clear();
            string[] arrRes = GlobalSettings.Default.nsReseller.Split(new string[] { "," }, StringSplitOptions.None);
            foreach (string strTemp in arrRes)
            {
                arrReseller.Add(strTemp);
            }
        }

        public static void LoadExcelData(string filelocation)
        {
            string[] arrCustomers = GlobalSettings.Default.exSheet.Split(new string[] { "," }, StringSplitOptions.None);
            foreach (string strCustom in arrCustomers)
            {
                LoadExcelData(filelocation, strCustom);
            }
            LoadGeneralSettings();
        }
               
        public static void LoadCSVData(string filelocation)
        {
            System.IO.FileInfo objFile = new System.IO.FileInfo(filelocation);

            LoadTextData(objFile.DirectoryName, objFile.Name);
            LoadGeneralSettings();
        }

        public static void LoadTable(DataTable dtData)
        {
            dsTable = dtData;
            ThreadStart objStart = new ThreadStart(LoadTable);
            Thread objThread = new Thread(objStart);
            objThread.CurrentCulture = Thread.CurrentThread.CurrentCulture;
            objThread.CurrentUICulture = Thread.CurrentThread.CurrentUICulture;
            objThread.Name = "loadTable";
            objThread.Start();
            RUN = true;
        }

        public static void LoadTableMassMarket(DataTable dtData)
        {
            dsTable = dtData;
            ThreadStart objStart = new ThreadStart(LoadTableMassMarket);
            Thread objThread = new Thread(objStart);
            objThread.CurrentCulture = Thread.CurrentThread.CurrentCulture;
            objThread.CurrentUICulture = Thread.CurrentThread.CurrentUICulture;
            objThread.Name = "loadTable";
            objThread.Start();
            RUN = true;
        }

        public static void LoadTable1800(DataTable dtData)
        {
            dsTable = dtData;
            ThreadStart objStart = new ThreadStart(LoadTable1800);
            Thread objThread = new Thread(objStart);
            objThread.CurrentCulture = Thread.CurrentThread.CurrentCulture;
            objThread.CurrentUICulture = Thread.CurrentThread.CurrentUICulture;
            objThread.Start(); 
            RUN = true;
        }

        private static void LoadTable()
        {
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Hashtable htCustomers = new Hashtable();
            Hashtable htProducts = new Hashtable();
            Hashtable htOrders = new Hashtable();
            Util.Login(GlobalSettings.Default.nsUser, GlobalSettings.Default.nsPassword, GlobalSettings.Default.nsAccount, GlobalSettings.Default.nsRole);

            int numUserIndex = 0;
            ArrayList arrUsers = Util.GetUsers(GlobalSettings.Default.nsSalesRoleId);
            if (arrUsers.Count == 0)
            {
                string[] arrTemp = GlobalSettings.Default.nsSalesRep.Split(new string[] { "," }, StringSplitOptions.None);
                foreach (string strTemp in arrTemp)
                {
                    Employee objEmp = new Employee();
                    objEmp.internalId = strTemp;
                    arrUsers.Add(objEmp);
                }
                //ShowError("SalesReps haven't been found. Role ID: " + GlobalSettings.Default.nsSalesRoleId);
                //return;
            }

            int numIndex = 0;
            ArrayList arrItems = new ArrayList();
            string strRef = "PO";
            bool booIsByPO = true;
            string strDateFormat = "yyyy-MM-dd";
            if (arrAcum.Contains(dsTable.TableName))
            {
                strDateFormat = "GLyyyy-MM-dd";
                strRef = "Ship Date";
                booIsByPO = false;
            }

            Customer objReseller = new Customer();
            int numLineId = 1;
            int numTotal = 0;
            SalesOrder objSalesOrder = new SalesOrder();
            foreach (DataRow objRow in dsTable.Rows)
            {
                string strEmployeeId = ((Employee)arrUsers[numUserIndex]).internalId;
                string strPoNumber = "";
                string strPoAux = "";
                bool booAdd = false;
                try
                {
                    objForm.IncrementProcess();
                    strPoNumber = objRow[strRef].ToString();
                    strPoAux = objRow["PO"].ToString();
                    if (objRow[strRef] is DateTime)
                    {
                        strPoNumber = ((DateTime)objRow[strRef]).ToString(strDateFormat);
                    }
                    if (!htOrders.ContainsKey(strPoNumber))
                    {
                        string strCustomer = objRow["Customer"].ToString();
                        string strId = Util.GetOrderId(strPoNumber, (arrReseller.IndexOf(strCustomer) + 1).ToString());
                        if (string.IsNullOrEmpty(strId))
                        {
                            if (!htCustomers.ContainsKey(strCustomer))
                            {
                                objReseller = Util.GetCustomer(strCustomer);
                                if (objReseller == null || objReseller.internalId == null)
                                {
                                    throw new Exception("Customer hasn't been found. Reseller ID: " + strCustomer);
                                }
                                else
                                {
                                    htCustomers.Add(strCustomer, objReseller.internalId);
                                }
                            }

                            string strBillTo = "";
                            string strBillPhone = "";
                            string strBillAddress1 = "";
                            string strBillAddress2 = "";
                            string strBillAddress3 = "";
                            string strBillCity = "";
                            string strBillState = "";
                            string strBillZip = "";
                            foreach (CustomerAddressbook objAddress in objReseller.addressbookList.addressbook)
                            {
                                if (objAddress.defaultBilling)
                                {
                                    strBillTo = "ttt";// objAddress.attention;
                                    strBillPhone = objAddress.phone;
                                    strBillAddress1 = objAddress.addr1;
                                    strBillAddress2 = objAddress.addr2;
                                    strBillAddress3 = objAddress.addr3;
                                    strBillCity = objAddress.city;
                                    strBillState = objAddress.state;
                                    strBillZip = objAddress.zip;
                                }
                            }

                            string strCustomerId = htCustomers[strCustomer].ToString();

                            string strProduct = objRow["SKU"].ToString();
                            if (string.IsNullOrEmpty(strProduct))
                            {
                                throw new Exception("Product is empty.");
                            }
                            string strProductId = "";
                            if (!htProducts.ContainsKey(strProduct))
                            {
                                ArrayList arrProducts = Util.GetItemId(strProduct);
                                if (arrProducts.Count == 0)
                                {

                                    throw new Exception("Product hasn't been found. Product: " + strProduct);
                                }
                                else if (arrProducts.Count == 1)
                                {
                                    strProductId = ((string[])arrProducts[0])[0];
                                    htProducts.Add(strProduct, strProductId);
                                }
                                else
                                {
                                    Form2 objForm2 = new Form2(arrProducts);
                                    System.Windows.Forms.DialogResult objResult = objForm2.ShowDialog(objForm);
                                    while (objResult != System.Windows.Forms.DialogResult.OK)
                                    {
                                        objForm2.ShowDialog(objForm);
                                    }
                                    if (objResult == System.Windows.Forms.DialogResult.OK)
                                    {
                                        strProductId = objForm2.ItemId;
                                    }
                                }
                            }
                            else
                            {
                                strProductId = htProducts[strProduct].ToString();
                            }

                            string strShipToName = objRow["ShipToName"].ToString();
                            string strShipPhone = objRow["ShipToPhone"].ToString();

                            string strShipAddress1 = objRow["ShipToAddress1"].ToString();
                            string strShipAddress2 = objRow["ShipToAddress2"].ToString();
                            string strShipCity = objRow["ShipToCity"].ToString();
                            string strShipZipCode = objRow["ShipToZipCode"].ToString();
                            string strShipState = objRow["ShipToState"].ToString();
                            string strShipEmail = objRow["ShipToEmail"].ToString();
                            string strShipType = objRow["Transport1"].ToString().Trim();
                            string strCourierType = objRow["Transport2"].ToString().Trim();

                            RecordRef objSubCustomer = new RecordRef();
                            objSubCustomer.internalId = objReseller.internalId;
                            objSubCustomer.type = RecordType.customer;
                            objSubCustomer.typeSpecified = true;

                            if (!string.IsNullOrEmpty(strShipToName))
                            {
                                string strFirstName = "";
                                string strLastName = "";

                                string[] arrNames = strShipToName.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                                if (arrNames.Length == 1 || arrNames.Length != 0)
                                {
                                    arrNames = arrNames[0].Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                                    if (arrNames.Length > 1)
                                    {
                                        strFirstName = arrNames[0];
                                        strLastName = arrNames[1];
                                    }
                                    else
                                    {
                                        strFirstName = strShipToName;
                                    }
                                }
                                else
                                {
                                    strFirstName = strShipToName;
                                }
                                try
                                {
                                    string strError = "";
                                    //objSubCustomer = Util.GetCustomer(strCustomerId, strShipPhone);
                                    objSubCustomer = null;

                                    if (string.IsNullOrEmpty(strBillTo))
                                    {
                                        throw new Exception("Billing information missed!");
                                    }

                                    if (objSubCustomer == null || objSubCustomer.internalId == null)
                                    {
                                        objSubCustomer = Util.AddCustomer(strCustomerId, strEmployeeId, strFirstName, strLastName, strShipEmail,
                                            strBillTo, strBillPhone, strBillAddress1, strBillAddress2, strBillAddress3, strBillCity, strBillState, strBillZip,
                                            strShipToName, strShipPhone, strShipAddress1, strShipAddress2, "", strShipCity, strShipState, strShipZipCode,
                                            out strError);
                                    }
                                    else
                                    {
                                        objSubCustomer = Util.UpdateCustomer(objSubCustomer.internalId, strEmployeeId, strCustomerId, strFirstName, strLastName, strShipEmail,
                                            strBillTo, strBillPhone, strBillAddress1, strBillAddress2, strBillAddress3, strBillCity, strBillState, strBillZip,
                                            strShipToName, strShipPhone, strShipAddress1, strShipAddress2, "", strShipCity, strShipState, strShipZipCode,
                                            out strError);
                                    }

                                    if (!string.IsNullOrEmpty(strError))
                                    {
                                        throw new Exception(strError);
                                    }
                                    strCustomerId = objSubCustomer.internalId;
                                }
                                catch (Exception objExc)
                                {
                                    throw new Exception("Subcustomer hasn't created or updated (" + strShipToName + "): " + objExc.Message);
                                }
                            }

                            DateTime strShipDate = DateTime.Parse(objRow["Ship Date"].ToString());
                            DateTime strDeliDate = strShipDate;
                            try
                            {
                                strDeliDate = DateTime.Parse(objRow["DeliveryDate"].ToString());
                            }
                            catch (Exception)
                            {
                            }
                            DateTime strInvoiceDate = strDeliDate;
                            try
                            {
                                strInvoiceDate = DateTime.Parse(objRow["InvoiceDate"].ToString());
                            }
                            catch (Exception)
                            {
                            }
                            DateTime objOrderDate = DateTime.Now;
                            try
                            {
                                objOrderDate = DateTime.Parse(objRow["F1"].ToString());
                            }
                            catch (Exception)
                            {
                            }

                            double numQty = double.Parse(objRow["Q"].ToString());

                            if (!string.IsNullOrEmpty(strCustomer) && !string.IsNullOrEmpty(strProduct) && !string.IsNullOrEmpty(strPoNumber))
                            {
                                if (booIsByPO)
                                {
                                    for (int i = 0; i < numQty; i++)
                                    {
                                        arrItems.Add(new object[] { strProductId, 1, strShipDate, strDeliDate, arrShipType.IndexOf(strShipType) + 1, objRow["PO"].ToString(), objRow["ORDERMESSAGETEXT"].ToString(), arrCourierType.IndexOf(strCourierType) + 1 });
                                    }
                                }
                                else
                                {
                                    arrItems.Add(new object[] { strProductId, numQty, strShipDate, strDeliDate, arrShipType.IndexOf(strShipType) + 1, objRow["PO"].ToString(), objRow["ORDERMESSAGETEXT"].ToString(), arrCourierType.IndexOf(strCourierType) + 1 });
                                }
                                string strNextRef = "";
                                if (numIndex + 1 < dsTable.Rows.Count)
                                {
                                    strNextRef = dsTable.Rows[numIndex + 1][strRef].ToString();
                                    if (dsTable.Rows[numIndex + 1][strRef] is DateTime)
                                    {
                                        strNextRef = ((DateTime)dsTable.Rows[numIndex + 1][strRef]).ToString(strDateFormat);
                                    }
                                }
                                if (numIndex + 1 >= dsTable.Rows.Count || !strNextRef.Equals(strPoNumber))
                                {
                                    // Insert order
                                    string strError = "";
                                    string strClase = "";
                                    strId = Util.AddSalesOrder(arrReseller.IndexOf(strCustomer) + 1, objSubCustomer, strPoNumber, strPoNumber, objOrderDate, DateTime.Now, objOrderDate, arrItems, strEmployeeId, strInvoiceDate, strClase, out strError);
                                    if (!string.IsNullOrEmpty(strId))
                                    {
                                        LastProcessResult = ProcessResult.Ok2;
                                        strMessage = strId;
                                        numUserIndex++;
                                        numUserIndex = (numUserIndex) % arrUsers.Count;
                                        if (!htOrders.ContainsKey(strPoNumber))
                                        {
                                            htOrders.Add(strPoNumber, strId);
                                        }
                                    }
                                    else
                                    {
                                        LastProcessResult = ProcessResult.Error;
                                        strMessage = strError;
                                    }
                                    arrItems.Clear();
                                }
                                else
                                {
                                    LastProcessResult = ProcessResult.Ok1;
                                    strMessage = "Linea agregada: " + strProduct;
                                }
                            }
                            else
                            {
                                LastProcessResult = ProcessResult.Error;
                                strMessage = "Información incompleta en SKU";
                            }
                        }
                        else
                        {
                            //objSalesOrder = (SalesOrder)Util.GetEntity(strId, RecordType.salesOrder);
                            //if (!htOrders.ContainsKey(strPoNumber))
                            //{
                            //    htOrders.Add(strPoNumber, strId);
                            //}
                            //strMessage = "Registro ya procesado: " + strId;
                            //LastProcessResult = ProcessResult.Processed;
                            objSalesOrder = (SalesOrder)Util.GetEntity(strId, RecordType.salesOrder);
                            numTotal = objSalesOrder.itemList.item.Length;
                            if (GlobalSettings.Default.AddLines.Equals("Yes"))
                            {
                                numTotal = 0;
                            }
                            if (!htOrders.ContainsKey(strPoNumber))
                            {
                                htOrders.Add(strPoNumber, strId);
                            }
                            booAdd = true;
                        }
                    }
                    else
                    {
                        booAdd = true;
                    }

                    if(booAdd)
                    {
                        if (htOrders[strPoNumber].Equals("error"))
                        {
                            LastProcessResult = ProcessResult.Error;
                            strMessage = "Linea omitida por error en linea anteriores a la orden";
                        }
                        else
                        {
                            // Depende del archivo de configuración si es True AddLines numTotal es 0
                            if (numTotal < 1)
                            {
                                //Adicional
                                string strProduct = objRow["SKU"].ToString();
                                if (string.IsNullOrEmpty(strProduct))
                                {
                                    throw new Exception("Product is empty.");
                                }
                                string strProductId = "";
                                if (!htProducts.ContainsKey(strProduct))
                                {
                                    ArrayList arrProducts = Util.GetItemId(strProduct);
                                    if (arrProducts.Count == 0)
                                    {
                                        throw new Exception("Product hasn't been found. Product: " + strProduct);
                                    }
                                    else if (arrProducts.Count == 1)
                                    {
                                        strProductId = ((string[])arrProducts[0])[0];
                                        htProducts.Add(strProduct, strProductId);
                                    }
                                    else
                                    {
                                        Form2 objForm2 = new Form2(arrProducts);
                                        System.Windows.Forms.DialogResult objResult = objForm2.ShowDialog(objForm);
                                        while (objResult != System.Windows.Forms.DialogResult.OK)
                                        {
                                            objForm2.ShowDialog(objForm);
                                        }
                                        if (objResult == System.Windows.Forms.DialogResult.OK)
                                        {
                                            strProductId = objForm2.ItemId;
                                        }
                                    }
                                }
                                else
                                {
                                    strProductId = htProducts[strProduct].ToString();
                                }

                                DateTime strShipDate = DateTime.Parse(objRow["Ship Date"].ToString());
                                DateTime strDeliDate = strShipDate;
                                try
                                {
                                    strDeliDate = DateTime.Parse(objRow["DeliveryDate"].ToString());
                                }
                                catch (Exception)
                                {
                                }
                                DateTime objOrderDate = DateTime.Now;
                                try
                                {
                                    objOrderDate = DateTime.Parse(objRow["F1"].ToString());
                                }
                                catch (Exception)
                                {
                                }
                                double numQty = double.Parse(objRow["Q"].ToString());
                                string strShipType = objRow["Transport1"].ToString().Trim();
                                string strCourierType = objRow["Transport2"].ToString().Trim();

                                if (booIsByPO)
                                {
                                    for (int i = 0; i < numQty; i++)
                                    {
                                        arrItems.Add(new object[] { strProductId, 1, strShipDate, strDeliDate, arrShipType.IndexOf(strShipType) + 1, objRow["PO"].ToString(), arrCourierType.IndexOf(strCourierType) + 1  });
                                    }
                                }
                                else
                                {
                                    arrItems.Add(new object[] { strProductId, numQty, strShipDate, strDeliDate, arrShipType.IndexOf(strShipType) + 1, objRow["PO"].ToString(), arrCourierType.IndexOf(strCourierType) + 1 });
                                }

                                string strNextRef = "";
                                if (numIndex + 1 < dsTable.Rows.Count)
                                {
                                    strNextRef = dsTable.Rows[numIndex + 1][strRef].ToString();
                                    if (dsTable.Rows[numIndex + 1][strRef] is DateTime)
                                    {
                                        strNextRef = ((DateTime)dsTable.Rows[numIndex + 1][strRef]).ToString(strDateFormat);
                                    }
                                }
                                if (numIndex + 1 >= dsTable.Rows.Count || !strNextRef.Equals(strPoNumber))
                                {
                                    // Insert order
                                    string strError = "";
                                    string strId = objSalesOrder.internalId;
                                    Util.UpdateSalesOrder(objSalesOrder, arrItems, out strError);
                                    if (!string.IsNullOrEmpty(strId))
                                    {
                                        LastProcessResult = ProcessResult.Added;
                                        strMessage = "Linea adicionada para la orden: " + strPoNumber;
                                        numUserIndex++;
                                        numUserIndex = (numUserIndex) % arrUsers.Count;
                                        if (!htOrders.ContainsKey(strPoNumber))
                                        {
                                            htOrders.Add(strPoNumber, strId);
                                        }
                                    }
                                    else
                                    {
                                        LastProcessResult = ProcessResult.Error;
                                        strMessage = strError;
                                    }
                                    arrItems.Clear();
                                }
                                else
                                {
                                    LastProcessResult = ProcessResult.Ok1;
                                    strMessage = "Linea agregada en la orden: " + strPoNumber;
                                }
                            }
                            else
                            {
                                strMessage = "Registro ya procesado: " + htOrders[strPoNumber];
                                LastProcessResult = ProcessResult.Processed;
                            }
                        }
                    }
                }
                catch (Exception objExc)
                {
                    arrItems.Clear();
                    if (!htOrders.ContainsKey(strPoNumber))
                    {
                        htOrders.Add(strPoNumber, "error");
                    }
                    else
                    {
                        htOrders[strPoNumber] = "error";
                    }
                    LastProcessResult = ProcessResult.Error;
                    strMessage = objExc.Message;
                    //ShowError(objExc.Message);
                }

                string strRefAux = "";
                if (numIndex + 1 < dsTable.Rows.Count)
                {
                    strRefAux = dsTable.Rows[numIndex + 1][strRef].ToString();
                    if (dsTable.Rows[numIndex + 1][strRef] is DateTime)
                    {
                        strRefAux = ((DateTime)dsTable.Rows[numIndex + 1][strRef]).ToString(strDateFormat);
                    }
                }
                if (numIndex + 1 >= dsTable.Rows.Count || !strRefAux.Equals(strPoNumber))
                {
                    numLineId = 0;
                }
                numIndex++;
                numLineId++;
            }

            objForm.IncrementProcess();
            RUN = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
        }


        private static void LoadTableMassMarket()
        {
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Hashtable htCustomers = new Hashtable();
            Hashtable htProducts = new Hashtable();
            Hashtable htOrders = new Hashtable();
            Util.Login(GlobalSettings.Default.nsUser, GlobalSettings.Default.nsPassword, GlobalSettings.Default.nsAccount, GlobalSettings.Default.nsRole);

            int numUserIndex = 0;
            ArrayList arrUsers = Util.GetUsers(GlobalSettings.Default.nsSalesRoleId);
            if (arrUsers.Count == 0)
            {
                string[] arrTemp = GlobalSettings.Default.nsSalesRep.Split(new string[] { "," }, StringSplitOptions.None);
                foreach (string strTemp in arrTemp)
                {
                    Employee objEmp = new Employee();
                    objEmp.internalId = strTemp;
                    arrUsers.Add(objEmp);
                }
                //ShowError("SalesReps haven't been found. Role ID: " + GlobalSettings.Default.nsSalesRoleId);
                //return;
            }

            int numIndex = 0;
            ArrayList arrItems = new ArrayList();
            string strRef = "PO";
            bool booIsByPO = true;
            string strDateFormat = "yyyy-MM-dd";
            if (arrAcum.Contains(dsTable.TableName))
            {
                strDateFormat = "GLyyyy-MM-dd";
                strRef = "Ship Date";
                booIsByPO = false;
            }

            Customer objReseller = new Customer();
            int numLineId = 1;
            int numTotal = 0;
            SalesOrder objSalesOrder = new SalesOrder();
            foreach (DataRow objRow in dsTable.Rows)
            {
                string strEmployeeId = ((Employee)arrUsers[numUserIndex]).internalId;
                string strPoNumber = "";
                string strPoAux = "";
                bool booAdd = false;
                try
                {
                    objForm.IncrementProcess();
                    strPoNumber = objRow[strRef].ToString();
                    strPoAux = objRow["PO"].ToString();
                    if (objRow[strRef] is DateTime)
                    {
                        strPoNumber = ((DateTime)objRow[strRef]).ToString(strDateFormat);
                    }
                    if (!htOrders.ContainsKey(strPoNumber))
                    {
                        string strCustomer = objRow["Reseller"].ToString();
                        string strId = Util.GetOrderId(strPoNumber, (arrReseller.IndexOf(strCustomer) + 1).ToString());
                        if (string.IsNullOrEmpty(strId))
                        {
                            if (!htCustomers.ContainsKey(strCustomer))
                            {
                                objReseller = Util.GetCustomer(strCustomer);
                                if (objReseller == null || objReseller.internalId == null)
                                {
                                    throw new Exception("Customer hasn't been found. Reseller ID: " + strCustomer);
                                }
                                else
                                {
                                    htCustomers.Add(strCustomer, objReseller.internalId);
                                }
                            }

                            string strBillTo = "";
                            string strBillPhone = "";
                            string strBillAddress1 = "";
                            string strBillAddress2 = "";
                            string strBillAddress3 = "";
                            string strBillCity = "";
                            string strBillState = "";
                            string strBillZip = "";
                            foreach (CustomerAddressbook objAddress in objReseller.addressbookList.addressbook)
                            {
                                if (objAddress.defaultBilling)
                                {
                                    strBillTo = "ttt";// objAddress.attention;
                                    strBillPhone = objAddress.phone;
                                    strBillAddress1 = objAddress.addr1;
                                    strBillAddress2 = objAddress.addr2;
                                    strBillAddress3 = objAddress.addr3;
                                    strBillCity = objAddress.city;
                                    strBillState = objAddress.state;
                                    strBillZip = objAddress.zip;
                                }
                            }

                            string strCustomerId = htCustomers[strCustomer].ToString();

                            string strProduct = objRow["SKU"].ToString();
                            if (string.IsNullOrEmpty(strProduct))
                            {
                                throw new Exception("Product is empty.");
                            }
                            string strProductId = "";
                            if (!htProducts.ContainsKey(strProduct))
                            {
                                ArrayList arrProducts = Util.GetItemId(strProduct);
                                if (arrProducts.Count == 0)
                                {

                                    throw new Exception("Product hasn't been found. Product: " + strProduct);
                                }
                                else if (arrProducts.Count == 1)
                                {
                                    strProductId = ((string[])arrProducts[0])[0];
                                    htProducts.Add(strProduct, strProductId);
                                }
                                else
                                {
                                    Form2 objForm2 = new Form2(arrProducts);
                                    System.Windows.Forms.DialogResult objResult = objForm2.ShowDialog(objForm);
                                    while (objResult != System.Windows.Forms.DialogResult.OK)
                                    {
                                        objForm2.ShowDialog(objForm);
                                    }
                                    if (objResult == System.Windows.Forms.DialogResult.OK)
                                    {
                                        strProductId = objForm2.ItemId;
                                    }
                                }
                            }
                            else
                            {
                                strProductId = htProducts[strProduct].ToString();
                            }

                            /*string strShipToName = objRow["ShipToName"].ToString();
                            string strShipPhone = objRow["ShipToPhone"].ToString();

                            string strShipAddress1 = objRow["ShipToAddress1"].ToString();
                            string strShipAddress2 = objRow["ShipToAddress2"].ToString();
                            string strShipCity = objRow["ShipToCity"].ToString();
                            string strShipZipCode = objRow["ShipToZipCode"].ToString();
                            string strShipState = objRow["ShipToState"].ToString();
                            string strShipEmail = objRow["ShipToEmail"].ToString();*/
                            string strShipType = "";
                            string strCourierType = "";

                            RecordRef objSubCustomer = new RecordRef();
                            objSubCustomer.internalId = objReseller.internalId;
                            objSubCustomer.type = RecordType.customer;
                            objSubCustomer.typeSpecified = true;

                            /*if (!string.IsNullOrEmpty(strShipToName))
                            {
                                string strFirstName = "";
                                string strLastName = "";

                                string[] arrNames = strShipToName.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                                if (arrNames.Length == 1 || arrNames.Length != 0)
                                {
                                    arrNames = arrNames[0].Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                                    if (arrNames.Length > 1)
                                    {
                                        strFirstName = arrNames[0];
                                        strLastName = arrNames[1];
                                    }
                                    else
                                    {
                                        strFirstName = strShipToName;
                                    }
                                }
                                else
                                {
                                    strFirstName = strShipToName;
                                }
                                try
                                {
                                    string strError = "";
                                    //objSubCustomer = Util.GetCustomer(strCustomerId, strShipPhone);
                                    objSubCustomer = null;

                                    if (string.IsNullOrEmpty(strBillTo))
                                    {
                                        throw new Exception("Billing information missed!");
                                    }

                                    if (objSubCustomer == null || objSubCustomer.internalId == null)
                                    {
                                        objSubCustomer = Util.AddCustomer(strCustomerId, strEmployeeId, strFirstName, strLastName, strShipEmail,
                                            strBillTo, strBillPhone, strBillAddress1, strBillAddress2, strBillAddress3, strBillCity, strBillState, strBillZip,
                                            strShipToName, strShipPhone, strShipAddress1, strShipAddress2, "", strShipCity, strShipState, strShipZipCode,
                                            out strError);
                                    }
                                    else
                                    {
                                        objSubCustomer = Util.UpdateCustomer(objSubCustomer.internalId, strEmployeeId, strCustomerId, strFirstName, strLastName, strShipEmail,
                                            strBillTo, strBillPhone, strBillAddress1, strBillAddress2, strBillAddress3, strBillCity, strBillState, strBillZip,
                                            strShipToName, strShipPhone, strShipAddress1, strShipAddress2, "", strShipCity, strShipState, strShipZipCode,
                                            out strError);
                                    }

                                    if (!string.IsNullOrEmpty(strError))
                                    {
                                        throw new Exception(strError);
                                    }
                                    strCustomerId = objSubCustomer.internalId;
                                }
                                catch (Exception objExc)
                                {
                                    throw new Exception("Subcustomer hasn't created or updated (" + strShipToName + "): " + objExc.Message);
                                }
                            }*/

                            DateTime strShipDate = DateTime.Parse(objRow["Ship Date"].ToString());
                            DateTime strDeliDate = strShipDate;
                            string strClase = objRow["Class"].ToString();
                            try
                            {
                                strDeliDate = DateTime.Parse(objRow["DeliveryDate"].ToString());
                            }
                            catch (Exception)
                            {
                            }
                            DateTime strInvoiceDate = strDeliDate;
                            try
                            {
                                strInvoiceDate = DateTime.Parse(objRow["InvoiceDate"].ToString());
                            }
                            catch (Exception)
                            {
                            }
                            DateTime objOrderDate = DateTime.Now;
                            try
                            {
                                objOrderDate = DateTime.Parse(objRow["F1"].ToString());
                            }
                            catch (Exception)
                            {
                            }

                            double numQty = double.Parse(objRow["Q"].ToString());

                            if (!string.IsNullOrEmpty(strCustomer) && !string.IsNullOrEmpty(strProduct) && !string.IsNullOrEmpty(strPoNumber))
                            {
                                if (booIsByPO)
                                {
                                    for (int i = 0; i < numQty; i++)
                                    {
                                        arrItems.Add(new object[] { strProductId, 1, strShipDate, strDeliDate, arrShipType.IndexOf(strShipType) + 1, objRow["PO"].ToString(), "", arrCourierType.IndexOf(strCourierType) + 1 });
                                    }
                                }
                                else
                                {
                                    arrItems.Add(new object[] { strProductId, numQty, strShipDate, strDeliDate, arrShipType.IndexOf(strShipType) + 1, objRow["PO"].ToString(), "", arrCourierType.IndexOf(strCourierType) + 1 });
                                }
                                string strNextRef = "";
                                if (numIndex + 1 < dsTable.Rows.Count)
                                {
                                    strNextRef = dsTable.Rows[numIndex + 1][strRef].ToString();
                                    if (dsTable.Rows[numIndex + 1][strRef] is DateTime)
                                    {
                                        strNextRef = ((DateTime)dsTable.Rows[numIndex + 1][strRef]).ToString(strDateFormat);
                                    }
                                }
                                if (numIndex + 1 >= dsTable.Rows.Count || !strNextRef.Equals(strPoNumber))
                                {
                                    // Insert order
                                    string strError = "";
                                    strId = Util.AddSalesOrder(arrReseller.IndexOf(strCustomer) + 1, objSubCustomer, strPoNumber, strPoNumber, objOrderDate, DateTime.Now, objOrderDate, arrItems, strEmployeeId, strInvoiceDate, strClase, out strError);
                                    if (!string.IsNullOrEmpty(strId))
                                    {
                                        LastProcessResult = ProcessResult.Ok2;
                                        strMessage = strId;
                                        numUserIndex++;
                                        numUserIndex = (numUserIndex) % arrUsers.Count;
                                        if (!htOrders.ContainsKey(strPoNumber))
                                        {
                                            htOrders.Add(strPoNumber, strId);
                                        }
                                    }
                                    else
                                    {
                                        LastProcessResult = ProcessResult.Error;
                                        strMessage = strError;
                                    }
                                    arrItems.Clear();
                                }
                                else
                                {
                                    LastProcessResult = ProcessResult.Ok1;
                                    strMessage = "Linea agregada: " + strProduct;
                                }
                            }
                            else
                            {
                                LastProcessResult = ProcessResult.Error;
                                strMessage = "Información incompleta en SKU";
                            }
                        }
                        else
                        {
                            //objSalesOrder = (SalesOrder)Util.GetEntity(strId, RecordType.salesOrder);
                            //if (!htOrders.ContainsKey(strPoNumber))
                            //{
                            //    htOrders.Add(strPoNumber, strId);
                            //}
                            //strMessage = "Registro ya procesado: " + strId;
                            //LastProcessResult = ProcessResult.Processed;
                            objSalesOrder = (SalesOrder)Util.GetEntity(strId, RecordType.salesOrder);
                            numTotal = objSalesOrder.itemList.item.Length;
                            if (GlobalSettings.Default.AddLines.Equals("Yes"))
                            {
                                numTotal = 0;
                            }
                            if (!htOrders.ContainsKey(strPoNumber))
                            {
                                htOrders.Add(strPoNumber, strId);
                            }
                            booAdd = true;
                        }
                    }
                    else
                    {
                        booAdd = true;
                    }

                    if (booAdd)
                    {
                        if (htOrders[strPoNumber].Equals("error"))
                        {
                            LastProcessResult = ProcessResult.Error;
                            strMessage = "Linea omitida por error en linea anteriores a la orden";
                        }
                        else
                        {
                            // Depende del archivo de configuración si es True AddLines numTotal es 0
                            if (numTotal < 1)
                            {
                                //Adicional
                                string strProduct = objRow["SKU"].ToString();
                                if (string.IsNullOrEmpty(strProduct))
                                {
                                    throw new Exception("Product is empty.");
                                }
                                string strProductId = "";
                                if (!htProducts.ContainsKey(strProduct))
                                {
                                    ArrayList arrProducts = Util.GetItemId(strProduct);
                                    if (arrProducts.Count == 0)
                                    {
                                        throw new Exception("Product hasn't been found. Product: " + strProduct);
                                    }
                                    else if (arrProducts.Count == 1)
                                    {
                                        strProductId = ((string[])arrProducts[0])[0];
                                        htProducts.Add(strProduct, strProductId);
                                    }
                                    else
                                    {
                                        Form2 objForm2 = new Form2(arrProducts);
                                        System.Windows.Forms.DialogResult objResult = objForm2.ShowDialog(objForm);
                                        while (objResult != System.Windows.Forms.DialogResult.OK)
                                        {
                                            objForm2.ShowDialog(objForm);
                                        }
                                        if (objResult == System.Windows.Forms.DialogResult.OK)
                                        {
                                            strProductId = objForm2.ItemId;
                                        }
                                    }
                                }
                                else
                                {
                                    strProductId = htProducts[strProduct].ToString();
                                }

                                DateTime strShipDate = DateTime.Parse(objRow["Ship Date"].ToString());
                                DateTime strDeliDate = strShipDate;
                                try
                                {
                                    strDeliDate = DateTime.Parse(objRow["DeliveryDate"].ToString());
                                }
                                catch (Exception)
                                {
                                }
                                DateTime objOrderDate = DateTime.Now;
                                try
                                {
                                    objOrderDate = DateTime.Parse(objRow["F1"].ToString());
                                }
                                catch (Exception)
                                {
                                }
                                double numQty = double.Parse(objRow["Q"].ToString());
                                string strShipType = objRow["Transport1"].ToString().Trim();
                                string strCourierType = objRow["Transport2"].ToString().Trim();

                                if (booIsByPO)
                                {
                                    for (int i = 0; i < numQty; i++)
                                    {
                                        arrItems.Add(new object[] { strProductId, 1, strShipDate, strDeliDate, arrShipType.IndexOf(strShipType) + 1, objRow["PO"].ToString(), arrCourierType.IndexOf(strCourierType) + 1 });
                                    }
                                }
                                else
                                {
                                    arrItems.Add(new object[] { strProductId, numQty, strShipDate, strDeliDate, arrShipType.IndexOf(strShipType) + 1, objRow["PO"].ToString(), arrCourierType.IndexOf(strCourierType) + 1 });
                                }

                                string strNextRef = "";
                                if (numIndex + 1 < dsTable.Rows.Count)
                                {
                                    strNextRef = dsTable.Rows[numIndex + 1][strRef].ToString();
                                    if (dsTable.Rows[numIndex + 1][strRef] is DateTime)
                                    {
                                        strNextRef = ((DateTime)dsTable.Rows[numIndex + 1][strRef]).ToString(strDateFormat);
                                    }
                                }
                                if (numIndex + 1 >= dsTable.Rows.Count || !strNextRef.Equals(strPoNumber))
                                {
                                    // Insert order
                                    string strError = "";
                                    string strId = objSalesOrder.internalId;
                                    Util.UpdateSalesOrder(objSalesOrder, arrItems, out strError);
                                    if (!string.IsNullOrEmpty(strId))
                                    {
                                        LastProcessResult = ProcessResult.Added;
                                        strMessage = "Linea adicionada para la orden: " + strPoNumber;
                                        numUserIndex++;
                                        numUserIndex = (numUserIndex) % arrUsers.Count;
                                        if (!htOrders.ContainsKey(strPoNumber))
                                        {
                                            htOrders.Add(strPoNumber, strId);
                                        }
                                    }
                                    else
                                    {
                                        LastProcessResult = ProcessResult.Error;
                                        strMessage = strError;
                                    }
                                    arrItems.Clear();
                                }
                                else
                                {
                                    LastProcessResult = ProcessResult.Ok1;
                                    strMessage = "Linea agregada en la orden: " + strPoNumber;
                                }
                            }
                            else
                            {
                                strMessage = "Registro ya procesado: " + htOrders[strPoNumber];
                                LastProcessResult = ProcessResult.Processed;
                            }
                        }
                    }
                }
                catch (Exception objExc)
                {
                    arrItems.Clear();
                    if (!htOrders.ContainsKey(strPoNumber))
                    {
                        htOrders.Add(strPoNumber, "error");
                    }
                    else
                    {
                        htOrders[strPoNumber] = "error";
                    }
                    LastProcessResult = ProcessResult.Error;
                    strMessage = objExc.Message;
                    //ShowError(objExc.Message);
                }

                string strRefAux = "";
                if (numIndex + 1 < dsTable.Rows.Count)
                {
                    strRefAux = dsTable.Rows[numIndex + 1][strRef].ToString();
                    if (dsTable.Rows[numIndex + 1][strRef] is DateTime)
                    {
                        strRefAux = ((DateTime)dsTable.Rows[numIndex + 1][strRef]).ToString(strDateFormat);
                    }
                }
                if (numIndex + 1 >= dsTable.Rows.Count || !strRefAux.Equals(strPoNumber))
                {
                    numLineId = 0;
                }
                numIndex++;
                numLineId++;
            }

            objForm.IncrementProcess();
            RUN = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
        }

        private static void LoadTable1800()
        {
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.WaitCursor;
            Hashtable htCustomers = new Hashtable();
            Hashtable htProducts = new Hashtable();
            Hashtable htOrders = new Hashtable();
            Util.Login(GlobalSettings.Default.nsUser, GlobalSettings.Default.nsPassword, GlobalSettings.Default.nsAccount, GlobalSettings.Default.nsRole);

            int numUserIndex = 0;
            ArrayList arrUsers = Util.GetUsers(GlobalSettings.Default.nsSalesRoleId);
            if (arrUsers.Count == 0)
            {
                string[] arrTemp = GlobalSettings.Default.nsSalesRep.Split(new string[] { "," }, StringSplitOptions.None);
                foreach (string strTemp in arrTemp)
                {
                    Employee objEmp = new Employee();
                    objEmp.internalId = strTemp;
                    arrUsers.Add(objEmp);
                }
                //ShowError("SalesReps haven't been found. Role ID: " + GlobalSettings.Default.nsSalesRoleId);
                //return;
            }

            int numIndex = 0;
            ArrayList arrItems = new ArrayList();
            string strRef = "PO";
            bool booIsByPO = true;
            string strDateFormat = "yyyy-MM-dd";
            if (arrAcum.Contains(dsTable.TableName))
            {
                strDateFormat = "GLyyyy-MM-dd";
                strRef = "Ship Date";
                booIsByPO = false;
            }

            Customer objReseller = new Customer();
            int numLineId = 1;
            int numTotal = 0;
            SalesOrder objSalesOrder = new SalesOrder();
            foreach (DataRow objRow in dsTable.Rows)
            {
                string strEmployeeId = ((Employee)arrUsers[numUserIndex]).internalId;
                string strPoNumber = "";
                string strPoAux = "";
                bool booAdd = false;
                try
                {
                    objForm.IncrementProcess();
                    string strPOType = objRow["PO_TYPE"].ToString();
                    string strOrderNumber = objRow["CustPONumber"].ToString();
                    if (strPOType.Equals("CANCEL"))
                    {
                        SalesOrder objClosedSalesOrder = Util.GetOrderByExternalID("so1800_" + strOrderNumber);
                        if (objClosedSalesOrder == null)
                        {
                            strMessage = "Sales Order not found. ExternalID: " + "so1800_" + strOrderNumber;
                            LastProcessResult = ProcessResult.Error;
                        }
                        else
                        {
                            strMessage = Util.CloseSalesOrder(objClosedSalesOrder);
                            if (string.IsNullOrEmpty(strMessage))
                            {
                                strMessage = "Sales Order is Closed. InternalID: " + objClosedSalesOrder.internalId;
                                LastProcessResult = ProcessResult.Closed;
                            }
                            else
                            {
                                LastProcessResult = ProcessResult.Error;
                            }
                        }
                    }
                    else
                    {
                        strPoNumber = objRow[strRef].ToString();
                        strPoAux = objRow["PO"].ToString();
                        if (objRow[strRef] is DateTime)
                        {
                            strPoNumber = ((DateTime)objRow[strRef]).ToString(strDateFormat);
                        }
                        if (!htOrders.ContainsKey(strPoNumber))
                        {
                            string strCustomer = objRow["Customer"].ToString();
                            string strId = Util.GetOrderId(strPoNumber, "51");
                            if (string.IsNullOrEmpty(strId))
                            {
                                if (!htCustomers.ContainsKey(strCustomer))
                                {
                                    objReseller = Util.GetCustomer(strCustomer);
                                    if (objReseller == null || objReseller.internalId == null)
                                    {
                                        throw new Exception("Customer hasn't been found. Reseller ID: " + strCustomer);
                                    }
                                    else
                                    {
                                        htCustomers.Add(strCustomer, objReseller.internalId);
                                    }
                                }

                                string strBillTo = "";
                                string strBillPhone = "";
                                string strBillAddress1 = "";
                                string strBillAddress2 = "";
                                string strBillAddress3 = "";
                                string strBillCity = "";
                                string strBillState = "";
                                string strBillZip = "";
                                foreach (CustomerAddressbook objAddress in objReseller.addressbookList.addressbook)
                                {
                                    if (objAddress.defaultBilling)
                                    {
                                        strBillTo = objAddress.attention;
                                        strBillPhone = objAddress.phone;
                                        strBillAddress1 = objAddress.addr1;
                                        strBillAddress2 = objAddress.addr2;
                                        strBillAddress3 = objAddress.addr3;
                                        strBillCity = objAddress.city;
                                        strBillState = objAddress.state;
                                        strBillZip = objAddress.zip;
                                    }
                                }

                                string strCustomerId = htCustomers[strCustomer].ToString();

                                string strProduct = objRow["SKU"].ToString();
                                if (string.IsNullOrEmpty(strProduct))
                                {
                                    throw new Exception("Product is empty.");
                                }
                                string strProductId = "";
                                if (!htProducts.ContainsKey(strProduct))
                                {
                                    ArrayList arrProducts = Util.GetItemByItemNumber(strProduct);
                                    if (arrProducts.Count == 0)
                                    {
                                        throw new Exception("Product hasn't been found. Product: " + strProduct);
                                    }
                                    else if (arrProducts.Count == 1)
                                    {
                                        strProductId = ((string[])arrProducts[0])[0];
                                        htProducts.Add(strProduct, strProductId);
                                    }
                                    else
                                    {
                                        Form2 objForm2 = new Form2(arrProducts);
                                        System.Windows.Forms.DialogResult objResult = objForm2.ShowDialog(objForm);
                                        while (objResult != System.Windows.Forms.DialogResult.OK)
                                        {
                                            objForm2.ShowDialog(objForm);
                                        }
                                        if (objResult == System.Windows.Forms.DialogResult.OK)
                                        {
                                            strProductId = objForm2.ItemId;
                                        }
                                    }
                                }
                                else
                                {
                                    strProductId = htProducts[strProduct].ToString();
                                }

                                string strShipToName = objRow["ShipToName"].ToString();
                                string strShipPhone = objRow["ShipToPhone"].ToString();

                                string strShipAddress1 = objRow["ShipToAddress1"].ToString();
                                string strShipAddress2 = objRow["ShipToAddress2"].ToString();
                                string strShipCity = objRow["ShipToCity"].ToString();
                                string strShipZipCode = objRow["ShipToZipCode"].ToString();
                                string strShipState = objRow["ShipToState"].ToString();
                                string strShipEmail = objRow["ShipToEmail"].ToString();
                                string strShipType = objRow["Transport1"].ToString().Trim();
                                string strCourierType = objRow["Transport2"].ToString().Trim();

                                //RecordRef objSubCustomer = new RecordRef();
                                //objSubCustomer.internalId = objReseller.internalId;
                                //objSubCustomer.type = RecordType.customer;
                                //objSubCustomer.typeSpecified = true;

                                Customer objSubCustomer = new Customer();
                                if (!string.IsNullOrEmpty(strShipToName))
                                {
                                    string strFirstName = "";
                                    string strLastName = "";

                                    string[] arrNames = strShipToName.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                                    if (arrNames.Length == 1 || arrNames.Length != 0)
                                    {
                                        arrNames = arrNames[0].Split(new string[] { " " }, StringSplitOptions.RemoveEmptyEntries);
                                        if (arrNames.Length > 1)
                                        {
                                            strFirstName = arrNames[0];
                                            strLastName = arrNames[1];
                                        }
                                        else
                                        {
                                            strFirstName = strShipToName;
                                        }
                                    }
                                    else
                                    {
                                        strFirstName = strShipToName;
                                    }
                                    try
                                    {
                                        string strError = "";
                                        //objSubCustomer = Util.GetCustomer(strCustomerId, strShipPhone);
                                        objSubCustomer = null;

                                        if (string.IsNullOrEmpty(strBillTo))
                                        {
                                            throw new Exception("Billing information missed!");
                                        }

                                        if (objSubCustomer == null || objSubCustomer.internalId == null)
                                        {
                                            objSubCustomer = Util.AddCustomer1800("cu1800_" + strOrderNumber, strCustomerId, strEmployeeId, strFirstName, strLastName, strShipEmail,
                                                strBillTo, strBillPhone, strBillAddress1, strBillAddress2, strBillAddress3, strBillCity, strBillState, strBillZip,
                                                strShipToName, strShipPhone, strShipAddress1, strShipAddress2, "", strShipCity, strShipState, strShipZipCode,
                                                out strError);
                                        }
                                        else
                                        {
                                            /*objSubCustomer = Util.UpdateCustomer(objSubCustomer.internalId, strEmployeeId, strCustomerId, strFirstName, strLastName, strShipEmail,
                                                strBillTo, strBillPhone, strBillAddress1, strBillAddress2, strBillAddress3, strBillCity, strBillState, strBillZip,
                                                strShipToName, strShipPhone, strShipAddress1, strShipAddress2, "", strShipCity, strShipState, strShipZipCode,
                                                out strError);*/
                                            throw new Exception("Function not implemented!");
                                        }

                                        if (!string.IsNullOrEmpty(strError))
                                        {
                                            throw new Exception(strError);
                                        }
                                        strCustomerId = objSubCustomer.internalId;
                                    }
                                    catch (Exception objExc)
                                    {
                                        throw new Exception("Subcustomer hasn't created or updated (" + strShipToName + "): " + objExc.Message);
                                    }
                                }

                                DateTime strShipDate = DateTime.Parse(objRow["ShipDate"].ToString());
                                DateTime strDeliDate = strShipDate;
                                try
                                {
                                    strDeliDate = DateTime.Parse(objRow["DeliveryDate"].ToString());
                                }
                                catch (Exception)
                                {
                                }
                                DateTime objOrderDate = DateTime.Now;
                                try
                                {
                                    string strOrderDate = ((DateTime)objRow["DATECRTD"]).ToString("MM/dd/yyyy") + " " + ((DateTime)objRow["TIMECRTD"]).ToString("HH:mm:ss");
                                    objOrderDate = DateTime.Parse(strOrderDate);
                                }
                                catch (Exception)
                                {
                                }
                                DateTime objInsertDate = DateTime.Now;
                                try
                                {
                                    string strInsertDate = ((DateTime)objRow["FILEDATE"]).ToString("MM/dd/yyyy") + " " + ((DateTime)objRow["FILETIME"]).ToString("HH:mm:ss");
                                    objInsertDate = DateTime.Parse(strInsertDate);
                                }
                                catch (Exception)
                                {
                                }

                                double numQty = double.Parse(objRow["Q"].ToString());

                                if (!string.IsNullOrEmpty(strCustomer) && !string.IsNullOrEmpty(strProduct) && !string.IsNullOrEmpty(strPoNumber))
                                {
                                    if (booIsByPO)
                                    {
                                        for (int i = 0; i < numQty; i++)
                                        {
                                            arrItems.Add(new object[] { strProductId, 1, strShipDate, strDeliDate, arrShipType.IndexOf(strShipType) + 1, objRow["PO"].ToString(), objRow["ORDERMESSAGETEXT"].ToString(), arrCourierType.IndexOf(strCourierType) + 1 });
                                        }
                                    }
                                    else
                                    {
                                        arrItems.Add(new object[] { strProductId, numQty, strShipDate, strDeliDate, arrShipType.IndexOf(strShipType) + 1, objRow["PO"].ToString(), objRow["ORDERMESSAGETEXT"].ToString(), arrCourierType.IndexOf(strCourierType) + 1 });
                                    }
                                    string strNextRef = "";
                                    if (numIndex + 1 < dsTable.Rows.Count)
                                    {
                                        strNextRef = dsTable.Rows[numIndex + 1][strRef].ToString();
                                        if (dsTable.Rows[numIndex + 1][strRef] is DateTime)
                                        {
                                            strNextRef = ((DateTime)dsTable.Rows[numIndex + 1][strRef]).ToString(strDateFormat);
                                        }
                                    }
                                    if (numIndex + 1 >= dsTable.Rows.Count || !strNextRef.Equals(strPoNumber))
                                    {
                                        // Insert order
                                        string strError = "";
                                        string strVendorCode = objRow["VendorCode"].ToString();
                                        strId = Util.AddSalesOrder1800(51, objSubCustomer, strVendorCode, strPoNumber, strOrderNumber, objOrderDate, objInsertDate, objOrderDate, arrItems, strEmployeeId, out strError);
                                        if (!string.IsNullOrEmpty(strId))
                                        {
                                            LastProcessResult = ProcessResult.Ok2;
                                            strMessage = strId;
                                            numUserIndex++;
                                            numUserIndex = (numUserIndex) % arrUsers.Count;
                                            if (!htOrders.ContainsKey(strPoNumber))
                                            {
                                                htOrders.Add(strPoNumber, strId);
                                            }
                                        }
                                        else
                                        {
                                            LastProcessResult = ProcessResult.Error;
                                            strMessage = strError;
                                        }
                                        arrItems.Clear();
                                    }
                                    else
                                    {
                                        LastProcessResult = ProcessResult.Ok1;
                                        strMessage = "Linea agregada: " + strProduct;
                                    }
                                }
                                else
                                {
                                    LastProcessResult = ProcessResult.Error;
                                    strMessage = "Información incompleta en SKU";
                                }
                            }
                            else
                            {
                                //objSalesOrder = (SalesOrder)Util.GetEntity(strId, RecordType.salesOrder);
                                //if (!htOrders.ContainsKey(strPoNumber))
                                //{
                                //    htOrders.Add(strPoNumber, strId);
                                //}
                                //strMessage = "Registro ya procesado: " + strId;
                                //LastProcessResult = ProcessResult.Processed;
                                objSalesOrder = (SalesOrder)Util.GetEntity(strId, RecordType.salesOrder);
                                numTotal = objSalesOrder.itemList.item.Length;
                                if (GlobalSettings.Default.AddLines.Equals("Yes"))
                                {
                                    numTotal = 0;
                                }
                                if (!htOrders.ContainsKey(strPoNumber))
                                {
                                    htOrders.Add(strPoNumber, strId);
                                }
                                booAdd = true;
                            }
                        }
                        else
                        {
                            booAdd = true;
                        }

                        if (booAdd)
                        {
                            if (htOrders[strPoNumber].Equals("error"))
                            {
                                LastProcessResult = ProcessResult.Error;
                                strMessage = "Linea omitida por error en linea anteriores a la orden";
                            }
                            else
                            {
                                // Depende del archivo de configuración si es True AddLines numTotal es 0
                                if (numTotal < 1)
                                {
                                    //Adicional
                                    string strProduct = objRow["SKU"].ToString();
                                    if (string.IsNullOrEmpty(strProduct))
                                    {
                                        throw new Exception("Product is empty.");
                                    }
                                    string strProductId = "";
                                    if (!htProducts.ContainsKey(strProduct))
                                    {
                                        ArrayList arrProducts = Util.GetItemId(strProduct);
                                        if (arrProducts.Count == 0)
                                        {
                                            throw new Exception("Product hasn't been found. Product: " + strProduct);
                                        }
                                        else if (arrProducts.Count == 1)
                                        {
                                            strProductId = ((string[])arrProducts[0])[0];
                                            htProducts.Add(strProduct, strProductId);
                                        }
                                        else
                                        {
                                            Form2 objForm2 = new Form2(arrProducts);
                                            System.Windows.Forms.DialogResult objResult = objForm2.ShowDialog(objForm);
                                            while (objResult != System.Windows.Forms.DialogResult.OK)
                                            {
                                                objForm2.ShowDialog(objForm);
                                            }
                                            if (objResult == System.Windows.Forms.DialogResult.OK)
                                            {
                                                strProductId = objForm2.ItemId;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        strProductId = htProducts[strProduct].ToString();
                                    }

                                    DateTime strShipDate = DateTime.Parse(objRow["Ship Date"].ToString());
                                    DateTime strDeliDate = strShipDate;
                                    try
                                    {
                                        strDeliDate = DateTime.Parse(objRow["DeliveryDate"].ToString());
                                    }
                                    catch (Exception)
                                    {
                                    }
                                    DateTime objOrderDate = DateTime.Now;
                                    try
                                    {
                                        objOrderDate = DateTime.Parse(objRow["F1"].ToString());
                                    }
                                    catch (Exception)
                                    {
                                    }
                                    double numQty = double.Parse(objRow["Q"].ToString());
                                    string strShipType = objRow["Transport1"].ToString().Trim();
                                    string strCourierType = objRow["Transport2"].ToString().Trim();

                                    if (booIsByPO)
                                    {
                                        for (int i = 0; i < numQty; i++)
                                        {
                                            arrItems.Add(new object[] { strProductId, 1, strShipDate, strDeliDate, arrShipType.IndexOf(strShipType) + 1, objRow["PO"].ToString(), arrCourierType.IndexOf(strCourierType) + 1 });
                                        }
                                    }
                                    else
                                    {
                                        arrItems.Add(new object[] { strProductId, numQty, strShipDate, strDeliDate, arrShipType.IndexOf(strShipType) + 1, objRow["PO"].ToString(), arrCourierType.IndexOf(strCourierType) + 1 });
                                    }

                                    string strNextRef = "";
                                    if (numIndex + 1 < dsTable.Rows.Count)
                                    {
                                        strNextRef = dsTable.Rows[numIndex + 1][strRef].ToString();
                                        if (dsTable.Rows[numIndex + 1][strRef] is DateTime)
                                        {
                                            strNextRef = ((DateTime)dsTable.Rows[numIndex + 1][strRef]).ToString(strDateFormat);
                                        }
                                    }
                                    if (numIndex + 1 >= dsTable.Rows.Count || !strNextRef.Equals(strPoNumber))
                                    {
                                        // Insert order
                                        string strError = "";
                                        string strId = objSalesOrder.internalId;
                                        Util.UpdateSalesOrder(objSalesOrder, arrItems, out strError);
                                        if (!string.IsNullOrEmpty(strId))
                                        {
                                            LastProcessResult = ProcessResult.Added;
                                            strMessage = "Linea adicionada para la orden: " + strPoNumber;
                                            numUserIndex++;
                                            numUserIndex = (numUserIndex) % arrUsers.Count;
                                            if (!htOrders.ContainsKey(strPoNumber))
                                            {
                                                htOrders.Add(strPoNumber, strId);
                                            }
                                        }
                                        else
                                        {
                                            LastProcessResult = ProcessResult.Error;
                                            strMessage = strError;
                                        }
                                        arrItems.Clear();
                                    }
                                    else
                                    {
                                        LastProcessResult = ProcessResult.Ok1;
                                        strMessage = "Linea agregada en la orden: " + strPoNumber;
                                    }
                                }
                                else
                                {
                                    strMessage = "Registro ya procesado: " + htOrders[strPoNumber];
                                    LastProcessResult = ProcessResult.Processed;
                                }
                            }
                        }
                    }
                }
                catch (Exception objExc)
                {
                    arrItems.Clear();
                    if (!htOrders.ContainsKey(strPoNumber))
                    {
                        htOrders.Add(strPoNumber, "error");
                    }
                    else
                    {
                        htOrders[strPoNumber] = "error";
                    }
                    LastProcessResult = ProcessResult.Error;
                    strMessage = objExc.Message;
                    //ShowError(objExc.Message);
                }

                string strRefAux = "";
                if (numIndex + 1 < dsTable.Rows.Count)
                {
                    strRefAux = dsTable.Rows[numIndex + 1][strRef].ToString();
                    if (dsTable.Rows[numIndex + 1][strRef] is DateTime)
                    {
                        strRefAux = ((DateTime)dsTable.Rows[numIndex + 1][strRef]).ToString(strDateFormat);
                    }
                }
                if (numIndex + 1 >= dsTable.Rows.Count || !strRefAux.Equals(strPoNumber))
                {
                    numLineId = 0;
                }
                numIndex++;
                numLineId++;
            }

            objForm.IncrementProcess();
            RUN = false;
            System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default;
        }

        private static bool LoadTextData(string strFilelocation, string strFileName)
        {
            bool booResult = true;
            string textConnStr = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=\"" + strFilelocation + "\";Extended Properties=\"text;HDR=Yes;FMT=Delimited\";";
            OleDbConnection textConn = new OleDbConnection(textConnStr);
            try
            {
                CultureInfo culture = Thread.CurrentThread.CurrentCulture;
                OleDbCommand excelCommand = new OleDbCommand();
                OleDbDataAdapter excelDataAdapter = new OleDbDataAdapter();
                //string excelConnStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + filelocation + "; Extended Properties =Excel 8.0;";

                textConn.Open();
                DataTable dtPatterns = new DataTable();
                excelCommand = new OleDbCommand("SELECT * FROM [" + strFileName + "]", textConn);
                excelDataAdapter.SelectCommand = excelCommand;
                excelDataAdapter.Fill(dtPatterns);

                dtPatterns.TableName = "format1800";
                dsData.Tables.Add(dtPatterns);
            }
            catch (Exception objExc)
            {
                dsData.Tables.Add(new DataTable("format1800"));
                ShowError(objExc.Message);
                booResult = false;
            }
            textConn.Close();
            return booResult;
        }

        private static void LoadTextData2(string strFilelocation, string strFileName)
        {
            try
            {
                CultureInfo culture = Thread.CurrentThread.CurrentCulture;
                DataTable dtPatterns = new DataTable();

                System.IO.StreamReader objReader = new System.IO.StreamReader(strFilelocation);
                string strHeader = objReader.ReadLine();
                string[] arrHeader = strHeader.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < arrHeader.Length; i++)
                {
                    dtPatterns.Columns.Add(arrHeader[i], "".GetType());
                }
                string strLine = objReader.ReadLine();
                while (!string.IsNullOrEmpty(strLine))
                {
                    string[] arrDetail = strLine.Split(new string[] { "," }, StringSplitOptions.RemoveEmptyEntries);
                    dtPatterns.Rows.Add(arrDetail);
                    strLine = objReader.ReadLine();
                }

                dtPatterns.TableName = "format1800";
                dsData.Tables.Add(dtPatterns);
            }
            catch (Exception objExc)
            {
                dsData.Tables.Add(new DataTable("format1800"));
            }
        }

        private static void LoadExcelData(string filelocation, string strCustomer)
        {
            string excelConnStr = "Provider=Microsoft.ACE.OLEDB.12.0; Data Source=" + filelocation + "; Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";";
            OleDbConnection excelConn = new OleDbConnection(excelConnStr);
            try
            {
                CultureInfo culture = Thread.CurrentThread.CurrentCulture;
                OleDbCommand excelCommand = new OleDbCommand();
                //OleDbDataAdapter excelDataAdapter = new OleDbDataAdapter();
                OleDbDataAdapter excelDataAdapter;
                //string excelConnStr = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" + filelocation + "; Extended Properties =Excel 8.0;";

                excelConn.Open();
                DataTable dtPatterns = new DataTable();
                excelCommand = new OleDbCommand("SELECT * FROM [" + strCustomer + "$] ", excelConn);
                excelDataAdapter = new OleDbDataAdapter(excelCommand);
                //excelDataAdapter.SelectCommand = excelCommand;
                excelDataAdapter.Fill(dtPatterns);

                dtPatterns.TableName = strCustomer;
                dsData.Tables.Add(dtPatterns);
            }
            catch (Exception objExc)
            {
                dsData.Tables.Add(new DataTable(strCustomer));
            }
            excelConn.Close();
        }
                
        public static void ShowError(string strText)
        {
            System.Windows.Forms.MessageBox.Show(strText, "Error", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning);
        }

        public static void ShowMessage(string strText)
        {
            System.Windows.Forms.MessageBox.Show(strText, "Info", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information);
        }

    }
}
