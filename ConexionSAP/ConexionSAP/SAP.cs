using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SAPbobsCOM;

namespace ConexionSAP
{
    class SAP
    {
        private Company oCom;
        public string Error="";
        public string CName = "";

        public SAP()
        {
            this.oCom = new Company();
        }


        public void Conectar()
        {
            try
            {
                this.oCom.Server = "LABORATORIO";
                this.oCom.DbServerType = BoDataServerTypes.dst_MSSQL2014;
                this.oCom.UserName = "manager";
                this.oCom.Password = "manager";
                this.oCom.CompanyDB = "SBODemoGT";

                //this.oCom.DbUserName = "sa";
                //this.oCom.DbPassword = "SAPB1Admin1";

                if (!this.oCom.Connected)
                {
                    int ErrorCode = this.oCom.Connect();

                    if (ErrorCode != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription() + "( " + ErrorCode.ToString() + ")";
                    }
                    else
                    {
                        this.CName = this.oCom.CompanyName;
                    }
                }
                else
                {
                    this.Error = "Ya esta conectado";
                }



            }
            catch (Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {

            }
        }



        public void Desconectar()
        {
            try
            {
                this.Error = "";
                if (this.oCom != null)
                {
                    this.oCom.Disconnect();
                }
            }catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (this.oCom != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(this.oCom);
                    this.oCom = null;
                }
                
            }
        }

        #region SN

        public void CrearSN()
        {
            SAPbobsCOM.BusinessPartners oSN = null;
            try
            {
                this.Error = "";
                oSN = (SAPbobsCOM.BusinessPartners)this.oCom.GetBusinessObject(BoObjectTypes.oBusinessPartners);
                oSN.CardCode = "CL03";
                oSN.CardName = "Cliente TEST";
                oSN.CardType = BoCardTypes.cCustomer;
                oSN.FederalTaxID = "123456789023";

                if (oSN.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription();
                }

            }catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oSN != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSN);
                    oSN = null;
                }
            }
        }


        public void EditarSN(string CardCode)
        {
            SAPbobsCOM.BusinessPartners oSN = null;
            try
            {
                this.Error = "";
                oSN = (SAPbobsCOM.BusinessPartners)this.oCom.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                if (oSN.GetByKey(CardCode))
                {
                    oSN.EmailAddress = "TEST@gmail.com";

                    if (oSN.Update() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription();
                    }
                }
                else
                {
                    this.Error = "SN no existe";
                }

            }catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oSN != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSN);
                    oSN = null;
                }
            }
        }


        public void ADDContactosSN(string CardCode)
        {
            SAPbobsCOM.BusinessPartners oSN = null;
            try
            {
                this.Error = "";
                oSN = (SAPbobsCOM.BusinessPartners)this.oCom.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                if (oSN.GetByKey(CardCode))
                {

                    if (oSN.ContactEmployees.Count > 1)
                    {
                        oSN.ContactEmployees.Add();
                    }else
                    {
                        if (oSN.ContactEmployees.Name != "")
                        {
                            oSN.ContactEmployees.Add();
                        }
                    }
                    
                    oSN.ContactEmployees.Name = "PRIMERO";
                    oSN.ContactEmployees.FirstName = "Pedro2";
                    oSN.ContactEmployees.MiddleName = "Juan2";
                    oSN.ContactEmployees.LastName = "Perez2";
                    oSN.ContactEmployees.Title = "Sr..";
                    //oSN.ContactEmployees.Address = "Test";
                    //oSN.ContactEmployees.Phone1 = "89451335";
                    //oSN.ContactEmployees.MobilePhone = "56784512";
                    //oSN.ContactEmployees.E_Mail = "Test@gmail.com";

                    if (oSN.Update() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription();
                    }
                }
                else
                {
                    this.Error = "SN no existe";
                }

            }
            catch (Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oSN != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSN);
                    oSN = null;
                }
            }
        }


        public void EDITContactosSN(string CardCode,int Line)
        {
            SAPbobsCOM.BusinessPartners oSN = null;
            try
            {
                this.Error = "";
                oSN = (SAPbobsCOM.BusinessPartners)this.oCom.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                if (oSN.GetByKey(CardCode))
                {


                    oSN.ContactEmployees.SetCurrentLine(Line);
                    
                    oSN.ContactEmployees.FirstName = "";
                    oSN.ContactEmployees.MiddleName = "";
                    oSN.ContactEmployees.LastName = "";
                    oSN.ContactEmployees.Title = "";
                    //oSN.ContactEmployees.Address = "Test";
                    //oSN.ContactEmployees.Phone1 = "89451335";
                    //oSN.ContactEmployees.MobilePhone = "56784512";
                    //oSN.ContactEmployees.E_Mail = "Test@gmail.com";

                    if (oSN.Update() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription();
                    }
                }
                else
                {
                    this.Error = "SN no existe";
                }

            }
            catch (Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oSN != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSN);
                    oSN = null;
                }
            }
        }


        public void agregarDireccionSN(string CardCode)
        {
            SAPbobsCOM.BusinessPartners oSN = null;
            try
            {
                this.Error = "";
                oSN = (SAPbobsCOM.BusinessPartners)this.oCom.GetBusinessObject(BoObjectTypes.oBusinessPartners);

                if (oSN.GetByKey(CardCode))
                {
                    oSN.Addresses.Add();
                    oSN.Addresses.AddressName = "Dirección 3";
                    oSN.Addresses.AddressType = SAPbobsCOM.BoAddressType.bo_ShipTo;
                    

                    if (oSN.Update() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription();
                    }
                }else
                {
                    this.Error = "SN no encontrado";
                }

               

            }
            catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oSN != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oSN);
                    oSN = null;
                }
            }
        }


        #endregion

        #region ITEM    

        public void CrearItem()
        {
            SAPbobsCOM.Items oItem = null;
            try
            {
                oItem = (SAPbobsCOM.Items)this.oCom.GetBusinessObject(BoObjectTypes.oItems);
                oItem.ItemCode = "ART001";
                oItem.ItemName = "PALA";

                oItem.InventoryItem = SAPbobsCOM.BoYesNoEnum.tNO;

                oItem.WhsInfo.WarehouseCode = "01";

                if (oItem.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription();
                }

                this.Error = "";

            }catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oItem != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                    oItem = null;
                }
            }
        }

        public void EditarItem(string ItemCode)
        {
            SAPbobsCOM.Items oItem = null;
            try
            {
                oItem = (SAPbobsCOM.Items)this.oCom.GetBusinessObject(BoObjectTypes.oItems);

                if (oItem.GetByKey(ItemCode))
                {
                    oItem.WhsInfo.WarehouseCode = "5";
                    oItem.WhsInfo.Add();


                    if (oItem.Update() != 0)
                    {
                        this.Error = this.oCom.GetLastErrorDescription();
                    }
                }else
                {
                    this.Error = "Item no existe";
                }
               

                this.Error = "";

            }
            catch (Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oItem != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oItem);
                    oItem = null;
                }
            }
        }


        #endregion


        #region DOCUMENTOS

        #region PEDIDO

        public void CrearPedido(out string DocEntry)
        {
            DocEntry = "";
            SAPbobsCOM.Documents oPedido=null;
            try
            {
                this.Error = "";
                oPedido = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oOrders);
                //oPedido.Series = 02;
                oPedido.CardCode = "CL01";
                oPedido.DocDate = DateTime.Today;
                oPedido.DocDueDate = DateTime.Today;
                oPedido.Comments = "test 2";

                oPedido.Lines.ItemCode = "A00004";
                oPedido.Lines.Quantity = 2;
                oPedido.Lines.TaxCode = "IVA";

                oPedido.Lines.Add();

                oPedido.Lines.ItemCode = "A00005";
                oPedido.Lines.Quantity = 1;
                oPedido.Lines.TaxCode = "EXE";

                if (oPedido.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription();
                }else
                {
                    DocEntry = this.oCom.GetNewObjectKey();
                }
            }
            catch(Exception e)
            {
                this.Error= e.Message;
            }
            finally
            {
                if (oPedido != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPedido);
                    oPedido = null;
                }
            }
        }

        public void agregarLineaPedido(int DocEntry)
        {
            SAPbobsCOM.Documents oPedido=null;
            try
            {
                this.Error = "";
                oPedido = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oOrders);

                if (oPedido.GetByKey(DocEntry))
                {
                    if (oPedido.DocumentStatus != SAPbobsCOM.BoStatus.bost_Close)
                    {
                        oPedido.Lines.Add();
                        oPedido.Lines.ItemCode = "B10000";
                        oPedido.Lines.Quantity = 1;

                        if (oPedido.Update() != 0)
                        {
                            this.Error = this.oCom.GetLastErrorDescription();
                        }
                    }else
                    {
                        this.Error = "Documento Cerrado";
                    }
                   

                }
                else
                {
                    this.Error = "Pedido no existe";
                }

            }catch(Exception e)
            {

            }
            finally
            {

            }
        }


        public void CrearPedidoDeTipoServicio(out string DocEntry)
        {
            DocEntry = "";
            SAPbobsCOM.Documents oPedido = null;
            try
            {
                this.Error = "";
                oPedido = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oOrders);
                //oPedido.Series = 02;
                oPedido.CardCode = "CL01";
                oPedido.DocDate = DateTime.Today;
                oPedido.DocDueDate = DateTime.Today;
                oPedido.Comments = "Pedido de tipo servicio";

                oPedido.DocType = SAPbobsCOM.BoDocumentTypes.dDocument_Service;

                oPedido.Lines.ItemDescription = "Servicio de ejemplo";
                oPedido.Lines.AccountCode = "_SYS00000000361";
                oPedido.Lines.LineTotal = 500;
                

                if (oPedido.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription();
                }
                else
                {
                    DocEntry = this.oCom.GetNewObjectKey();
                }
            }
            catch (Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oPedido != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oPedido);
                    oPedido = null;
                }
            }
        }


        #endregion

        #region ENTREGAS

        public void CrearEntrega(out string DocEntry)
        {
            DocEntry = "";
            SAPbobsCOM.Documents oEntrega = null;
            try
            {
                this.Error = "";
                oEntrega = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oDeliveryNotes);
                oEntrega.CardCode = "CL01";
                oEntrega.DocDate = DateTime.Today;
                oEntrega.DocDueDate = DateTime.Today;
                oEntrega.Comments = "test 2";

                oEntrega.Lines.ItemCode = "A00004";
                oEntrega.Lines.Quantity = 2;
                oEntrega.Lines.TaxCode = "IVA";

                oEntrega.Lines.Add();

                oEntrega.Lines.ItemCode = "A00005";
                oEntrega.Lines.Quantity = 1;
                oEntrega.Lines.TaxCode = "EXE";

                if (oEntrega.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription();
                }
                else
                {
                    DocEntry = this.oCom.GetNewObjectKey();
                }
            }
            catch (Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oEntrega != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oEntrega);
                    oEntrega = null;
                }
            }
        }

        #endregion

        #region DEVOLUCION

        public void CrearDevolucion(out string DocEntry)
        {
            DocEntry = "";
            SAPbobsCOM.Documents oDevolucion = null;
            try
            {
                this.Error = "";
                oDevolucion = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oReturns);
                oDevolucion.CardCode = "CL01";
                oDevolucion.DocDate = DateTime.Today;
                oDevolucion.DocDueDate = DateTime.Today;
                oDevolucion.Comments = "test 2";

                oDevolucion.Lines.ItemCode = "A00004";
                oDevolucion.Lines.Quantity = 2;
                oDevolucion.Lines.TaxCode = "IVA";

                

                if (oDevolucion.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription();
                }
                else
                {
                    DocEntry = this.oCom.GetNewObjectKey();
                }
            }
            catch (Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {
                if (oDevolucion != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oDevolucion);
                    oDevolucion = null;
                }
            }
        }

        #endregion

        #region SALIDAS

        public void CrearSalida(out string DocEntry)
        {
            DocEntry = "";
            SAPbobsCOM.Documents oSalida = null;
            try
            {
                this.Error = "";
                oSalida = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oInventoryGenExit);
                oSalida.DocDate = DateTime.Today;
                oSalida.DocDueDate = DateTime.Today;
                oSalida.GroupNumber = -1;

                oSalida.Lines.ItemCode = "B10000";
                oSalida.Lines.Quantity = 2;
                oSalida.Lines.CostingCode = "10001";
                oSalida.Lines.CostingCode2 = "20001";

                //Lotes
                oSalida.Lines.BatchNumbers.SetCurrentLine(0);
                oSalida.Lines.BatchNumbers.Quantity = 2;
                oSalida.Lines.BatchNumbers.BatchNumber = "L01";

                if (oSalida.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription();
                }else
                {
                    DocEntry = this.oCom.GetNewObjectKey();
                }

            }catch(Exception e)
            {

            }
            finally
            {

            }
        }

        #endregion

        #region FACTURA CON DOCUMENTO BASE

        public void CrearFacturaConDocumentoBase(out string DocEntry)
        {
            DocEntry = "";
            SAPbobsCOM.Documents oFacturas = null;
            try
            {
                this.Error = "";
                oFacturas = (SAPbobsCOM.Documents)this.oCom.GetBusinessObject(BoObjectTypes.oInvoices);
                oFacturas.CardCode = "CL01";
                oFacturas.DocDate = DateTime.Today;
                oFacturas.DocDueDate = DateTime.Today;

                oFacturas.Lines.BaseType = (int)SAPbobsCOM.BoObjectTypes.oOrders;
                oFacturas.Lines.BaseEntry = 555;
                oFacturas.Lines.BaseLine = 0;
                oFacturas.Lines.TaxCode = "IVA";

                oFacturas.Lines.Add();
                oFacturas.Lines.BaseType = (int)SAPbobsCOM.BoObjectTypes.oOrders;
                oFacturas.Lines.BaseEntry = 555;
                oFacturas.Lines.BaseLine = 1;
                oFacturas.Lines.TaxCode = "IVA";

                oFacturas.Lines.Add();
                oFacturas.Lines.BaseType = (int)SAPbobsCOM.BoObjectTypes.oOrders;
                oFacturas.Lines.BaseEntry = 555;
                oFacturas.Lines.BaseLine = 2;
                oFacturas.Lines.TaxCode = "IVA";

                oFacturas.Lines.BatchNumbers.SetCurrentLine(0);
                oFacturas.Lines.BatchNumbers.Quantity = 1;
                oFacturas.Lines.BatchNumbers.BatchNumber = "L01";

                if (oFacturas.Add() != 0)
                {
                    this.Error = this.oCom.GetLastErrorDescription();
                }else
                {
                    DocEntry = this.oCom.GetNewObjectKey();
                }

            }
            catch(Exception e)
            {
                this.Error = e.Message;
            }
            finally
            {

            }
        }

        #endregion

        #endregion


    }
}
