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



    }
}
