using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ConexionSAP
{
    public partial class Form1 : Form
    {
        private SAP sap;
        public Form1()
        {
            InitializeComponent();
            sap = new SAP();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                sap.Conectar();
                if (sap.Error != "")
                {
                    MessageBox.Show("Error: "+sap.Error);
                }else
                {
                    MessageBox.Show("Conectados a " + sap.CName);
                }
            }
            catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                this.sap.Desconectar();
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error " + this.sap.Error);
                }else
                {
                    MessageBox.Show("Desconectados");
                }
            }catch(Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                this.sap.CrearSN();
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + this.sap.Error);
                }
            }catch(Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                this.sap.EditarSN("CL01");
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + this.sap.Error);
                }else
                {
                    MessageBox.Show("Actualizado.");
                }
            }catch(Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try
            {
                this.sap.EDITContactosSN("CL03", 2);
                if (this.sap.Error != "")
                {
                    MessageBox.Show("Error: " + this.sap.Error);
                }
                else
                {
                    MessageBox.Show("Actualizado.");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                this.sap.CrearItem();
                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }else
                {
                    MessageBox.Show("Item Agregado");
                }
            }catch(Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            try
            {
                this.sap.EditarItem("ART001");
                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }
                else
                {
                    MessageBox.Show("Item Editado");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            try
            {
                this.sap.agregarDireccionSN("V22000");
                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }
                else
                {
                    MessageBox.Show("Se agrego una dirección");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                this.sap.CrearPedido(out DocEntry);

                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }else
                {
                    MessageBox.Show("Pedido #" + DocEntry + " Creado con exito");
                }

            }catch(Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                this.sap.CrearPedidoDeTipoServicio(out DocEntry);

                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }
                else
                {
                    MessageBox.Show("Pedido #" + DocEntry + " Creado con exito");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            try
            {
                this.sap.agregarLineaPedido(1);

                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }
                else
                {
                    MessageBox.Show("Linea agregada");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                this.sap.CrearEntrega(out DocEntry);

                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }
                else
                {
                    MessageBox.Show("Entrega #" + DocEntry + " Creado con exito");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                this.sap.CrearDevolucion(out DocEntry);

                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }
                else
                {
                    MessageBox.Show("Devolución #" + DocEntry + " Creado con exito");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                this.sap.CrearSalida(out DocEntry);

                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }
                else
                {
                    MessageBox.Show("Salida #" + DocEntry + " Creado con exito");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                this.sap.CrearFacturaConDocumentoBase(out DocEntry);

                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }
                else
                {
                    MessageBox.Show("Factura #" + DocEntry + " Creado con exito");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                this.sap.CrearTransferencia(out DocEntry);

                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }
                else
                {
                    MessageBox.Show("Transferencia #" + DocEntry + " Creado con exito");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                this.sap.CrearPago(out DocEntry,669);

                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }
                else
                {
                    MessageBox.Show("Pago #" + DocEntry + " Creado con exito");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button18_Click(object sender, EventArgs e)
        {
            try
            {
                string Datos = "";
                this.sap.Record(out Datos, 669);

                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }
                else
                {
                    MessageBox.Show(Datos);
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                this.sap.CrearPago(out DocEntry, "647");

                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }
                else
                {
                    MessageBox.Show("Pago #" + DocEntry + " Creado con exito");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            try
            {
                string DocEntry = "";
                this.sap.CrearFacturaConDocumentoBase(out DocEntry,"556");

                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }
                else
                {
                    MessageBox.Show("Factura #" + DocEntry + " Creado con exito");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            try
            {
                //string DocEntry = "";
                this.sap.EjemploTransaction();

                if (this.sap.Error != "")
                {
                    MessageBox.Show(this.sap.Error);
                }
                else
                {
                    MessageBox.Show("exito");
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
    }
}
