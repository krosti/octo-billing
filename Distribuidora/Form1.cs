using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Base;
using Intermedia;
using nmExcel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools.Excel;
using System.Globalization;
using System.Xml;

namespace Distribuidora
{//todo
    public partial class Form1 : Form
    {
        public Form1()
        {

            InitializeComponent();
            ClaseBase = new Basededatos(ref sigla);
            ClaseMedia = new Intermedia2(ref sigla);
            llenarValores();
        }
        public string sigla = "dtb";
        public Basededatos ClaseBase;
        public Intermedia2 ClaseMedia;

        private void llenarValores()//LLeno los combobox con los datos de zonas, sit. iva, rubros, etc.
        {
            /*DataTable clientes = (ClaseMedia.obtener_clientes(sigla).Copy());
            clientes.WriteXml(@"C:\Claudio\Clientes.xml");

            DataSet dataclientes = new DataSet();

            System.IO.FileStream fs = new System.IO.FileStream(@"C:\Claudio\Clientes.xml", System.IO.FileMode.Open);
            dataclientes.ReadXml(fs);
            fs.Close();

            dataGridView4.DataSource = dataclientes;
            dataGridView4.DataMember = "Table";

            DataTable productos = (ClaseMedia.obtener_productos(sigla).Copy());
            productos.WriteXml(@"C:\Claudio\Productos.xml");

            DataSet dataproductos = new DataSet();

            System.IO.FileStream fs2 = new System.IO.FileStream(@"C:\Claudio\Productos.xml", System.IO.FileMode.Open);
            dataproductos.ReadXml(fs2);
            fs2.Close();

            dataGridView3.DataSource = dataproductos;
            dataGridView3.DataMember = "Table";*/

            DataTable valores = ClaseMedia.obtener_rubros(sigla);
            foreach (DataRow row in valores.Rows)
            {
                comboBox3.Items.Add(row[0].ToString()+" - "+row[1].ToString());
                comboBox14.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                comboBox20.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                comboBox21.Items.Add(row[0].ToString() + " - " + row[1].ToString());
            }
            comboBox20.Items.Add("TODOS");
            comboBox20.SelectedIndex = 0;
            valores = ClaseMedia.obtener_subrubros(sigla);
            foreach (DataRow row in valores.Rows)
            {
                comboBox17.Items.Add(row[0].ToString() + " - " + row[1].ToString());
            }
            comboBox17.SelectedIndex = 0;
            valores = ClaseMedia.obtener_sitivas(sigla);
            foreach (DataRow row in valores.Rows)
            {
                comboBox2.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                comboBox8.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                comboBox11.Items.Add(row[0].ToString() + " - " + row[1].ToString());
            }
            valores = ClaseMedia.obtener_zonas(sigla);
            foreach (DataRow row in valores.Rows)
            {
                comboBox1.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                comboBox4.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                comboBox9.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                comboBox10.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                comboBox12.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                comboBox15.Items.Add(row[0].ToString() + " - " + row[1].ToString());
            }
            valores = ClaseMedia.obtener_tasas(sigla);
            foreach (DataRow row in valores.Rows)
            {
                comboBox7.Items.Add(row[0].ToString() + " - " + row[1].ToString());
            }
            comboBox9.Items.Add("LISTAR TODAS");
            comboBox9.SelectedIndex = comboBox9.Items.Count-1;
            comboBox12.Items.Add("LISTAR TODAS");
            comboBox12.SelectedIndex = comboBox9.Items.Count - 1;
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox7.SelectedIndex = 0;
            textBox1.Text = ClaseMedia.obtener_nro_cliente(sigla).PadLeft(5,Convert.ToChar("0"));

            
            //COMPROBANTES
            comboBox6.Items.Add("A");
            comboBox6.Items.Add("B");
            comboBox10.Items.Add("A");
            comboBox10.Items.Add("B");
            comboBox13.Items.Add("A");
            comboBox13.Items.Add("B");
            comboBox13.Items.Add("TODOS");
            comboBox6.SelectedIndex = 0;
            comboBox10.SelectedIndex = 0;
            comboBox13.SelectedIndex = 0;

            comboBox5.Items.Add("Factura");
            comboBox5.Items.Add("Nota de Credito");
            comboBox5.Items.Add("Nota de Debito");
            //comboBox5.Items.Add("Remito"); POR AHORA NO SE USAN
            //comboBox5.Items.Add("Recibo");
            comboBox5.SelectedIndex = 0;
            //
            crear_factura();
            //
            comboBox16.Items.Add("Efectivo");
            comboBox16.Items.Add("Cta. Cte.");
            comboBox16.SelectedIndex = 1;
            //
            comboBox19.Items.Add("Clientes");
            comboBox19.Items.Add("Productos");
            comboBox19.SelectedIndex = 0;
            //
            comboBox22.Items.Add("Rubros");
            comboBox22.Items.Add("Zonas");
            comboBox22.Items.Add("Subrubros");
            comboBox22.SelectedIndex = 0;
        }

        private void button5_Click(object sender, EventArgs e)//ABM un nuevo Cliente
        {
            bool exito = false;
            string[] registro;
            registro = new string[10];
            registro[0] = textBox1.Text;//Codigo
            registro[1] = textBox2.Text;//Razon
            registro[2] = textBox3.Text;//Domicilio
            registro[3] = textBox4.Text;//Localidad
            registro[4] = comboBox1.SelectedIndex.ToString();//Zona
            registro[5] = textBox6.Text;//Telefono
            registro[6] = textBox7.Text;//CUIT
            registro[7] = (Convert.ToInt16(comboBox2.SelectedIndex)+1).ToString();//IVA
            registro[8] = textBox9.Text;//Saldo
            registro[9] = textBox42.Text; //Bonificacion

            if (groupBox1.Text == "Alta")
            {
                if (ClaseMedia.insertar_cliente(sigla, registro) == 0)
                {
                    label8.Text = "Cargado correctamente";
                    textBox1.Text = ClaseMedia.obtener_nro_cliente(sigla).PadLeft(5,Convert.ToChar("0"));//Obtengo el Proximo codigo de cliente a cargar
                    exito = true;
                }
                else
                {
                    label8.Text = "No se pudo cargar";
                }

            }
            else if (groupBox1.Text == "Modificar")
            {
                if (ClaseMedia.cliente_existente(sigla, textBox1.Text) == 0)
                {
                    if (ClaseMedia.modificar_cliente(sigla, registro) == 0)
                    {
                        label8.Text = "Modificado correctamente";
                        exito = true;
                    }
                }
                else
                {
                    label8.Text = "No se pudo modificar";
                }
            }
            else if (groupBox1.Text == "Eliminar")
            {
                if (ClaseMedia.clienteexistente(sigla, textBox1.Text) == 0)
                {
                    if (textBox9.Text == "0")
                    {
                        if (ClaseMedia.borrar_cliente(sigla, textBox1.Text) == 0)
                        {
                            label8.Text = "Eliminado correctamente";
                            exito = true;
                        }
                    }
                    else
                    {
                        label8.Text = "El cliente tiene movimientos pendientes";
                    }
                }
                else
                {
                    label8.Text = "No se pudo eliminar";
                }
            }
            if (exito == true)
            {
                if (groupBox1.Text != "Alta")
                {
                    textBox1.Text = "";
                }
                //limpio los campos, en caso que se cargo
                textBox2.Text = "";
                textBox3.Text = "";
                textBox4.Text = "";
                textBox6.Text = "";
                textBox7.Text = "";
                textBox9.Text = "";
                textBox42.Text = "";
                comboBox1.SelectedIndex = -1;
                comboBox2.SelectedIndex = -1;
            }
        }

        private void button1_Click(object sender, EventArgs e)//Boton para dar de alta un nuevo cliente
        {
            label8.Text = "";
            groupBox1.Text = "Alta";
            textBox1.ReadOnly = true; textBox2.ReadOnly = false;
            textBox3.ReadOnly = false; textBox4.ReadOnly = false;
            textBox6.ReadOnly = false; textBox9.ReadOnly = false;
            textBox7.ReadOnly = false; 
            textBox1.Clear(); textBox2.Clear();
            textBox3.Clear(); textBox4.Clear();
            textBox6.Clear(); comboBox1.SelectedIndex = 0;
            textBox7.Clear(); comboBox2.SelectedIndex = 0;
            textBox9.Clear(); textBox42.Clear(); textBox42.ReadOnly = false;
            textBox1.Text = ClaseMedia.obtener_nro_cliente(sigla).PadLeft(5,Convert.ToChar("0"));
        }

        private void button2_Click(object sender, EventArgs e)//Boton para modificar un nuevo cliente
        {
            label8.Text = "";
            groupBox1.Text = "Modificar";
            textBox1.ReadOnly = false; textBox2.ReadOnly = false;
            textBox3.ReadOnly = false; textBox4.ReadOnly = false;
            textBox6.ReadOnly = false; textBox9.ReadOnly = false;
            textBox7.ReadOnly = false; textBox42.Clear(); textBox42.ReadOnly = false;
            textBox1.Clear(); textBox2.Clear();
            textBox3.Clear(); textBox4.Clear();
            textBox6.Clear(); comboBox1.SelectedIndex = 0;
            textBox7.Clear(); comboBox2.SelectedIndex = 0;
            textBox9.Clear();
        }

        private void button3_Click(object sender, EventArgs e)//Boton para dar de baja un nuevo cliente
        {
            label8.Text = "";
            groupBox1.Text = "Eliminar";
            textBox1.ReadOnly = false; textBox2.ReadOnly = true;
            textBox3.ReadOnly = true; textBox4.ReadOnly = true;
            textBox6.ReadOnly = true; textBox9.ReadOnly = true;
            textBox7.ReadOnly = true; textBox42.Clear(); textBox42.ReadOnly = true;
            textBox1.Clear(); textBox2.Clear();
            textBox3.Clear(); textBox4.Clear();
            textBox6.Clear(); comboBox1.SelectedIndex = 0;
            textBox7.Clear(); comboBox2.SelectedIndex = 0;
            textBox9.Clear();
        }

        private void textBox1_Leave(object sender, EventArgs e)//Relleno los campos con datos del cliente cuando escribe el codigo (leave)
        {
            if ((groupBox1.Text == "Modificar") || (groupBox1.Text == "Eliminar"))
            {
                if (ClaseMedia.clienteexistente(sigla, textBox1.Text) != 0)
                {
                    label8.Text = "Cliente inexistente";
                    textBox2.Text = "";
                    textBox3.Text = "";
                    textBox4.Text = "";
                    comboBox1.SelectedIndex = 0;
                    textBox6.Text = "";
                    textBox7.Text = "";
                    comboBox2.SelectedIndex = 0;
                    textBox9.Text = "";
                }
                else
                {
                    string[] cliente;
                    cliente = new string[9];
                    cliente = ClaseMedia.obtener_cliente(sigla, textBox1.Text);
                    textBox2.Text = cliente[1];
                    textBox3.Text = cliente[2];
                    textBox4.Text = cliente[3];
                    comboBox1.SelectedIndex = Convert.ToInt32(cliente[4]);
                    textBox6.Text = cliente[5];
                    textBox7.Text = cliente[6];
                    comboBox2.SelectedIndex = Convert.ToInt32(cliente[7])-1;
                    textBox9.Text = cliente[8];
                    textBox42.Text = cliente[9];
                    label8.Text = "";
                }
            }
        }
        
        private void button10_Click(object sender, EventArgs e)//Boton de ABM de Productos
        {
            bool exito = false;
            string[] registro;
            registro = new string[11];
            string rubro = comboBox3.SelectedIndex.ToString();
            registro[0] = textBox5.Text;//Codigo
            registro[1] = rubro;//Rubro
            registro[2] = textBox8.Text;//Descripcion
            registro[3] = textBox10.Text;//Precio Minorista
            registro[4] = textBox11.Text;//Precio Mayorista
            registro[5] = textBox12.Text;//Impuesto $
            registro[6] = textBox21.Text;//Impuesto %
            registro[7] = textBox13.Text;//Stock Actual
            registro[8] = textBox14.Text;//Stock Minimo
            registro[9] = (comboBox17.SelectedIndex + 1).ToString();//subrubro 
            registro[10] = (comboBox7.SelectedIndex+1).ToString();

            if (groupBox2.Text == "Alta")
            {
                if (ClaseMedia.insertar_producto(sigla, registro) == 0)
                {
                    label20.Text = "Cargado correctamente";
                    exito = true;
                    textBox5.Text = ClaseMedia.obtener_nro_producto(sigla, comboBox3.SelectedIndex.ToString().PadLeft(4,Convert.ToChar("0")));//Obtengo el proximo codigo a cargar
                }
                else
                {
                    label20.Text = "No se pudo cargar";
                }

            }
            else if (groupBox2.Text == "Modificar")
            {
                if (ClaseMedia.producto_existente(sigla, textBox5.Text) == 0)
                {
                    if (ClaseMedia.modificar_producto(sigla, registro) == 0)
                    {
                        label20.Text = "Modificado correctamente";
                        exito = true;
                    }
                }
                else
                {
                    label20.Text = "No se pudo modificar";
                }
            }
            else if (groupBox2.Text == "Eliminar")
            {
                if (ClaseMedia.producto_existente(sigla, textBox1.Text) == 0)
                {
                    if (ClaseMedia.borrar_producto(sigla, textBox1.Text) == 0)
                    {
                        label20.Text = "Eliminado correctamente";
                        exito = true;
                    }
                }
                else
                {
                    label20.Text = "No se pudo eliminar";
                }
            }
            if (exito == true)
            {
                if (groupBox2.Text != "Alta")
                {
                    textBox5.Text = "";
                }
                textBox8.Text = "";
                textBox10.Text = "";
                textBox11.Text = "";
                textBox12.Text = "";
                textBox13.Text = "";
                textBox14.Text = "";
                textBox21.Text = "";
                comboBox3.SelectedIndex = 0;
                comboBox7.SelectedIndex = 0;
            }
        }

        private void textBox5_Leave(object sender, EventArgs e)//Relleno los campos con datos del producto (si existe)
        {
            if ((groupBox2.Text == "Modificar") || (groupBox2.Text == "Eliminar"))
            {
                if (ClaseMedia.producto_existente(sigla, textBox5.Text) != 0)
                {
                    label20.Text = "Producto inexistente";
                    textBox8.Text = "";
                    textBox10.Text = "";
                    textBox11.Text = "";
                    textBox12.Text = "";
                    textBox13.Text = "";
                    textBox14.Text = "";
                    textBox21.Text = "";
                }
                else
                {
                    string[] producto;
                    producto = new string[13];
                    producto = ClaseMedia.obtener_producto(sigla, textBox5.Text);
                    textBox8.Text = producto[2];
                    textBox10.Text = producto[3];
                    textBox11.Text = producto[4];
                    textBox12.Text = producto[5];
                    textBox21.Text = producto[6];
                    textBox13.Text = producto[7];
                    textBox14.Text = producto[8];
                    comboBox3.SelectedIndex = Convert.ToInt32(producto[1]);
                    comboBox7.SelectedIndex = Convert.ToInt32(producto[10])-1;
                    comboBox17.SelectedIndex = Convert.ToInt32(producto[9]) - 1;
                    label20.Text = "";
                }
            }
        }

        private void button9_Click(object sender, EventArgs e)//Alta de producto
        {
            groupBox2.Text = "Alta";
            textBox5.ReadOnly = false; textBox21.ReadOnly = false;
            textBox8.ReadOnly = false; textBox12.ReadOnly = false;
            textBox10.ReadOnly = false; textBox13.ReadOnly = false;
            textBox11.ReadOnly = false; textBox14.ReadOnly = false;
            textBox8.Clear();
            textBox10.Clear(); textBox11.Clear();
            textBox12.Clear(); textBox13.Clear();
            textBox14.Clear(); textBox21.Clear();
            comboBox3.SelectedIndex = 0;
            textBox5.Text = ClaseMedia.obtener_nro_producto(sigla, comboBox3.SelectedIndex.ToString()).PadLeft(4, Convert.ToChar("0"));
        }

        private void button8_Click(object sender, EventArgs e)//Modif. producto
        {
            groupBox2.Text = "Modificar";
            textBox5.ReadOnly = false; textBox21.ReadOnly = false;
            textBox8.ReadOnly = false; textBox12.ReadOnly = false;
            textBox10.ReadOnly = false; textBox13.ReadOnly = false;
            textBox11.ReadOnly = false; textBox14.ReadOnly = false;
            textBox5.Clear(); textBox8.Clear();
            textBox10.Clear(); textBox11.Clear();
            textBox12.Clear(); textBox13.Clear();
            textBox14.Clear(); textBox21.Clear();
            comboBox1.SelectedIndex = 0;
            //textBox5.Text = comboBox3.SelectedIndex.ToString() + "-";
        }

        private void button7_Click(object sender, EventArgs e)//Baja de producto
        {
            groupBox2.Text = "Eliminar";
            textBox5.ReadOnly = false; textBox21.ReadOnly = true;
            textBox8.ReadOnly = true; textBox12.ReadOnly = true;
            textBox10.ReadOnly = true; textBox13.ReadOnly = true;
            textBox11.ReadOnly = true; textBox14.ReadOnly = true;
            textBox5.Clear(); textBox8.Clear();
            textBox10.Clear(); textBox11.Clear();
            textBox12.Clear(); textBox13.Clear();
            textBox14.Clear(); textBox21.Clear();
            comboBox1.SelectedIndex = 0;
            textBox5.Text = comboBox3.SelectedIndex.ToString() + "-";
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            FormResizer objFormResizer = new FormResizer();
            objFormResizer.ResizeForm(this, 900, 1600);
        }

        private void comboBox3_SelectedIndexChanged(object sender, EventArgs e)//Busco el codigo del nuevo producto a cargar cuando se cambia de rubro
        {
            if (groupBox2.Text == "Alta")
            {
                textBox5.Text = ClaseMedia.obtener_nro_producto(sigla, comboBox3.SelectedIndex.ToString()).PadLeft(4,Convert.ToChar("0"));                
            }
            else if ((groupBox2.Text == "Modificar") || (groupBox2.Text == "Eliminar"))
            {
                //textBox5.Text = comboBox3.SelectedIndex.ToString() + "-";
            }
        }

        private void button66_Click(object sender, EventArgs e)
        {
            groupBox16.Enabled = true;
            textBox48.Focus();
        }


        //
        //FACTURACION
        //

        bool cliente_existe = false;
        //VARIABLES GLOBALES DE LA FACTURA
        DataTable factura = new DataTable();
        double subtotal = 0;
        double impuestos = 0;
        double iva = 0;
        double ivano = 0;
        double total = 0;
        int cant_renglones = 0;

        private void crear_factura()
        {
            factura.Columns.Add("Nro", Type.GetType("System.String"));
            factura.Columns.Add("Cantidad", Type.GetType("System.String"));
            factura.Columns.Add("Codigo", Type.GetType("System.String"));
            factura.Columns.Add("Descripcion", Type.GetType("System.String"));
            factura.Columns.Add("Imp. Unit.", Type.GetType("System.String"));
            factura.Columns.Add("Bonificacion", Type.GetType("System.String"));
            factura.Columns.Add("Bonif. Unit.", Type.GetType("System.String"));
            factura.Columns.Add("Precio s/IVA", Type.GetType("System.String"));
            factura.Columns.Add("Precio Unit. Final", Type.GetType("System.String"));
            factura.Columns.Add("Subtotal", Type.GetType("System.String"));
        }

        private void textBox48_Leave(object sender, EventArgs e)//Relleno los datos del cliente en la factura
        {
            if (textBox48.Text == "8888")//Para no imprimirla
            {
                checkBox2.Checked = true;
                checkBox2.Visible = true;
                textBox48.Focus();
            }
            string[] cliente;
            cliente = new string[9];
            cliente = ClaseMedia.obtener_cliente(sigla, textBox48.Text);
            if (cliente[0] != null)
            {
                textBox47.Text = cliente[1];
                textBox49.Text = cliente[2];
                textBox52.Text = cliente[6];
                textBox50.Text = cliente[3];
                textBox18.Text = cliente[9];
                if (cliente[4] != "999")//Si es consumidor final la zona no la trato de ubicar.
                {
                    comboBox4.SelectedIndex = Convert.ToInt16(cliente[4]);
                }
                else // La factura va a ser en efectivo
                {
                    comboBox16.SelectedIndex = 0;
                }
                comboBox8.SelectedIndex = Convert.ToInt16(cliente[7])-1;
                label23.Text = "";
                cliente_existe = true;
            }
            else
            {
                if (textBox48.Text != "8888")
                {
                    label23.Text = "Cliente inexistente";
                }
                else
                {
                    textBox48.Text = "";
                }
                textBox47.Text = "";
                textBox49.Text = "";
                textBox52.Text = "";
                textBox50.Text = "";
                comboBox4.SelectedIndex = -1;
                comboBox8.SelectedIndex = -1;
                cliente_existe = false;
            }
        }

        private void textBox53_Leave(object sender, EventArgs e)//Al escribir el codigo del producto
        {
            string[] producto;
            producto = new string[13];
            producto = ClaseMedia.obtener_producto(sigla, textBox53.Text);
            if (producto[0] != null)
            {
                label89.Text = producto[2];
                //Miro si el cliente tiene BONIFICACION, le hago el descuento en el precio del producto
                textBox20.Text = (Convert.ToDouble(producto[4])-(Convert.ToDouble(producto[4])*(Convert.ToDouble(textBox18.Text)*0.01))).ToString("#0.00");
                textBox15.Text = producto[6];
                textBox22.Text = producto[5];
                label23.Text = "";
                if (comboBox8.SelectedIndex == 3)//Si es exento le saco el IVA a todos los productos
                {
                    textBox34.Text = ClaseMedia.obtener_valor_tasa(sigla,"3");
                }
                else // Le pongo el IVA que tiene el producto
                {
                    textBox34.Text = ClaseMedia.obtener_valor_tasa(sigla, producto[10]);
                }
            }
            else
            {
                textBox53.Clear();
                label23.Text = "Producto Inexistente";
            }
        }

        private void button60_Click(object sender, EventArgs e)
        {
            if (cliente_existe == true)
            {
                groupBox18.Enabled = true;
                textBox53.Focus();
                if (comboBox8.SelectedIndex == 0)//Si es RI le hago A
                {
                    comboBox6.SelectedIndex = 0;
                }
                else
                {
                    comboBox6.SelectedIndex = 1;
                }
                groupBox16.Enabled = false;
                if (textBox18.Text == "")
                {
                    textBox18.Text = "0";
                }
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e) // Para cargar productos especiales (fletes o cosas asi)
        {
            label23.Text = "";
            if (checkBox1.Checked)
            {
                textBox61.Visible = true;
                textBox16.Visible = true;
                label21.Visible = true;
                textBox61.Focus();
                textBox53.Clear();
                textBox54.Clear();
                textBox20.Clear();
                textBox15.Clear();
                textBox22.Clear();
                label89.Text = "";
                textBox19.Text = "0";
                textBox19.Enabled = false;
            }
            else
            {
                textBox61.Visible = false;
                textBox16.Visible = false;
                label21.Visible = false;
                textBox53.Focus();
                textBox19.Enabled = true;

            }
        }

        private void button64_Click(object sender, EventArgs e) //Carga de nuevo renglon
        {
            if (cant_renglones < 12)
            {
                if (((textBox53.Text != "") && (textBox54.Text != "") && (textBox19.Text != "")) || ((checkBox1.Checked) && ((textBox61.Text != "") && (textBox16.Text != ""))))
                {
                    DataRow renglon;
                    renglon = factura.NewRow();
                    renglon[0] = factura.Rows.Count + 1;
                    if (textBox53.Text != "")//Si no es un producto especial
                    {
                        renglon[1] = textBox54.Text;
                        renglon[2] = textBox53.Text;
                        renglon[3] = label89.Text;//Descripcion
                        //Le agrego el % y $ de impuestos que tenga el producto
                        renglon[4] = ((Convert.ToDouble(textBox20.Text) * (Convert.ToDouble(textBox15.Text) * 0.01)) + (Convert.ToDouble(textBox22.Text))).ToString();
                        renglon[5] = (Convert.ToDouble(textBox19.Text) * 0.01).ToString("#0.00");//bonificacion UNITARIA
                        //Hago el calculo del precio unitario con impuestos y bonificacion
                        renglon[6] = ((Convert.ToDouble(textBox20.Text) + Convert.ToDouble(renglon[4])) * Convert.ToDouble(renglon[5])).ToString("#0.00");
                        renglon[7] = ((Convert.ToDouble(textBox20.Text) + Convert.ToDouble(renglon[4])) - Convert.ToDouble(renglon[6])).ToString("#0.00");
                        if (textBox34.Text != "0")
                        {
                            renglon[8] = ((Convert.ToDouble(renglon[7])) + (Convert.ToDouble(renglon[7]) * Convert.ToDouble(textBox34.Text) / 100)).ToString("#0.00");
                        }
                        else //Si es exento no tiene IVA
                        {
                            renglon[8] = renglon[7];
                        }
                        if (checkBox2.Visible == true)//Si voy a facturar en negro le saco los impuestos y el IVA, ya que no lleva nada de eso.
                        {
                            renglon[4] = 0.ToString();//Impuestos
                            renglon[6] = textBox20.Text;
                            renglon[7] = textBox20.Text;
                            renglon[8] = textBox20.Text;
                        }

                    }
                    else //Cargo el producto especial
                    {
                        renglon[0] = "1";
                        renglon[1] = "1";
                        renglon[2] = "-";
                        renglon[3] = textBox61.Text;
                        renglon[4] = "0";
                        renglon[5] = "0";
                        renglon[6] = "0";
                        renglon[7] = textBox16.Text;
                        renglon[8] = textBox16.Text;
                    }
                    if (comboBox6.SelectedIndex == 0) //Si es comprobante A el Subtotal va sin IVA
                    {
                        renglon[9] = (Convert.ToDouble(renglon[7]) * Convert.ToDouble(renglon[1])).ToString("#0.00");
                    }
                    else //Sino paso el precio con IVA Indiscriminado
                    {
                        renglon[9] = (Convert.ToDouble(renglon[8]) * Convert.ToDouble(renglon[1])).ToString("#0.00");
                    }

                    factura.Rows.Add(renglon);
                    dataGridView2.DataSource = factura;

                    if (comboBox6.SelectedIndex == 1)//Si es tipo B
                    {
                        if (comboBox8.SelectedIndex == 4)//Si es RNI le calculo el IVA de RNI
                        {
                            subtotal += Convert.ToDouble(renglon[9]);//Producto con IVA
                            ivano = (subtotal * 0.105);
                        }
                        else if (comboBox8.SelectedIndex == 5)//Si es NC le calculo el IVA de NC (se agrega un 10.5 al sub+IVA)
                        {
                            subtotal += Convert.ToDouble(renglon[9]);//Producto con IVA
                            ivano = (subtotal + iva) * 0.105;
                        }
                        else
                        {
                            subtotal += Convert.ToDouble(renglon[9]);//Producto con IVA
                            iva += 0;
                        }
                    }
                    else// Si es tipo A
                    {
                        subtotal += Convert.ToDouble(renglon[7]) * Convert.ToDouble(renglon[1]);//Precio del producto SIN IVA por la cantidad
                        iva += (Convert.ToDouble(renglon[8]) - Convert.ToDouble(renglon[7])) * Convert.ToDouble(renglon[1]);
                        //Voy acumulando el iva aparte, en B siempre se mantiene en 0
                    }

                    total = subtotal + impuestos + iva + ivano;
                    textBox55.Text = subtotal.ToString("#0.00");
                    textBox56.Text = impuestos.ToString("#0.00");
                    textBox57.Text = iva.ToString("#0.00");
                    textBox58.Text = ivano.ToString("#0.00");
                    textBox59.Text = total.ToString("#0.00");
                    cant_renglones++;
                    //Falta limitarlo a 14 renglones MAXIMO (es un IF)
                    textBox53.Focus();
                    label23.Text = "";
                }
                else
                {
                    label23.Text = "Indique el producto";
                }
            }
            else
            {
                MessageBox.Show("La cantidad maxima de renglones fue alcanzada");
            }
        }

        private void button61_Click(object sender, EventArgs e)//Carga e impreison de la Factura
        {
            nmExcel._Worksheet workSheet2;
            DateTime m_FechaHora = DateTime.Now;
            Microsoft.Office.Interop.Excel.Application Aplic = new Microsoft.Office.Interop.Excel.Application();
            Aplic.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook Libro = Aplic.Workbooks.Open("C:/Claudio/Fact.xlsx", Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet2 = (nmExcel.Worksheet)Aplic.ActiveSheet;

            //cargar datos factura A
            //Tipo de comprobante
            workSheet2.PageSetup.LeftMargin = 0;
            workSheet2.PageSetup.TopMargin = 0;
            workSheet2.Cells[4, 9] = dateTimePicker4.Text;
            if (comboBox6.SelectedIndex == 0)//Si es comprobante A
            {
                workSheet2.Cells[5, 9] = "Nro:"+ClaseMedia.obtener_nro_facta(sigla);
            }
            else
            {
                workSheet2.Cells[5, 9] = "Nro:" + ClaseMedia.obtener_nro_factb(sigla);
            }
            if (comboBox5.Text != "Factura")
            {
                workSheet2.Cells[5, 10] = comboBox5.Text; 
            }
            //Nombre
            workSheet2.Cells[7, 2] = "Señor(es): " + textBox47.Text;
            //Domicilio
            workSheet2.Cells[8, 2] = "Domicilio: " + textBox49.Text;
            //Sit. IVA:
            workSheet2.Cells[9, 2] = "IVA: " + comboBox8.Text;
            //Venta:
            workSheet2.Cells[10, 2] = "Condicion de Venta: " + comboBox16.Text;
            //Localidad
            workSheet2.Cells[7, 7] = "Localidad: "+textBox50.Text;
            //CUIT
            workSheet2.Cells[8, 7] = "CUIT: " + textBox52.Text;
            //ZONA
            workSheet2.Cells[9, 7] = "Zona: " + comboBox4.Text;
                
            //Cabecera de renglones
            workSheet2.Cells[12, 3] = "Cant";
            workSheet2.Cells[12, 4] = "Código";
            workSheet2.Cells[12, 6] = "Detalle";
            if (comboBox6.SelectedIndex == 0) //Si es A muestro el precio con IVA
            {
                workSheet2.Cells[12, 8] = "Prec. Unitario";
            }
            workSheet2.Cells[12, 9] = "Uni. Final";
            workSheet2.Cells[12, 10] = "Total";
                
            //Renglones
            for (int i = 0; i < cant_renglones; i++)
            {
                workSheet2.Cells[i + 13, 3] = dataGridView2[1, i].Value.ToString();
                workSheet2.Cells[i + 13, 4] = dataGridView2[2, i].Value.ToString();
                workSheet2.Cells[i + 13, 5] = dataGridView2[3, i].Value.ToString();

                if (comboBox6.SelectedIndex == 0) //Si es A muestro el precio con IVA
                {
                    workSheet2.Cells[i + 13, 8] = String.Format("{0:0.00}", dataGridView2[7, i].Value);
                    workSheet2.Cells[i + 13, 9] = String.Format("{0:0.00}", dataGridView2[8, i].Value);
                }
                else
                {
                    workSheet2.Cells[i + 13, 9] = String.Format("{0:0.00}", dataGridView2[8, i].Value);
                }
                workSheet2.Cells[i + 13, 10] = String.Format("{0:0.00}", dataGridView2[9, i].Value);
            }
            if (comboBox6.SelectedIndex == 0)
            {
                //SubTotal
                workSheet2.Cells[26, 3] = "Subtotal";
                workSheet2.Cells[27, 3] = String.Format("{0:0.00}", subtotal);
                //IVA
                workSheet2.Cells[26, 6] = "IVA";
                workSheet2.Cells[27, 6] = String.Format("{0:0.00}", iva);
                //Total (cabecera solo en A)
                workSheet2.Cells[26, 10] = "Total";
            }
            if ((comboBox8.SelectedIndex == 4)||(comboBox8.SelectedIndex == 5))
            {
                workSheet2.Cells[27, 5] = "Percepciones:";
                workSheet2.Cells[27, 7] = String.Format("{0:0.00}", ivano);
            }
            //TOTAL
            workSheet2.Cells[27, 10] = String.Format("{0:0.00}", total);
            //myPrinters.SetDefaultPrinter(cbimpresoras.Text);
            workSheet2.PrintOut(1, 1, 1, false);
            if (Aplic != null)
            {
                Aplic.Quit();
                System.Diagnostics.Process[] pProcess;
                pProcess = System.Diagnostics.Process.GetProcessesByName("Excel");

                foreach (System.Diagnostics.Process oPro in pProcess)
                {
                    if (oPro.StartTime >= m_FechaHora)
                        oPro.Kill();
                }
            }
            
            //FIN DE IMPRESION DE LA FACTURA
            if (MessageBox.Show("Confirmar Impresion", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
            //Si el usuario confirma que la factura se imprimio bien, la guardo en la Base de Datos
            {
                string[] fact;
                fact = new string[20];
                if (comboBox6.SelectedIndex == 0)//Si es comprobante A
                {
                    fact[0] = ClaseMedia.obtener_nro_facta(sigla);
                    fact[1] = "A";
                }
                else//Si es comprobante B
                {
                    fact[0] = ClaseMedia.obtener_nro_factb(sigla);
                    fact[1] = "B";
                }
                if (checkBox2.Visible == false)
                {
                    fact[17] = "1";
                }
                else
                {
                    fact[17] = "0";
                }
                //Cargo datos de la factura
                fact[2] = comboBox5.Text;
                fact[3] = textBox48.Text;
                fact[4] = textBox47.Text;
                fact[5] = textBox52.Text;
                fact[6] = (comboBox8.SelectedIndex + 1).ToString();
                fact[7] = textBox49.Text;
                fact[8] = textBox50.Text;
                fact[9] = comboBox4.SelectedIndex.ToString();
                fact[10] = textBox55.Text;
                fact[11] = textBox56.Text;
                fact[12] = textBox57.Text;
                fact[13] = textBox58.Text;
                fact[14] = textBox59.Text;
                fact[15] = dateTimePicker4.Text;
                fact[16] = textBox18.Text;
                if (comboBox16.SelectedIndex == 0)
                {
                    fact[18] = textBox59.Text;
                }
                else
                {
                    fact[18] = 0.ToString();
                }
                fact[19] = "0";
                ClaseMedia.cargar_comprobante(sigla, fact);
                string[] renglon;
                renglon = new string[11];
                foreach (DataRow r in factura.Rows)//Foreach para los renglones de la factura
                {
                    renglon[0] = r[0].ToString();
                    renglon[1] = fact[0];
                    renglon[2] = r[2].ToString();
                    renglon[3] = r[3].ToString();
                    renglon[4] = r[1].ToString();
                    renglon[5] = r[4].ToString();
                    renglon[6] = textBox34.Text;
                    renglon[7] = r[5].ToString();
                    renglon[8] = r[7].ToString();
                    renglon[9] = r[8].ToString();
                    renglon[10] = dateTimePicker4.Text;
                    ClaseMedia.cargar_renglon(sigla, renglon);
                }
                //Actualizacion de saldos, es + o - segun el comprobante
                if (comboBox16.SelectedIndex == 1) //Si fue pago a Cta. Cte.
                {

                    if (comboBox5.Text == "Nota de Credito")
                    {
                        ClaseMedia.actualizar_saldo(sigla, fact[3], fact[14],0);
                    }
                    else if (comboBox5.Text == "Nota de Debito")
                    {
                        ClaseMedia.actualizar_saldo(sigla, fact[3], "-" + fact[14],1);
                    }
                    else if (comboBox5.Text == "Factura")
                    {
                        ClaseMedia.actualizar_stock(sigla, factura);
                        ClaseMedia.actualizar_saldo(sigla, fact[3], "-" + fact[14],1);
                    }
                }
                else //Si fue pago en Efectivo, actualizo stock y guardo el Pago.
                {
                    if (comboBox5.Text == "Factura")
                    {
                        ClaseMedia.actualizar_stock(sigla, factura);
                        string[] cobro;
                        cobro = new string[7];
                        cobro[0] = "''";
                        cobro[1] = textBox48.Text;
                        cobro[2] = total.ToString();
                        cobro[3] = 0.ToString();
                        cobro[4] = 0.ToString();
                        cobro[5] = "Pago Efectivo "+comboBox5.Text+" nro: "+fact[0];
                        cobro[6] = dateTimePicker4.Text;
                        ClaseMedia.cargar_cobro(sigla, cobro);
                    }
                }
                //Falta hacer el remito y recibo. Aunque no deberian tocar nada

                factura.Rows.Clear();
                dataGridView2.DataSource = null;
                groupBox18.Enabled = false;
                groupBox17.Enabled = true;
                dateTimePicker4.Focus();
                label23.Text = "Factura Guardada";
                textBox34.Clear();
                textBox19.Text = "0";
                subtotal = 0;
                impuestos = 0;
                iva = 0;
                ivano = 0;
                total = 0;
                cant_renglones = 0;
            }
        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            checkBox2.Checked = true;
            checkBox2.Visible = false;
            label23.Text = "";
            textBox48.Focus();
        }

        private void textBox48_Enter(object sender, EventArgs e)
        {
            label23.Text = "";
        }

        private void button63_Click(object sender, EventArgs e)
        {
            if (cant_renglones > 0)
            {
                int fila = Convert.ToInt16(dataGridView2.CurrentCell.RowIndex);
                DataRow renglon = factura.Rows[fila];

                if (comboBox6.SelectedIndex == 1)//Si es tipo B
                {
                    if (comboBox8.SelectedIndex == 4)//Si es RNI le calculo el IVA de RNI
                    {
                        subtotal -= Convert.ToDouble(renglon[7]) * Convert.ToDouble(renglon[1]);//Precio del producto SIN IVA por la cantidad
                        iva -= (Convert.ToDouble(renglon[8]) - Convert.ToDouble(renglon[7])) * Convert.ToDouble(renglon[1]);
                        ivano = (subtotal * 0.105);
                    }
                    else if (comboBox8.SelectedIndex == 5)//Si es NC le calculo el IVA de NC (se agrega un 10.5 al sub+IVA)
                    {
                        subtotal -= Convert.ToDouble(renglon[7]) * Convert.ToDouble(renglon[1]);//Precio del producto SIN IVA por la cantidad
                        iva -= (Convert.ToDouble(renglon[8]) - Convert.ToDouble(renglon[7])) * Convert.ToDouble(renglon[1]);
                        ivano = (subtotal + iva) * 0.105;
                    }
                    else
                    {
                        subtotal -= Convert.ToDouble(renglon[9]);//Producto con IVA
                        iva -= 0;
                    }
                }
                else// Si es tipo A
                {
                    subtotal -= Convert.ToDouble(renglon[7]) * Convert.ToDouble(renglon[1]);//Precio del producto SIN IVA por la cantidad
                    iva -= (Convert.ToDouble(renglon[8]) - Convert.ToDouble(renglon[7])) * Convert.ToDouble(renglon[1]);
                    //Voy acumulando el iva aparte, en B siempre se mantiene en 0
                }

                total = subtotal + impuestos + iva + ivano;

                factura.Rows.Remove(renglon);

                textBox55.Text = String.Format("{0:0.00}", subtotal);
                textBox56.Text = String.Format("{0:0.00}", impuestos);
                textBox57.Text = String.Format("{0:0.00}", iva);
                textBox58.Text = String.Format("{0:0.00}", ivano);
                textBox59.Text = String.Format("{0:0.00}", total);
                cant_renglones --;
            }
        }

        //
        //FIN FACTURACION
        //

        //
        //LISTADO DESDE HASTA CON ULTIMOS MOVIMIENTOS DE LOS CLIENTES
        //
        private void button4_Click(object sender, EventArgs e)
        {
            bool todas;
            if (comboBox9.SelectedIndex == comboBox9.Items.Count - 1)
            {
                todas = true;
            }
            else
            {
                todas = false;
            }
            DataTable clientes = ClaseMedia.obtener_clientes_listado_arepartir(sigla, dateTimePicker1.Text, dateTimePicker2.Text, comboBox9.SelectedIndex.ToString(), todas);
            nmExcel._Worksheet workSheet2;
            DateTime m_FechaHora = DateTime.Now;
            Microsoft.Office.Interop.Excel.Application Aplic = new Microsoft.Office.Interop.Excel.Application();
            Aplic.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook Libro = Aplic.Workbooks.Open("C:/Claudio/ListadoMovimientos.xlsx", Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet2 = (nmExcel.Worksheet)Aplic.ActiveSheet;
            int i = 4;
            workSheet2.Cells[1, 2] = dateTimePicker1.Text;
            workSheet2.Cells[1, 5] = dateTimePicker2.Text;
            workSheet2.Cells[1, 9] = comboBox9.Text;
            DateTime fecha;
            foreach (DataRow row in clientes.Rows)
            {
                workSheet2.Cells[i, 1] = row[0].ToString() + " - " + row[1].ToString();
                workSheet2.Cells[i, 4] = "Saldo Actual:";
                workSheet2.Cells[i, 5] = row[2].ToString();
                i++;
                DataTable facturas = ClaseMedia.obtener_comprobantes_cliente_adeuda(sigla, row[0].ToString());
                workSheet2.Cells[i, 1] = "Comprobantes Impagos:";
                foreach (DataRow row2 in facturas.Rows)
                {
                    fecha = Convert.ToDateTime(ClaseMedia.girar_mes_dia(row2[0].ToString()));
                    if ((fecha.Date < dateTimePicker1.Value.Date) || (fecha.Date > dateTimePicker2.Value.Date))
                    {
                        workSheet2.Cells[i, 4] = row2[0].ToString();
                        workSheet2.Cells[i, 5] = row2[1].ToString();
                        workSheet2.Cells[i, 6] = row2[2].ToString();
                        workSheet2.Cells[i, 8] = row2[5].ToString();
                        workSheet2.Cells[i, 9] = row2[7].ToString();
                        i++;
                    }
                }
                facturas = ClaseMedia.obtener_comprobantes_arepartir(sigla, dateTimePicker1.Text, dateTimePicker2.Text, row[0].ToString());
                workSheet2.Cells[i, 1] = "Movimientos en el Período";
                foreach (DataRow row2 in facturas.Rows)
                {
                    workSheet2.Cells[i, 4] = row2[6].ToString();
                    workSheet2.Cells[i, 5] = row2[0].ToString();
                    workSheet2.Cells[i, 6] = row2[2].ToString();
                    workSheet2.Cells[i, 8] = row2[5].ToString();
                    workSheet2.Cells[i, 9] = row2[7].ToString();
                    i++;
                }
                i++;
            }
            workSheet2.PrintOut(1,1,1,false);
            if (Aplic != null)
            {
                Aplic.Quit();
                System.Diagnostics.Process[] pProcess;
                pProcess = System.Diagnostics.Process.GetProcessesByName("Excel");

                foreach (System.Diagnostics.Process oPro in pProcess)
                {
                    if (oPro.StartTime >= m_FechaHora)
                        oPro.Kill();
                }
            }
        }


        //
        //SECCION DE COBROS
        //
        private void textBox26_Leave(object sender, EventArgs e)//Relleno los datos del cliente
        {
            string[] cliente;
            cliente = new string[9];
            cliente = ClaseMedia.obtener_cliente(sigla, textBox26.Text);
            if (cliente[0] != null)
            {
                textBox27.Text = cliente[1];
                textBox25.Text = cliente[8];
                textBox23.Text = cliente[6];
                textBox24.Text = cliente[3];
                if (cliente[4] != "999")//Si es consumidor final la zona no la trato de ubicar.
                {
                    comboBox4.SelectedIndex = Convert.ToInt16(cliente[4]);
                }
                comboBox11.SelectedIndex = Convert.ToInt16(cliente[7]) - 1;
                label47.Text = "";
                dataGridView5.DataSource = ClaseMedia.obtener_comprobantes_cliente_adeuda(sigla, textBox26.Text);
            }
            else
            {
                label47.Text = "Cliente inexistente";
                textBox23.Text = "";
                textBox24.Text = "";
                textBox27.Text = "";
                textBox25.Text = "";
                comboBox10.SelectedIndex = 0;
                comboBox11.SelectedIndex = 0;
            }
        }

        private void radioButton5_CheckedChanged(object sender, EventArgs e)
        {
            if (groupBox6.Enabled == true)
            {
                groupBox6.Enabled = false;
            }
            else
            {
                groupBox6.Enabled = true;
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if ((dataGridView5.Rows.Count > 0)&&(dataGridView5.CurrentRow.Selected == true)&&(textBox17.Text != ""))
            {
                string[] cobro;
                cobro = new string[8];
                cobro[0] = "''";
                cobro[1] = textBox26.Text;
                cobro[2] = textBox17.Text;
                if (radioButton4.Checked)
                {
                    cobro[3] = 0.ToString();
                }
                else
                {
                    cobro[3] = 1.ToString();
                }
                cobro[4] = textBox29.Text;
                cobro[5] = textBox30.Text;
                cobro[6] = dateTimePicker6.Text;
                cobro[7] = dataGridView5.CurrentRow.Cells[1].Value.ToString();

                string nro = dataGridView5.CurrentRow.Cells[1].Value.ToString(); //Obtengo el nro de factura
                string tipo = dataGridView5.CurrentRow.Cells[2].Value.ToString(); //Obtengo el tipo de factura
                string deuda = dataGridView5.CurrentRow.Cells[7].Value.ToString();//Obtengo el valor de lo que falta pagar de la Factura que va a pagar 
                string total = dataGridView5.CurrentRow.Cells[5].Value.ToString();//Obtengo el valor del total de la factura 

                if ((tipo == "Factura") || (tipo == "Nota de Debito"))
                {
                    if ((Convert.ToDouble(deuda) >= 0) && (Convert.ToDouble(textBox17.Text) + Convert.ToDouble(deuda) <= Convert.ToDouble(total)) && (Convert.ToDouble(textBox17.Text) > 0))
                    {
                        string[] comp;
                        comp = new string[2];
                        comp[0] = nro;
                        comp[1] = (Convert.ToDouble(deuda) + Convert.ToDouble(textBox17.Text)).ToString();//Actualizo el campo Pagado de la factura y guardo

                        bool chequebien = true;
                        if (radioButton5.Checked)//Si va a cobrar con un cheque lo guardo
                        {
                            string[] cheque;
                            cheque = new string[8];
                            cheque[0] = textBox28.Text;
                            cheque[1] = textBox29.Text;
                            cheque[2] = dateTimePicker5.Text;
                            cheque[3] = dateTimePicker3.Text;
                            cheque[4] = textBox31.Text;
                            cheque[5] = textBox17.Text;
                            cheque[6] = "0";
                            cheque[7] = "0";
                            if (ClaseMedia.cargar_cheque(sigla, cheque) != 0)
                            {
                                chequebien = false;
                            }
                        }
                        if (chequebien == true)
                        {
                            ClaseMedia.actualizar_saldo_factura(sigla, comp);

                            if (ClaseMedia.cargar_cobro(sigla, cobro) == 0)
                            {
                                ClaseMedia.actualizar_saldo(sigla, textBox26.Text, textBox17.Text, 1);
                                label47.Text = "Cargado";
                                dataGridView5.DataSource = ClaseMedia.obtener_comprobantes_cliente_adeuda(sigla, textBox26.Text);
                                //limpio datos del cliente
                                textBox23.Text = "";
                                textBox24.Text = "";
                                textBox27.Text = "";
                                textBox25.Text = "";
                                comboBox10.SelectedIndex = 0;
                                comboBox11.SelectedIndex = 0;
                                //limpio los otros datos
                                textBox17.Text = "";
                                textBox30.Text = "";
                                textBox28.Text = "";
                                textBox29.Text = "";
                                radioButton4.Checked = true;

                            }
                            else
                            {
                                label47.Text = "No se pudo cargar";
                            }
                        }
                        else
                        {
                            label47.Text = "Error: El cheque y el pago no se pudieron cargar. Verifique que el cheque no se haya utilizado para cobrar";
                        }
                    }
                    else
                    {
                        label47.Text = "Error: El comprobante ya fue pagado o el monto a pagar debe ser mayor a cero.";
                    }
                }
                else
                {
                    label47.Text = "Error: El comprobante debe ser Factura o Nota de Debito";
                }
            }
            else
            {
                label47.Text = "Error: Seleccione un comprobante a cobrar e indique el monto";
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            bool todas;
            if (comboBox9.SelectedIndex == comboBox9.Items.Count - 1)
            {
                todas = true;
            }
            else
            {
                todas = false;
            }
            DataTable clientes = ClaseMedia.obtener_clientes_listado_arepartir(sigla, dateTimePicker1.Text, dateTimePicker2.Text, comboBox9.SelectedIndex.ToString(), todas);
            nmExcel._Worksheet workSheet2;
            DateTime m_FechaHora = DateTime.Now;
            Microsoft.Office.Interop.Excel.Application Aplic = new Microsoft.Office.Interop.Excel.Application();
            Aplic.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook Libro = Aplic.Workbooks.Open("C:/Claudio/ListadoMovimientos.xlsx", Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet2 = (nmExcel.Worksheet)Aplic.ActiveSheet;
            int i = 4;
            workSheet2.Cells[1, 2] = dateTimePicker1.Text;
            workSheet2.Cells[1, 5] = dateTimePicker2.Text;
            workSheet2.Cells[1, 9] = comboBox9.Text;
            DateTime fecha;
            foreach (DataRow row in clientes.Rows)
            {
                workSheet2.Cells[i, 1] = row[0].ToString() + " - " + row[1].ToString();
                workSheet2.Cells[i, 4] = "Saldo Actual:";
                workSheet2.Cells[i, 5] = row[2].ToString();
                i++;
                DataTable facturas = ClaseMedia.obtener_comprobantes_cliente_adeuda(sigla, row[0].ToString());
                workSheet2.Cells[i, 1] = "Comprobantes Impagos:";
                foreach (DataRow row2 in facturas.Rows)
                {
                    fecha = Convert.ToDateTime(ClaseMedia.girar_mes_dia(row2[0].ToString()));
                    if ((fecha.Date < dateTimePicker1.Value.Date) || (fecha.Date > dateTimePicker2.Value.Date))
                    {
                        workSheet2.Cells[i, 4] = row2[0].ToString();
                        workSheet2.Cells[i, 5] = row2[1].ToString();
                        workSheet2.Cells[i, 6] = row2[2].ToString();
                        workSheet2.Cells[i, 8] = row2[5].ToString();
                        workSheet2.Cells[i, 9] = row2[7].ToString();
                        i++;
                    }
                }
                facturas = ClaseMedia.obtener_comprobantes_arepartir(sigla, dateTimePicker1.Text, dateTimePicker2.Text, row[0].ToString());
                workSheet2.Cells[i, 1] = "Movimientos en el Período";
                foreach (DataRow row2 in facturas.Rows)
                {
                    workSheet2.Cells[i, 4] = row2[6].ToString();
                    workSheet2.Cells[i, 5] = row2[0].ToString();
                    workSheet2.Cells[i, 6] = row2[2].ToString();
                    workSheet2.Cells[i, 8] = row2[5].ToString();
                    workSheet2.Cells[i, 9] = row2[7].ToString();
                    i++;
                }
                i++;
            }
            Aplic.Visible = true;
            workSheet2.PrintPreview();
            Aplic.Visible = false;
            if (Aplic != null)
            {
                Aplic.Quit();
                System.Diagnostics.Process[] pProcess;
                pProcess = System.Diagnostics.Process.GetProcessesByName("Excel");

                foreach (System.Diagnostics.Process oPro in pProcess)
                {
                    if (oPro.StartTime >= m_FechaHora)
                        oPro.Kill();
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            bool todas;
            if (comboBox12.SelectedIndex == comboBox12.Items.Count - 1)
            {
                todas = true;
            }
            else
            {
                todas = false;
            }
            bool impresas = false;
            if (checkBox3.Checked)
            {
                impresas = true;
            }
            DataTable facturas;
            if (checkBox4.Checked)
            {
                facturas = ClaseMedia.obtener_comprobantes(sigla, dateTimePicker7.Text, dateTimePicker8.Text, comboBox13.Text, impresas, comboBox12.SelectedIndex.ToString(), todas, true);
            }
            else
            {
                facturas = ClaseMedia.obtener_comprobantes(sigla, dateTimePicker7.Text, dateTimePicker8.Text, comboBox13.Text, impresas, comboBox12.SelectedIndex.ToString(), todas, false);
            }
            nmExcel._Worksheet workSheet2;
            DateTime m_FechaHora = DateTime.Now;
            Microsoft.Office.Interop.Excel.Application Aplic = new Microsoft.Office.Interop.Excel.Application();
            Aplic.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook Libro = Aplic.Workbooks.Open("C:/Claudio/ListadoFacturas.xlsx", Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet2 = (nmExcel.Worksheet)Aplic.ActiveSheet;
            int i = 5;
            workSheet2.Cells[1, 2] = dateTimePicker7.Text;
            workSheet2.Cells[1, 5] = dateTimePicker8.Text;
            workSheet2.Cells[1, 9] = comboBox12.Text;
            workSheet2.Cells[2, 4] = comboBox13.Text;
            double subtotal = 0;
            double iva = 0;
            double total = 0;
            if (impresas == true)
            {
                workSheet2.Cells[2, 7] = "Si";
            }
            else
            {
                workSheet2.Cells[2, 7] = "No";
            }
            foreach (DataRow row in facturas.Rows)
            {
                workSheet2.Cells[i, 1] = row[8].ToString();
                workSheet2.Cells[i, 2] = row[1].ToString();
                workSheet2.Cells[i, 3] = row[0].ToString();
                workSheet2.Cells[i, 4] = row[2].ToString();
                workSheet2.Cells[i, 5] = row[3].ToString() + " - " + row[4].ToString();
                workSheet2.Cells[i, 7] = row[5].ToString();
                workSheet2.Cells[i, 8] = row[6].ToString();
                workSheet2.Cells[i, 9] = row[7].ToString();
                if (row[2].ToString() != "Nota de Credito")
                {
                    subtotal += Convert.ToDouble(row[5].ToString());
                    iva += Convert.ToDouble(row[6].ToString());
                    total += Convert.ToDouble(row[7].ToString());
                }
                else
                {
                    subtotal -= Convert.ToDouble(row[5].ToString());
                    iva -= Convert.ToDouble(row[6].ToString());
                    total -= Convert.ToDouble(row[7].ToString());
                }
                if (row[9].ToString() == "0")
                {
                    workSheet2.Cells[i, 10] = "N";
                }
                else
                {
                    workSheet2.Cells[i, 10] = "S";
                }
                i++;
            }
            i++;
            workSheet2.Cells[i, 6] = "TOTALES:";
            workSheet2.Cells[i, 7] = subtotal.ToString();
            workSheet2.Cells[i, 8] = iva.ToString();
            workSheet2.Cells[i, 9] = total.ToString();
            workSheet2.PrintOut(1, 1, 1, false);
            if (Aplic != null)
            {
                Aplic.Quit();
                System.Diagnostics.Process[] pProcess;
                pProcess = System.Diagnostics.Process.GetProcessesByName("Excel");

                foreach (System.Diagnostics.Process oPro in pProcess)
                {
                    if (oPro.StartTime >= m_FechaHora)
                        oPro.Kill();
                }
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            bool todas;
            if (comboBox12.SelectedIndex == comboBox12.Items.Count - 1)
            {
                todas = true;
            }
            else
            {
                todas = false;
            }
            bool impresas = false;
            if (checkBox3.Checked)
            {
                impresas = true;
            }
            DataTable facturas;
            if (checkBox4.Checked)
            {
                facturas = ClaseMedia.obtener_comprobantes(sigla, dateTimePicker7.Text, dateTimePicker8.Text, comboBox13.Text, impresas, comboBox12.SelectedIndex.ToString(), todas,true);
            }
            else
            {
                facturas = ClaseMedia.obtener_comprobantes(sigla, dateTimePicker7.Text, dateTimePicker8.Text, comboBox13.Text, impresas, comboBox12.SelectedIndex.ToString(), todas,false);
            }
            nmExcel._Worksheet workSheet2;
            DateTime m_FechaHora = DateTime.Now;
            Microsoft.Office.Interop.Excel.Application Aplic = new Microsoft.Office.Interop.Excel.Application();
            Aplic.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook Libro = Aplic.Workbooks.Open("C:/Claudio/ListadoFacturas.xlsx", Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet2 = (nmExcel.Worksheet)Aplic.ActiveSheet;
            int i = 5;
            workSheet2.Cells[1, 2] = dateTimePicker7.Text;
            workSheet2.Cells[1, 5] = dateTimePicker8.Text;
            workSheet2.Cells[1, 9] = comboBox12.Text;
            workSheet2.Cells[2, 4] = comboBox13.Text;
            double subtotal = 0;
            double iva = 0;
            double total = 0;
            if (impresas == true) //Si muestro facturas no impresas lo aclaro
            {
                workSheet2.Cells[2, 7] = "Si";
            }
            else
            {
                workSheet2.Cells[2, 7] = "No";
            }
            if (checkBox4.Checked) //Si solo muestro no impresas lo aclaro
            {
                workSheet2.Cells[2, 7] = "Solo no impresas";
            }
            foreach (DataRow row in facturas.Rows)
            {
                workSheet2.Cells[i, 1] = row[8].ToString();
                workSheet2.Cells[i, 2] = row[1].ToString();
                workSheet2.Cells[i, 3] = row[0].ToString();
                workSheet2.Cells[i, 4] = row[2].ToString();
                workSheet2.Cells[i, 5] = row[3].ToString() + " - " + row[4].ToString();
                workSheet2.Cells[i, 7] = row[5].ToString();
                workSheet2.Cells[i, 8] = row[6].ToString();
                workSheet2.Cells[i, 9] = row[10].ToString();
                workSheet2.Cells[i, 10] = row[7].ToString();
                if (row[2].ToString() != "Nota de Credito")
                {
                    subtotal += Convert.ToDouble(row[5].ToString());
                    iva += Convert.ToDouble(row[6].ToString());
                    ivano += Convert.ToDouble(row[10].ToString());
                    total += Convert.ToDouble(row[7].ToString());
                }
                else
                {
                    subtotal -= Convert.ToDouble(row[5].ToString());
                    iva -= Convert.ToDouble(row[6].ToString());
                    ivano -= Convert.ToDouble(row[10].ToString());
                    total -= Convert.ToDouble(row[7].ToString());
                }
                if (row[9].ToString() == "False")
                {
                    workSheet2.Cells[i, 11] = "N";
                }
                else
                {
                    workSheet2.Cells[i, 11] = "S";
                }
                i++;
            }
            i++;
            workSheet2.Cells[i, 6] = "TOTALES:";
            workSheet2.Cells[i, 7] = subtotal.ToString();
            workSheet2.Cells[i, 8] = iva.ToString();
            workSheet2.Cells[i, 9] = ivano.ToString();
            workSheet2.Cells[i, 10] = total.ToString();
            Aplic.Visible = true;
            workSheet2.PrintPreview();
            Aplic.Visible = false;
            if (Aplic != null)
            {
                Aplic.Quit();
                System.Diagnostics.Process[] pProcess;
                pProcess = System.Diagnostics.Process.GetProcessesByName("Excel");

                foreach (System.Diagnostics.Process oPro in pProcess)
                {
                    if (oPro.StartTime >= m_FechaHora)
                        oPro.Kill();
                }
            }
        }

        public void solonumeros(KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && (e.KeyChar != '\b'))
            {
                e.Handled = true;
            }
            else
            {
                e.Handled = false;
            }
        }
        public void solonumerosypunto(KeyPressEventArgs e)
        {
            CultureInfo cc = System.Threading.Thread.CurrentThread.CurrentCulture;
            if ((char.IsNumber(e.KeyChar)) || (e.KeyChar.ToString() == cc.NumberFormat.NumberDecimalSeparator) || (e.KeyChar == '\b') || (e.KeyChar.ToString() == "-"))
                e.Handled = false;
            else
                e.Handled = true;
        }

        private void textBox48_KeyPress(object sender, KeyPressEventArgs e)
        {
            /*dataGridView1.DataSource = null;
            if (e.KeyChar == 'H')
            {
                ayuda = "clientes";
                dataGridView1.Visible = true;
                dataGridView1.DataSource = ClaseMedia.obtener_clientes(sigla);
                dataGridView1.AutoResizeColumns();
                label23.Text = "";
            }*/
            solonumeros(e);
        }

        private void textBox53_KeyPress(object sender, KeyPressEventArgs e)
        {
            /*dataGridView1.DataSource = null;
            if (e.KeyChar == 'H')
            {
                ayuda = "productos";
                dataGridView1.Visible = true;
                dataGridView1.DataSource = ClaseMedia.obtener_productos(sigla);
                dataGridView1.AutoResizeColumns();
                label23.Text = "";
            }*/
            solonumeros(e);
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            label23.Text = "";
            if (dataGridView1.CurrentCell.RowIndex > 0)
            {
                if (comboBox19.SelectedIndex == 0)
                {
                    textBox48.Text = dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value.ToString();
                    textBox48.Focus();

                }
                if (comboBox19.SelectedIndex == 1)
                {
                    textBox53.Text = dataGridView1[0, dataGridView1.CurrentCell.RowIndex].Value.ToString();
                    textBox53.Focus();

                }
            }
        }

        private void cancelar_factura()
        {
            factura.Rows.Clear();
            dataGridView2.DataSource = null;
            groupBox18.Enabled = false;
            groupBox17.Enabled = true;
            dateTimePicker4.Focus();
            label23.Text = "";
            subtotal = 0;
            impuestos = 0;
            iva = 0;
            ivano = 0;
            total = 0;
            cant_renglones = 0;
            textBox47.Clear();
            textBox48.Clear();
            textBox49.Clear();
            textBox50.Clear();
            textBox52.Clear();
            textBox18.Text = "0";
            textBox55.Clear();
            textBox56.Clear();
            textBox57.Clear();
            textBox58.Clear();
            textBox59.Clear();
            textBox53.Clear();
            textBox54.Clear();
            textBox34.Clear();
            textBox20.Text = "0"; ;
            comboBox4.SelectedIndex = 0;
            comboBox8.SelectedIndex = 0;
        }

        private void button67_Click(object sender, EventArgs e)
        {
            cancelar_factura();
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (groupBox2.Text != "Alta")
            {
                textBox5.Text = dataGridView3[1, dataGridView3.CurrentCell.RowIndex].Value.ToString();
                textBox5.Focus();
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            if (textBox32.Text == "")
            {
                dataGridView3.DataSource = null;
                dataGridView3.DataSource = ClaseMedia.obtener_productos_segunrubro(sigla, comboBox14.SelectedIndex.ToString());
            }
            else
            {
                dataGridView3.DataSource = null;
                dataGridView3.DataSource = ClaseMedia.obtener_productos_segunnombre(sigla, textBox32.Text);
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (textBox33.Text == "")
            {
                dataGridView4.DataSource = null;
                dataGridView4.DataSource = ClaseMedia.obtener_clientes_segunzona(sigla, comboBox15.SelectedIndex.ToString());
            }
            else
            {
                dataGridView4.DataSource = null;
                dataGridView4.DataSource = ClaseMedia.obtener_clientes_segunnombre(sigla, textBox33.Text);
            }
        }

        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (groupBox1.Text != "Alta")
            {
                textBox1.Text = dataGridView4[0, dataGridView4.CurrentCell.RowIndex].Value.ToString();
                textBox1.Focus();
            }
        }

        private void textBox18_KeyPress(object sender, KeyPressEventArgs e)//Controlo que en la bonificacion del cliente solo ingresen numeros
        {
            solonumerosypunto(e);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if ((textBox36.Text != "")&&(textBox35.Text != ""))
            {
                if (Convert.ToDouble(textBox36.Text) > 0)//Primero controlo que el pago sea mayor a cero
                {
                    bool montobien = true;
                    if (radioButton1.Checked)//Aca hago todos los controles si va a pagar con cheque
                    {
                        if (radioButton6.Checked)//Si va a pagar con un cheque nuevo lo pago
                        {
                            string[] cheque;
                            cheque = new string[8];
                            cheque[0] = textBox51.Text;
                            cheque[1] = textBox46.Text;
                            cheque[2] = dateTimePicker12.Text;
                            cheque[3] = dateTimePicker11.Text;
                            cheque[4] = textBox45.Text;
                            cheque[5] = textBox60.Text;
                            cheque[6] = "1";
                            cheque[7] = "0";
                            if (textBox36.Text == textBox60.Text)//Si el monto del cheque y del pago son iguales, lo cargo
                            {
                                if (chequeusado == false)//Si el cheque aun no fue usado para pagar
                                {
                                    ClaseMedia.cargar_cheque(sigla, cheque);
                                }
                                else
                                {
                                    montobien = false;
                                }
                            }
                            else
                            {
                                montobien = false;
                            }
                        }
                        if (radioButton3.Checked)//Si va a pagar con un cheque existente le actualizo el campo Pago a true, asi no se vuelve a usar despues
                        {
                            if (textBox36.Text == textBox60.Text)
                            {
                                ClaseMedia.marcar_cheque_pago(sigla, textBox46.Text);
                            }
                            else
                            {
                                montobien = false;
                            }
                        }
                    }
                    if (montobien == true)
                    {
                        string[] pago;
                        pago = new string[5];
                        pago[0] = "''";
                        pago[1] = textBox35.Text;
                        pago[2] = textBox36.Text;
                        pago[3] = dateTimePicker9.Text;
                        pago[4] = textBox46.Text;
                        if (ClaseMedia.cargar_pago(sigla, pago) == 0)
                        {
                            label47.Text = "Cargado";
                            textBox35.Clear();
                            textBox36.Clear();
                            textBox45.Clear();
                            textBox46.Clear();
                            textBox51.Clear();
                            textBox60.Clear();
                        }
                        else
                        {
                            label47.Text = "No se pudo cargar";
                        }
                    }
                    else
                    {
                        label47.Text = "El monto del cheque debe ser igual al del pago";
                    }
                }
                else
                {
                    label47.Text = "El monto del pago debe ser mayor a 0";
                }
            }
            else
            {
                label47.Text = "Indique monto y descripcion del pago";
            }
        }

        private void textBox36_Enter(object sender, EventArgs e)
        {
            label47.Text = "";
        }

        private void tabPage3_Enter(object sender, EventArgs e)
        {
            cancelar_factura();
        }

        private void button18_Click(object sender, EventArgs e)
        {
            string[] fac;
            fac = new string[1];
            fac = ClaseMedia.obtener_factura(sigla, textBox38.Text);
            if (fac[0] != null)
            {
                if (fac[19] == "False")
                {
                    textBox40.Text = fac[2];
                    textBox41.Text = fac[3];
                    textBox37.Text = fac[4];
                    textBox39.Text = fac[14];
                    textBox43.Text = fac[18];
                    dateTimePicker10.Text = fac[15];
                    label47.Text = "";
                }
                else
                {
                    label47.Text = "Factura ya anulada";
                }
            }
            else
            {
                label47.Text = "Factura inexistente";
            }
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("Anular Factura?", "", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                if (ClaseMedia.anular_factura(sigla, textBox38.Text, textBox40.Text) == 0)
                {
                    if (textBox40.Text == "Nota de Credito")
                    {
                        ClaseMedia.actualizar_saldo(sigla, textBox41.Text, textBox39.Text,0);
                    }
                    else if ((textBox40.Text == "Nota de Debito") || (textBox40.Text == "Factura"))
                    {
                        ClaseMedia.actualizar_saldo(sigla, textBox41.Text, textBox39.Text,1);
                    }
                    label47.Text = "Factura Anulada";
                }
                else
                {
                    label47.Text = "No se pudo anular";
                }
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (textBox44.Text == "")
            {
                dataGridView1.DataSource = null;
                if (comboBox19.SelectedIndex == 0)//Busco clientes
                {
                    dataGridView1.DataSource = ClaseMedia.obtener_clientes_segunzona(sigla, comboBox18.SelectedIndex.ToString());
                }
                else//Busco productos
                {
                    dataGridView1.DataSource = ClaseMedia.obtener_productos_segunrubro(sigla, comboBox18.SelectedIndex.ToString());
                }
            }
            else
            {
                dataGridView1.DataSource = null;
                if (comboBox19.SelectedIndex == 0)//busco clientes
                {
                    dataGridView1.DataSource = ClaseMedia.obtener_clientes_segunnombre(sigla, textBox44.Text);
                }
                else//busco productos
                {
                    dataGridView1.DataSource = ClaseMedia.obtener_productos_segunnombre(sigla, textBox44.Text);
                }
            }
        }

        private void comboBox19_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox18.Items.Clear();
            if (comboBox19.SelectedIndex == 1)//Va a buscar productos
            {
                DataTable valores = ClaseMedia.obtener_rubros(sigla);
                foreach (DataRow row in valores.Rows)
                {
                    comboBox18.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                }
            }
            else//Cargo las zonas de los clientes
            {
                DataTable valores = ClaseMedia.obtener_zonas(sigla);
                foreach (DataRow row in valores.Rows)
                {
                    comboBox18.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                }
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            bool todos;
            DateTime m_FechaHora = DateTime.Now;
            if (comboBox20.SelectedIndex == comboBox20.Items.Count - 1)
            {
                todos = true;
            }
            else
            {
                todos = false;
            }
            DataTable precios = ClaseMedia.obtener_precios(sigla, comboBox20.SelectedIndex.ToString(), todos);
            nmExcel._Worksheet workSheet2;
            Microsoft.Office.Interop.Excel.Application Aplic = new Microsoft.Office.Interop.Excel.Application();
            Aplic.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook Libro = Aplic.Workbooks.Open("C:/Claudio/ListadoPrecios.xlsx", Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet2 = (nmExcel.Worksheet)Aplic.ActiveSheet;
            int i = 4;
            workSheet2.Cells[1, 9] = comboBox20.Text;
            string rubro = "AAA";
            string subrubro = "AAA";
            double preciociva;
            foreach (DataRow row in precios.Rows)
            {
                if (row[1].ToString() != rubro)
                {
                    i++;
                    workSheet2.Cells[i, 1] = "Rubro: ";
                    workSheet2.Cells[i, 2] = row[1].ToString();
                    rubro = row[1].ToString();
                    i++;
                    workSheet2.Cells[i, 1] = "Subrubro: ";
                    workSheet2.Cells[i, 2] = row[4].ToString();
                    subrubro = row[4].ToString();
                    i++;
                }
                if (row[4].ToString() != subrubro)
                {
                    i++;
                    workSheet2.Cells[i, 1] = "Subrubro: ";
                    workSheet2.Cells[i, 2] = row[4].ToString();
                    subrubro = row[4].ToString();
                    i++;
                }
                workSheet2.Cells[i, 1] = row[0].ToString();
                workSheet2.Cells[i, 2] = row[2].ToString();
                workSheet2.Cells[i, 5] = row[3].ToString();
                workSheet2.Cells[i, 6] = row[5].ToString();
                preciociva = (Convert.ToDouble(row[3]) * Convert.ToDouble(row[5]) / 100) + Convert.ToDouble(row[3]);
                workSheet2.Cells[i, 7] = preciociva.ToString("#0.00");
                //workSheet2.Cells[i, 8] = row[6].ToString();
                //workSheet2.Cells[i, 9] = (preciociva * Convert.ToDouble(row[6])).ToString("#0.00");
                i++;
            }
            Aplic.Visible = true;
            workSheet2.PrintPreview();
            Aplic.Visible = false;
            if (Aplic != null)
            {
                Aplic.Quit();
                System.Diagnostics.Process[] pProcess;
                pProcess = System.Diagnostics.Process.GetProcessesByName("Excel");

                foreach (System.Diagnostics.Process oPro in pProcess)
                {
                    if (oPro.StartTime >= m_FechaHora)
                        oPro.Kill();
                }
            }
        }

        bool chequeusado = true;
        private void textBox46_Leave(object sender, EventArgs e)
        {
            if (radioButton3.Checked)//Si van a usar un cheque existente verifico que no se haya usado para pagar y en ese caso relleno los otros campos
            {
                string[] cheque;
                cheque = new string[1];
                cheque = ClaseMedia.obtener_cheque(sigla, textBox46.Text);
                if (cheque[0] != null)
                {
                    if ((cheque[6] == "False")&&(cheque[7] == "False"))
                    {
                        textBox51.Text = cheque[0];
                        textBox45.Text = cheque[4];
                        dateTimePicker12.Text = cheque[2];
                        dateTimePicker11.Text = cheque[3];
                        label47.Text = "";
                        textBox60.Text = cheque[5];
                        chequeusado = false;
                    }
                    else //Si ya fue usado le aviso y limpio
                    {
                        textBox45.Clear();
                        textBox46.Clear();
                        textBox51.Clear();
                        textBox60.Clear();
                        chequeusado = true;
                        label47.Text = "Cheque cobrado o ya usado para pagar";
                    }
                }
                else//Sino significa que el cheque no existe
                {
                    textBox45.Clear();
                    textBox46.Clear();
                    textBox51.Clear();
                    textBox60.Clear();
                    chequeusado = true;
                    label47.Text = "El cheque ingresado NO existe en la base de datos";
                }
            }
            else //Si va a usar un cheque nuevo verifico que no exista
            {
                string[] cheque;
                cheque = new string[1];
                cheque = ClaseMedia.obtener_cheque(sigla, textBox46.Text);
                if (cheque[0] != null)
                {
                    textBox45.Clear();
                    textBox46.Clear();
                    textBox51.Clear();
                    textBox60.Clear();
                    chequeusado = true;
                    label47.Text = "El cheque ingresado NO es nuevo";
                }
                else
                {
                    chequeusado = false;
                }
            }
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (groupBox11.Enabled == true)
            {
                groupBox11.Enabled = false;
            }
            else
            {
                groupBox11.Enabled = true;
            }
        }

        private void textBox46_Enter(object sender, EventArgs e)
        {
            label47.Text = "";
        }

        private void button22_Click(object sender, EventArgs e)
        {
            DataTable pagos = ClaseMedia.obtener_pagos(sigla, dateTimePicker13.Text, dateTimePicker14.Text);//Busco todos los pagos desde hasta pedidos
            DateTime m_FechaHora = DateTime.Now;
            nmExcel._Worksheet workSheet2;
            Microsoft.Office.Interop.Excel.Application Aplic = new Microsoft.Office.Interop.Excel.Application();
            Aplic.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook Libro = Aplic.Workbooks.Open("C:/Claudio/ListadoPagos.xlsx", Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet2 = (nmExcel.Worksheet)Aplic.ActiveSheet;
            int i = 5;
            workSheet2.Cells[1, 2] = dateTimePicker13.Text;
            workSheet2.Cells[2, 2] = dateTimePicker14.Text;
            string[] cheque;
            cheque = new string[1];
            double totalpagos = 0;
            foreach (DataRow row in pagos.Rows)
            {
                workSheet2.Cells[i, 1] = row[2].ToString();
                workSheet2.Cells[i, 2] = row[0].ToString();
                workSheet2.Cells[i, 5] = row[1].ToString();
                if (row[3].ToString() != "")//Pongo inicial segun sea cheque o efectivo. Si es cheque pongo el nro y el banco
                {
                    workSheet2.Cells[i, 6] = "CH";
                    cheque = ClaseMedia.obtener_cheque(sigla, row[3].ToString());
                    workSheet2.Cells[i, 7] = row[3].ToString();
                    workSheet2.Cells[i, 8] = cheque[0];
                }
                else
                {
                    workSheet2.Cells[i, 6] = "E";
                }
                i++;
                totalpagos += Convert.ToDouble(row[1]);
            }
            i++;
            workSheet2.Cells[i, 4] = "Total:";
            workSheet2.Cells[i, 5] = totalpagos.ToString();
            Aplic.Visible = true;
            workSheet2.PrintPreview();
            Aplic.Visible = false;
            if (Aplic != null)
            {
                Aplic.Quit();
                System.Diagnostics.Process[] pProcess;
                pProcess = System.Diagnostics.Process.GetProcessesByName("Excel");

                foreach (System.Diagnostics.Process oPro in pProcess)
                {
                    if (oPro.StartTime >= m_FechaHora)
                        oPro.Kill();
                }
            }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            DataTable cobros = ClaseMedia.obtener_cobros2(sigla, dateTimePicker15.Text, dateTimePicker16.Text);
            DateTime m_FechaHora = DateTime.Now;
            nmExcel._Worksheet workSheet2;
            Microsoft.Office.Interop.Excel.Application Aplic = new Microsoft.Office.Interop.Excel.Application();
            Aplic.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook Libro = Aplic.Workbooks.Open("C:/Claudio/ListadoCobros.xlsx", Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet2 = (nmExcel.Worksheet)Aplic.ActiveSheet;
            int i = 5;
            workSheet2.Cells[1, 2] = dateTimePicker15.Text;
            workSheet2.Cells[2, 2] = dateTimePicker16.Text;
            string[] cheque;
            string[] cliente;
            cheque = new string[1];
            double totalcobros = 0;
            foreach (DataRow row in cobros.Rows)
            {
                workSheet2.Cells[i, 1] = row[5].ToString();
                workSheet2.Cells[i, 2] = row[4].ToString();
                cliente = ClaseMedia.obtener_cliente(sigla, row[0].ToString());
                workSheet2.Cells[i, 5] = row[0].ToString()+"-"+cliente[1];
                workSheet2.Cells[i, 6] = row[6].ToString();
                if (row[2].ToString() == "1")
                {
                    cheque = ClaseMedia.obtener_cheque(sigla, row[3].ToString());
                    workSheet2.Cells[i, 7] = row[3].ToString();
                    workSheet2.Cells[i, 8] = cheque[0];
                    if (cheque[7] == "False")
                    {
                        workSheet2.Cells[i, 9] = "No";
                    }
                    else
                    {
                        workSheet2.Cells[i, 9] = "Si";
                    }
                }
                else
                {
                    workSheet2.Cells[i, 7] = "Efectivo";
                }
                workSheet2.Cells[i, 10] = row[1].ToString();
                totalcobros += Convert.ToDouble(row[1]);
                i++;
            }
            i++;
            workSheet2.Cells[i, 9] = "Total:";
            workSheet2.Cells[i, 10] = totalcobros.ToString();
            Aplic.Visible = true;
            workSheet2.PrintPreview();
            Aplic.Visible = false;
            if (Aplic != null)
            {
                Aplic.Quit();
                System.Diagnostics.Process[] pProcess;
                pProcess = System.Diagnostics.Process.GetProcessesByName("Excel");

                foreach (System.Diagnostics.Process oPro in pProcess)
                {
                    if (oPro.StartTime >= m_FechaHora)
                        oPro.Kill();
                }
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            string[] cheque;
            cheque = ClaseMedia.obtener_cheque(sigla, textBox67.Text);
            if (cheque[0] != null)
            {
                if ((cheque[6] != "True")&&(cheque[7] != "True"))
                {
                    textBox64.Text = cheque[0];
                    textBox63.Text = cheque[5];
                    textBox62.Text = cheque[4];
                    dateTimePicker18.Text = cheque[2];
                    dateTimePicker17.Text = cheque[3];
                }
                else
                {
                    label47.Text = "Cheque entregado como pago o  ya cobrado";
                }
            }
            else
            {
                label47.Text = "Cheque inexistente";
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            if (ClaseMedia.marcar_cheque_cobro(sigla, textBox67.Text) == 0)
            {
                label47.Text = "Cheque cobrado";
                textBox37.Text = "";
                textBox38.Text = "";
                textBox39.Text = "";
                textBox40.Text = "";
                textBox41.Text = "";
                textBox43.Text = "";
            }
            else
            {
                label47.Text = "Hubo un error";
                textBox38.Text = "";
                textBox39.Text = "";
                textBox40.Text = "";
                textBox41.Text = "";
                textBox43.Text = "";
            }
        }

        private void button28_Click(object sender, EventArgs e)
        {
            DataTable cheques_actuales = ClaseMedia.obtener_cheques_actuales(sigla);
            DataTable cheques_pagados = ClaseMedia.obtener_cheques_pagados(sigla);
            DataTable cheques_cobrados = ClaseMedia.obtener_cheques_cobrados(sigla);
            DateTime m_FechaHora = DateTime.Now;
            nmExcel._Worksheet workSheet2;
            Microsoft.Office.Interop.Excel.Application Aplic = new Microsoft.Office.Interop.Excel.Application();
            Aplic.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook Libro = Aplic.Workbooks.Open("C:/Claudio/ListadoCheques.xlsx", Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet2 = (nmExcel.Worksheet)Aplic.ActiveSheet;
            int i = 3;
            workSheet2.Cells[i, 1] = "CHEQUES ACTUALES";
            i++;
            foreach (DataRow row in cheques_actuales.Rows)
            {
                workSheet2.Cells[i, 1] = row[1].ToString();
                workSheet2.Cells[i, 2] = row[0].ToString();
                workSheet2.Cells[i, 3] = row[4].ToString();
                workSheet2.Cells[i, 4] = row[3].ToString();
                workSheet2.Cells[i, 5] = row[2].ToString();
                workSheet2.Cells[i, 6] = row[5].ToString();
                i++;
            }
            i++;
            workSheet2.Cells[i, 1] = "CHEQUES COBRADOS";
            i++;
            foreach (DataRow row in cheques_cobrados.Rows)
            {
                workSheet2.Cells[i, 1] = row[1].ToString();
                workSheet2.Cells[i, 2] = row[0].ToString();
                workSheet2.Cells[i, 3] = row[4].ToString();
                workSheet2.Cells[i, 4] = row[3].ToString();
                workSheet2.Cells[i, 5] = row[2].ToString();
                workSheet2.Cells[i, 6] = row[5].ToString();
                i++;
            }
            i++;
            workSheet2.Cells[i, 1] = "CHEQUES USADOS PARA PAGAR";
            i++;
            foreach (DataRow row in cheques_pagados.Rows)
            {
                workSheet2.Cells[i, 1] = row[1].ToString();
                workSheet2.Cells[i, 2] = row[0].ToString();
                workSheet2.Cells[i, 3] = row[4].ToString();
                workSheet2.Cells[i, 4] = row[3].ToString();
                workSheet2.Cells[i, 5] = row[2].ToString();
                workSheet2.Cells[i, 6] = row[5].ToString();
                i++;
            }
            Aplic.Visible = true;
            workSheet2.PrintPreview();
            Aplic.Visible = false;
            if (Aplic != null)
            {
                Aplic.Quit();
                System.Diagnostics.Process[] pProcess;
                pProcess = System.Diagnostics.Process.GetProcessesByName("Excel");

                foreach (System.Diagnostics.Process oPro in pProcess)
                {
                    if (oPro.StartTime >= m_FechaHora)
                        oPro.Kill();
                }
            }
        }

        private void button21_Click(object sender, EventArgs e)
        {
            bool todos;//Para saber si quiere todos los rubros
            if (comboBox20.SelectedIndex == comboBox20.Items.Count - 1)
            {
                todos = true;
            }
            else
            {
                todos = false;
            }
            DataTable precios = ClaseMedia.obtener_precios(sigla, comboBox20.SelectedIndex.ToString(), todos);//Obtengo los precios del rubro o de todos
            DateTime m_FechaHora = DateTime.Now;//guardo la hora para matar el procexo exacto despues
            nmExcel._Worksheet workSheet2;
            Microsoft.Office.Interop.Excel.Application Aplic = new Microsoft.Office.Interop.Excel.Application();
            Aplic.Visible = false;
            Microsoft.Office.Interop.Excel.Workbook Libro = Aplic.Workbooks.Open("C:/Claudio/ListadoPrecios.xlsx", Type.Missing, true, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            workSheet2 = (nmExcel.Worksheet)Aplic.ActiveSheet;
            int i = 4;
            workSheet2.Cells[1, 9] = comboBox20.Text;
            string rubro = "AAA";//Valores iniciales para rubro y subrubro, que despues voy reemplazando
            string subrubro = "AAA";
            double preciociva; 
            foreach (DataRow row in precios.Rows)
            {
                if (row[1].ToString() != rubro)//Cuando cambio de rubro pongo nuevas cabeceras
                {
                    i++;
                    workSheet2.Cells[i, 1] = "Rubro: ";
                    workSheet2.Cells[i, 2] = row[1].ToString();
                    rubro = row[1].ToString();
                    i++;
                    workSheet2.Cells[i, 1] = "Subrubro: ";
                    workSheet2.Cells[i, 2] = row[4].ToString();
                    subrubro = row[4].ToString();
                    i++;
                }
                if (row[4].ToString() != subrubro)//Cuando cambio de subrubro pongo nuevas cabeceras
                {
                    i++;
                    workSheet2.Cells[i, 1] = "Subrubro: ";
                    workSheet2.Cells[i, 2] = row[4].ToString();
                    subrubro = row[4].ToString();
                    i++;
                }
                workSheet2.Cells[i, 1] = row[0].ToString();
                workSheet2.Cells[i, 2] = row[2].ToString();
                workSheet2.Cells[i, 5] = row[3].ToString();
                workSheet2.Cells[i, 6] = row[5].ToString();
                preciociva = (Convert.ToDouble(row[3]) * Convert.ToDouble(row[5]) / 100) + Convert.ToDouble(row[3]);
                workSheet2.Cells[i, 7] = preciociva.ToString("#0.00");
                //ESTA PARTE MUESTRA LA CANTIDAD Y EL PRECIO POR LA CANTIDAD. AHORA LO COMENTO PORQUE EL STOCK ESTA DESACTUALIZADO Y HACE LENTO EL LISTADO
                //workSheet2.Cells[i, 8] = row[6].ToString();
                //workSheet2.Cells[i, 9] = (preciociva * Convert.ToDouble(row[6])).ToString("#0.00");
                i++;
            }
            Aplic.Visible = true;
            workSheet2.PrintPreview();
            Aplic.Visible = false;
            if (Aplic != null)
            {
                Aplic.Quit();
                System.Diagnostics.Process[] pProcess;
                pProcess = System.Diagnostics.Process.GetProcessesByName("Excel");

                foreach (System.Diagnostics.Process oPro in pProcess)
                {
                    if (oPro.StartTime >= m_FechaHora)
                        oPro.Kill();
                }
            }
        }

        private void textBox36_KeyPress(object sender, KeyPressEventArgs e)
        {
            solonumerosypunto(e);
        }

        private void textBox60_KeyPress(object sender, KeyPressEventArgs e)
        {
            solonumerosypunto(e);
        }

        private void button30_Click(object sender, EventArgs e)
        {
            bool exito = false;
            string[] registro;
            registro = new string[2];
            registro[0] = textBox72.Text;//Codigo
            registro[1] = textBox73.Text;//Descripcion

            if (groupBox20.Text == "Alta")
            {
                Array.Resize(ref registro, 3);
                registro[2] = "0";
                if (ClaseMedia.insertar_rubro(sigla, registro) == 0)
                {
                    label110.Text = "Cargado correctamente";
                    textBox72.Text = ClaseMedia.obtener_nro_rubro(sigla);//Obtengo el Proximo codigo de rubro a cargar
                    exito = true;
                }
                else
                {
                    label110.Text = "No se pudo cargar";
                }

            }
            else if (groupBox20.Text == "Modificar")
            {
                if (ClaseMedia.rubroexistente(sigla, textBox72.Text) == 0)
                {
                    if (ClaseMedia.modificar_rubro(sigla, registro) == 0)
                    {
                        label110.Text = "Modificado correctamente";
                        exito = true;
                    }
                }
                else
                {
                    label110.Text = "No se pudo modificar";
                }
            }
            else if (groupBox20.Text == "Eliminar")
            {
                if (ClaseMedia.rubroexistente(sigla, textBox72.Text) == 0)
                {
                    if (ClaseMedia.rubroconproductos(sigla, textBox72.Text) != 0)
                    {
                        if (ClaseMedia.borrar_rubro(sigla, textBox72.Text) == 0)
                        {
                            label110.Text = "Eliminado correctamente";
                            exito = true;
                        }
                    }
                    else
                    {
                        label110.Text = "El rubro tiene productos cargados";
                    }
                }
                else
                {
                    label110.Text = "No se pudo eliminar";
                }
            }
            if (exito == true)
            {
                if (groupBox20.Text != "Alta")
                {
                    textBox72.Text = "";
                }
                //limpio los campos, en caso que se cargo
                textBox73.Text = "";
                DataTable valores = ClaseMedia.obtener_rubros(sigla);
                comboBox3.Items.Clear();
                comboBox14.Items.Clear();
                comboBox20.Items.Clear();
                comboBox21.Items.Clear();
                foreach (DataRow row in valores.Rows)
                {
                    comboBox3.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                    comboBox14.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                    comboBox20.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                    comboBox21.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                }
                comboBox20.Items.Add("TODOS");
                comboBox3.SelectedIndex = 0;
                comboBox14.SelectedIndex = 0;
                comboBox20.SelectedIndex = 0;
                comboBox21.SelectedIndex = 0;
            }
        }

        private void textBox72_Leave(object sender, EventArgs e)
        {
            if ((groupBox20.Text == "Modificar") || (groupBox20.Text == "Eliminar"))
            {
                if (ClaseMedia.rubroexistente(sigla, textBox72.Text) != 0)
                {
                    label110.Text = "Rubro inexistente";
                    textBox72.Text = "";
                    textBox73.Text = "";
                }
                else
                {
                    string[] rubro;
                    rubro = new string[9];
                    rubro = ClaseMedia.obtener_rubro(sigla, textBox72.Text);
                    textBox73.Text = rubro[1];
                    label110.Text = "";
                }
            }
        }

        private void button33_Click(object sender, EventArgs e)
        {
            groupBox20.Text = "Alta";
            textBox72.Text = "";
            textBox73.Text = "";
        }

        private void button32_Click(object sender, EventArgs e)
        {
            groupBox20.Text = "Modificar";
            textBox72.Text = "";
            textBox73.Text = "";
        }

        private void button31_Click(object sender, EventArgs e)
        {
            groupBox20.Text = "Eliminar";
            textBox72.Text = "";
            textBox73.Text = "";
        }

        private void button41_Click(object sender, EventArgs e)
        {
            groupBox22.Text = "Alta";
            textBox68.Text = "";
            textBox69.Text = "";
        }

        private void button40_Click(object sender, EventArgs e)
        {
            groupBox22.Text = "Modificar";
            textBox68.Text = "";
            textBox69.Text = "";
        }

        private void button39_Click(object sender, EventArgs e)
        {
            groupBox22.Text = "Eliminar";
            textBox68.Text = "";
            textBox69.Text = "";
        }

        private void textBox68_Leave(object sender, EventArgs e)
        {
            if ((groupBox22.Text == "Modificar") || (groupBox22.Text == "Eliminar"))
            {
                if (ClaseMedia.zonaexistente(sigla, textBox68.Text) != 0)
                {
                    label115.Text = "Zona inexistente";
                    textBox68.Text = "";
                    textBox69.Text = "";
                }
                else
                {
                    string[] zona;
                    zona = new string[2];
                    zona = ClaseMedia.obtener_zona(sigla, textBox68.Text);
                    textBox69.Text = zona[1];
                    label115.Text = "";
                }
            }
        }

        private void button38_Click(object sender, EventArgs e)
        {
            bool exito = false;
            string[] registro;
            registro = new string[2];
            registro[0] = textBox68.Text;//Codigo
            registro[1] = textBox69.Text;//Descripcion

            if (groupBox22.Text == "Alta")
            {
                if (ClaseMedia.insertar_zona(sigla, registro) == 0)
                {
                    label115.Text = "Cargado correctamente";
                    textBox68.Text = ClaseMedia.obtener_nro_zona(sigla);//Obtengo el Proximo codigo de rubro a cargar
                    exito = true;
                }
                else
                {
                    label115.Text = "No se pudo cargar";
                }

            }
            else if (groupBox22.Text == "Modificar")
            {
                if (ClaseMedia.zonaexistente(sigla, textBox68.Text) == 0)
                {
                    if (ClaseMedia.modificar_zona(sigla, registro) == 0)
                    {
                        label115.Text = "Modificado correctamente";
                        exito = true;
                    }
                }
                else
                {
                    label110.Text = "No se pudo modificar";
                }
            }
            else if (groupBox22.Text == "Eliminar")
            {
                if (ClaseMedia.zonaexistente(sigla, textBox68.Text) == 0)
                {
                    if (ClaseMedia.zonaconclientes(sigla, textBox68.Text) != 0)
                    {
                        if (ClaseMedia.borrar_zona(sigla, textBox68.Text) == 0)
                        {
                            label115.Text = "Eliminado correctamente";
                            exito = true;
                        }
                    }
                    else
                    {
                        label115.Text = "Hay clientes que pertenecen a la zona";
                    }
                }
                else
                {
                    label115.Text = "No se pudo eliminar";
                }
            }
            if (exito == true)
            {
                if (groupBox22.Text != "Alta")
                {
                    textBox68.Text = "";
                }
                //limpio los campos, en caso que se cargo
                textBox69.Text = "";
                DataTable valores = ClaseMedia.obtener_zonas(sigla);
                comboBox1.Items.Clear();
                comboBox4.Items.Clear();
                comboBox9.Items.Clear();
                comboBox10.Items.Clear();
                comboBox12.Items.Clear();
                comboBox15.Items.Clear();
                foreach (DataRow row in valores.Rows)
                {
                    comboBox1.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                    comboBox4.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                    comboBox9.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                    comboBox10.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                    comboBox12.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                    comboBox15.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                }
                comboBox9.Items.Add("LISTAR TODAS");
                comboBox12.Items.Add("LISTAR TODAS");
                comboBox1.SelectedIndex = 0;
                comboBox4.SelectedIndex = 0;
                comboBox9.SelectedIndex = 0;
                comboBox10.SelectedIndex = 0;
                comboBox12.SelectedIndex = 0;
                comboBox15.SelectedIndex = 0;
            }
        }

        private void button37_Click(object sender, EventArgs e)
        {
            groupBox21.Text = "Alta";
            comboBox21.SelectedIndex = 0;
            textBox65.Text = "";
            textBox66.Text = "";
        }

        private void button36_Click(object sender, EventArgs e)
        {
            groupBox21.Text = "Modificar";
            comboBox21.SelectedIndex = 0;
            textBox65.Text = "";
            textBox66.Text = "";
        }

        private void button35_Click(object sender, EventArgs e)
        {
            groupBox21.Text = "Eliminar";
            comboBox21.SelectedIndex = 0;
            textBox65.Text = "";
            textBox66.Text = "";
        }

        private void textBox65_Leave(object sender, EventArgs e)
        {
            if ((groupBox21.Text == "Modificar") || (groupBox21.Text == "Eliminar"))
            {
                if (ClaseMedia.subrubroexistente(sigla, textBox68.Text) != 0)
                {
                    label109.Text = "Subrubro inexistente";
                    textBox65.Text = "";
                    textBox66.Text = "";
                }
                else
                {
                    string[] subrubro;
                    subrubro = new string[2];
                    subrubro = ClaseMedia.obtener_subrubro(sigla, textBox65.Text);
                    textBox66.Text = subrubro[1];
                    comboBox21.SelectedIndex = Convert.ToInt32(subrubro[2]);
                    label109.Text = "";
                }
            }
        }

        private void button34_Click(object sender, EventArgs e)
        {
            bool exito = false;
            string[] registro;
            registro = new string[3];
            registro[0] = textBox65.Text;//Codigo
            registro[1] = textBox66.Text;//Descripcion
            registro[2] = comboBox21.SelectedIndex.ToString();//Rubro

            if (groupBox21.Text == "Alta")
            {
                if (ClaseMedia.insertar_subrubro(sigla, registro) == 0)
                {
                    label109.Text = "Cargado correctamente";
                    textBox65.Text = ClaseMedia.obtener_nro_subrubro(sigla);//Obtengo el Proximo codigo de rubro a cargar
                    exito = true;
                }
                else
                {
                    label109.Text = "No se pudo cargar";
                }

            }
            else if (groupBox21.Text == "Modificar")
            {
                if (ClaseMedia.subrubroexistente(sigla, textBox65.Text) == 0)
                {
                    if (ClaseMedia.modificar_subrubro(sigla, registro) == 0)
                    {
                        label109.Text = "Modificado correctamente";
                        exito = true;
                    }
                }
                else
                {
                    label109.Text = "No se pudo modificar";
                }
            }
            else if (groupBox21.Text == "Eliminar")
            {
                if (ClaseMedia.subrubroexistente(sigla, textBox65.Text) == 0)
                {
                    if (ClaseMedia.subrubroconproductos(sigla, textBox65.Text) != 0)
                    {
                        if (ClaseMedia.borrar_subrubro(sigla, textBox65.Text) == 0)
                        {
                            label109.Text = "Eliminado correctamente";
                            exito = true;
                        }
                    }
                    else
                    {
                        label109.Text = "El rubro tiene productos cargados";
                    }
                }
                else
                {
                    label109.Text = "No se pudo eliminar";
                }
            }
            if (exito == true)
            {
                if (groupBox21.Text != "Alta")
                {
                    textBox65.Text = "";
                }
                //limpio los campos, en caso que se cargo
                textBox66.Text = "";
                DataTable valores = ClaseMedia.obtener_subrubros(sigla);
                comboBox17.Items.Clear();
                foreach (DataRow row in valores.Rows)
                {
                    comboBox17.Items.Add(row[0].ToString() + " - " + row[1].ToString());
                }
                comboBox17.SelectedIndex = 0;
            }
        }

        private void button42_Click(object sender, EventArgs e)
        {
            if (comboBox22.Text == "Rubros")
            {
                dataGridView6.DataSource = ClaseMedia.obtener_rubros(sigla);
            }
            else if (comboBox22.Text == "Subrubros")
            {
                dataGridView6.DataSource = ClaseMedia.obtener_subrubros(sigla);
            }
            else if (comboBox22.Text == "Zonas")
            {
                dataGridView6.DataSource = ClaseMedia.obtener_zonas(sigla);
            }

        }

    }
}