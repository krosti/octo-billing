using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using Base;

namespace Intermedia
{
    public class Intermedia2
    {
        public Basededatos ClaseBase;


        public Intermedia2(ref string basedd)
        {
            Initialize(ref basedd);
        }

        public void Initialize(ref string basedd)
        {
            ClaseBase = new Basededatos(ref basedd);
        }

        public string girar_mes_dia(string fecha)
        {
            try
            {
                string dia = Convert.ToString(fecha[8]) + Convert.ToString(fecha[9]);
                string mes = Convert.ToString(fecha[5]) + Convert.ToString(fecha[6]);
                string año = Convert.ToString(fecha[0]) + Convert.ToString(fecha[1]) + Convert.ToString(fecha[2]) + Convert.ToString(fecha[3]);
                return (año + "-" + mes + "-" + dia);
            }
            catch
            {
                return (null);
            }
        }

        public string acomodar_fecha(string fecha)
        {
            try
            {
                int cant = 0; ;
                string mes = "";
                for (int i = 0; i < fecha.Length; i++)
                {
                    cant = i;
                    string car = Convert.ToString(fecha[i]);
                    if (car != "/")
                    {
                        mes += car;
                    }
                    else { break; }
                }
                cant++;
                string dia = "";
                for (int i = cant; i < fecha.Length; i++)
                {
                    cant = i;
                    string car = Convert.ToString(fecha[i]);
                    if (car != "/")
                    {
                        dia += car;
                    }
                    else { break; }
                }
                cant++;
                string año = ""; int an = cant + 4;
                for (int i = cant; i < an; i++)
                {
                    string car = Convert.ToString(fecha[i]);
                    if (car != "/")
                    {
                        año += car;
                    }
                    else { break; }
                }
                if (dia.Length < 2)
                {
                    dia = "0" + dia;
                }
                if (mes.Length < 2)
                {
                    mes = "0" + mes;
                }

                /*
                string dia = Convert.ToString(fecha[0]) + Convert.ToString(fecha[1]);
                string mes = Convert.ToString(fecha[3]) + Convert.ToString(fecha[4]);
                string año = Convert.ToString(fecha[6]) + Convert.ToString(fecha[7]) + Convert.ToString(fecha[8]) + Convert.ToString(fecha[9]);*/
                return (año + "-" + mes + "-" + dia);
            }
            catch
            {
                return (null);
            }
        }

        public string obtener_fecha_inicio(ref string sigla)
        {
            try
            {
                string tabla = "datos" + sigla;
                string clave = "id";
                string valor = "1";
                string[] registro;
                registro = new string[1];
                this.ClaseBase.obtenerRegistro(tabla, clave, valor, ref registro);
                return (registro[1]);
            }
            catch
            {
                return ("Error");
            }
        }

        public string obtener_fecha_fin(ref string sigla)
        {
            try
            {
                string tabla = "datos" + sigla;
                string clave = "id";
                string valor = "1";
                string[] registro;
                registro = new string[1];
                this.ClaseBase.obtenerRegistro(tabla, clave, valor, ref registro);
                return (registro[2]);
            }
            catch
            {
                return ("Error");
            }
        }

        public int insertar_usuario(string sigla, string[] usuario)
        {
            try
            {
                string tabla = "usuarios" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref usuario);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int insertar_sitiva(string sigla, string[] usuario)
        {
            try
            {
                string tabla = "sit_frente_iva" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref usuario);
                return (0);
            }
            catch
            {
                return (1);
            }
        }


        public int insertar_pdv(string[] registro)
        {
            try
            {
                string tabla = "clientes";
                ClaseBase.cargarRegistro(ref tabla, ref registro);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int insertar_provincia(string sigla, string[] registro)
        {
            try
            {
                string tabla = "provincias" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref registro);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int insertar_codigoPostal(string sigla, string[] registro)
        {
            try
            {
                string tabla = "codpostal" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref registro);
                return (0);
            }
            catch
            {
                return (1);
            }
        }




        public int insertar_cod_comprobantes(string sigla, string[] registro)
        {
            try
            {
                string tabla = "codigo_comprobantes" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref registro);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int insertar_tipo_comprobantes(string sigla, string[] registro)
        {
            try
            {
                string tabla = "tipo_comprobantes" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref registro);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int insertar_pdv_tipocomp(string sigla, string[] registro)
        {
            try
            {
                string tabla = "pdv_tipocomp" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref registro);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        

        public int insertar_tipo_tasa(string sigla, string[] registro)
        {
            try
            {
                string tabla = "tipos_tasa" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref registro);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int insertar_valor_tasa(string sigla, string[] registro)
        {
            try
            {
                string tabla = "valores_tasa" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref registro);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int borrar_usuario(string sigla, string usuario)
        {
            try
            {
                string tabla = "usuarios" + sigla;
                string clave = "usuario";
                ClaseBase.borrarregistrio(ref tabla, ref clave, ref usuario);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int borrar_codPostal(string sigla, string usuario)
        {
            try
            {
                string tabla = "codpostal" + sigla;
                string clave = "id";
                ClaseBase.borrarregistrio(ref tabla, ref clave, ref usuario);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int borrar_provincia(string sigla, string usuario)
        {
            try
            {
                string tabla = "provincias" + sigla;
                string clave = "id";
                ClaseBase.borrarregistrio(ref tabla, ref clave, ref usuario);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        

        public int borrar_sitiva(string sigla, string pdv)
        {
            try
            {
                string tabla = "sit_frente_iva" + sigla;
                string clave = "id";
                ClaseBase.borrarregistrio(ref tabla, ref clave, ref pdv);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int borrar_producto(string sigla, string pdv)
        {
            try
            {
                string tabla = "mercaderias" + sigla;
                string clave = "codigo";
                ClaseBase.borrarregistrio(ref tabla, ref clave, ref pdv);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int borrar_cod_comprobantes(string sigla, string pdv)
        {
            try
            {
                string tabla = "codigo_comprobantes" + sigla;
                string clave = "id";
                ClaseBase.borrarregistrio(ref tabla, ref clave, ref pdv);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int borrar_tipo_comprobantes(string sigla, string pdv)
        {
            try
            {
                string tabla = "tipo_comprobantes" + sigla;
                string clave = "id";
                ClaseBase.borrarregistrio(ref tabla, ref clave, ref pdv);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int borrar_pdv_tipocomp(string sigla, string pdv)
        {
            try
            {
                string tabla = "pdv_tipocomp" + sigla;
                string clave = "id";
                ClaseBase.borrarregistrio(ref tabla, ref clave, ref pdv);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int borrar_tipotasa(string sigla, string usuario)
        {
            try
            {
                string tabla = "tipos_tasa" + sigla;
                string clave = "id";
                ClaseBase.borrarregistrio(ref tabla, ref clave, ref usuario);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int borrar_valortasa(string sigla, string usuario)
        {
            try
            {
                string tabla = "valores_tasa" + sigla;
                string clave = "id";
                ClaseBase.borrarregistrio(ref tabla, ref clave, ref usuario);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

       

        
        public int codigo_comprobantes_existente(string sigla, string clave)
        {
            try
            {
                string tabla = "codigo_comprobantes" + sigla;
                string campo = "id";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }

        public int producto_existente(string sigla, string clave)
        {
            try
            {
                string tabla = "mercaderias" + sigla;
                string campo = "codigo";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }

        public int tipo_comprobantes_existente(string sigla, string clave)
        {
            try
            {
                string tabla = "tipo_comprobantes" + sigla;
                string campo = "id";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }

        public int valor_tasa_existente(string sigla, string clave)
        {
            try
            {
                string tabla = "valores_tasa" + sigla;
                string campo = "id";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }

        public int cliente_existente(string sigla, string clave)
        {
            try
            {
                string tabla = "clientes" + sigla;
                string campo = "codigo";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }

        public int cuit_existente(string sigla, string clave)
        {
            try
            {
                string tabla = "clientes" + sigla;
                string campo = "cuit";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }

        public int cuit_existente2(string sigla, string cuit, string id)
        {
            try
            {
                string tabla = "clientes" + sigla;
                if (ClaseBase.elemento_existente2(tabla, cuit, id) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }


        public int codPostal_existente(string sigla, string clave)
        {
            try
            {
                string tabla = "codpostal" + sigla;
                string campo = "id";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }


        public string[] obtener_tipo_comprobante(string sigla, string cuit)
        {
            try
            {
                string tabla = "tipo_comprobantes" + sigla;
                string campo = "id";
                string[] registro;
                registro = new string[1];
                ClaseBase.obtenerRegistro(tabla, campo, cuit, ref registro);
                return (registro);
            }
            catch
            {
                return (null);
            }
        }

        public string[] obtener_producto(string sigla, string cuit)
        {
            try
            {
                string tabla = "mercaderias" + sigla;
                string campo = "codigo";
                string[] registro;
                registro = new string[1];
                ClaseBase.obtenerRegistro(tabla, campo, cuit, ref registro);
                return (registro);
            }
            catch
            {
                return (null);
            }
        }

        public string[] obtener_codpostal(string sigla, string cuit)
        {
            try
            {
                string tabla = "codpostal" + sigla;
                string campo = "id";
                string[] registro;
                registro = new string[1];
                ClaseBase.obtenerRegistro(tabla, campo, cuit, ref registro);
                return (registro);
            }
            catch
            {
                return (null);
            }
        }

        


        public string[] obtener_tipo_emisor_receptor(string sigla, string emisor, string receptor)
        {
            try
            {
                string tabla = "emisor_receptor" + sigla;
                string campo = "emisor";
                string clave = "receptor";
                string[] registro;
                registro = new string[1];
                ClaseBase.obtenerRegistro_segun(tabla, campo, emisor, clave, receptor, ref registro);
                return (registro);
            }
            catch
            {
                return (null);
            }
        }

        

        public string[] obtener_sitiva(string sigla, string id)
        {
            try
            {
                string tabla = "sit_frente_iva" + sigla;
                string campo = "id";
                string[] registro;
                registro = new string[1];
                ClaseBase.obtenerRegistro(tabla, campo, id, ref registro);
                return (registro);
            }
            catch
            {
                return (null);
            }
        }



        public int sitiva_existente(string sigla, string clave)
        {
            try
            {
                string tabla = "sit_frente_iva" + sigla;
                string campo = "id";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }

        public int tipotasa_existente(string sigla, string clave)
        {
            try
            {
                string tabla = "tipos_tasa" + sigla;
                string campo = "id";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }

        public int pdv_tipocomp_existente(string sigla, string clave)
        {
            try
            {
                string tabla = "pdv_tipocomp" + sigla;
                string campo = "id";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }

        

        public int cargar_accion_impresora(string sigla, string acc, string imp)
        {
            try
            {
                string tabla = "accion_impresora" + sigla;
                string[] valores;
                valores = new string[2];
                valores[0] = acc;
                valores[1] = imp;
                ClaseBase.cargarRegistro(ref tabla, ref valores);
                return (0);
            }
            catch
            {
                return (-1);
            }
        }

        public int modificar_usuario(string sigla, string[] usuario)
        {
            try
            {
                string tabla = "usuarios" + sigla;
                string[] campos;
                campos = new string[5];
                campos[0] = "usuario";
                campos[1] = "nombre";
                campos[2] = "apellido";
                campos[3] = "password";
                campos[4] = "categoria";
                ClaseBase.modificarRegistro(ref tabla, ref campos, ref usuario);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int modificar_codPostal(string sigla, string[] usuario)
        {
            try
            {
                string tabla = "codpostal" + sigla;
                string[] campos;
                campos = new string[4];
                campos[0] = "id";
                campos[1] = "numero";
                campos[2] = "localidad";
                campos[3] = "idprovincia";
                ClaseBase.modificarRegistro(ref tabla, ref campos, ref usuario);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int modificar_provincia(string sigla, string[] usuario)
        {
            try
            {
                string tabla = "provincias" + sigla;
                string[] campos;
                campos = new string[2];
                campos[0] = "id";
                campos[1] = "nombre";
                ClaseBase.modificarRegistro(ref tabla, ref campos, ref usuario);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int modificar_sitiva(string sigla, string[] usuario)
        {
            try
            {
                string tabla = "sit_frente_iva" + sigla;
                string[] campos;
                campos = new string[3];
                campos[0] = "id";
                campos[1] = "descripcion";
                campos[2] = "sigla";
                ClaseBase.modificarRegistro(ref tabla, ref campos, ref usuario);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        

        public int modificar_cod_comprobantes(string sigla, string[] pdv)
        {
            try
            {
                string tabla = "codigo_comprobantes" + sigla;
                string[] campos;
                campos = new string[2];
                campos[0] = "id";
                campos[1] = "tipo";
                ClaseBase.modificarRegistro(ref tabla, ref campos, ref pdv);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int modificar_tipo_comprobantes(string sigla, string[] pdv)
        {
            try
            {
                string tabla = "tipo_comprobantes" + sigla;
                string[] campos;
                campos = new string[5];
                campos[0] = "id";
                campos[1] = "tipo";
                campos[2] = "sigla";
                campos[3] = "nombre";
                campos[4] = "signo";
                ClaseBase.modificarRegistro(ref tabla, ref campos, ref pdv);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int modificar_pdv_tipocomp(string sigla, string[] pdv)
        {
            try
            {
                string tabla = "pdv_tipocomp" + sigla;
                string[] campos;
                campos = new string[5];
                campos[0] = "id";
                campos[1] = "idpdv";
                campos[2] = "idtipo_comp";
                campos[3] = "ultimo_emitido";
                campos[4] = "fecha_ult";
                ClaseBase.modificarRegistro(ref tabla, ref campos, ref pdv);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int modificar_tipotasa(string sigla, string[] pdv)
        {
            try
            {
                string tabla = "tipos_tasa" + sigla;
                string[] campos;
                campos = new string[2];
                campos[0] = "id";
                campos[1] = "nombre";
                ClaseBase.modificarRegistro(ref tabla, ref campos, ref pdv);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int modificar_valortasa(string sigla, string[] pdv)
        {
            try
            {
                string tabla = "valores_tasa" + sigla;
                string[] campos;
                campos = new string[5];
                campos[0] = "id";
                campos[1] = "tasa_id";
                campos[2] = "desde";
                campos[3] = "hasta";
                campos[4] = "valor";
                ClaseBase.modificarRegistro(ref tabla, ref campos, ref pdv);
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public string obtener_fecha_valor_tasa(string sigla, string tasa)
        {
            try
            {
                string tabla = "valores_tasa" + sigla;
                string clave = "tasa_id";
                string campo = "hasta";
                return (ClaseBase.fecha_valor_tasa_hasta(tabla, clave, campo, tasa));
            }
            catch
            {
                return (Convert.ToString(-1));
            }
        }

        public DataTable obtener_codigo_comprobantes(string sigla)
        {
            try
            {
                string tabla = "codigo_comprobantes" + sigla;
                string fields = "*";
                return (ClaseBase.obtener_tabla(tabla, fields));
            }
            catch
            {
                return (null);
            }
        }


        /// ////////////////////////
        /// ///////////////////////
        /// NUEVOS
        /// ///////////////////////////
        /// /////////////////////////////

        public int borrar_cliente(string sigla, string cliente)
        {
            try
            {
                string tabla = "clientes" + sigla;
                string clave = "codigo";
                ClaseBase.borrarregistrio(ref tabla, ref clave, ref cliente);
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public int borrar_rubro(string sigla, string rubro)
        {
            try
            {
                string tabla = "rubros" + sigla;
                string clave = "codigo";
                ClaseBase.borrarregistrio(ref tabla, ref clave, ref rubro);
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public int borrar_subrubro(string sigla, string rubro)
        {
            try
            {
                string tabla = "subrubros" + sigla;
                string clave = "codigo";
                ClaseBase.borrarregistrio(ref tabla, ref clave, ref rubro);
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public int borrar_zona(string sigla, string zona)
        {
            try
            {
                string tabla = "zonas" + sigla;
                string clave = "codigo";
                ClaseBase.borrarregistrio(ref tabla, ref clave, ref zona);
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public int clienteexistente(string sigla, string clave)
        {
            try
            {
                string tabla = "clientes" + sigla;
                string campo = "codigo";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }
        public int rubroexistente(string sigla, string clave)
        {
            try
            {
                string tabla = "rubros" + sigla;
                string campo = "codigo";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }
        public int subrubroexistente(string sigla, string clave)
        {
            try
            {
                string tabla = "subrubros" + sigla;
                string campo = "codigo";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }
        public int zonaexistente(string sigla, string clave)
        {
            try
            {
                string tabla = "zonas" + sigla;
                string campo = "codigo";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }
        public int rubroconproductos(string sigla, string clave)
        {
            try
            {
                string tabla = "mercaderias" + sigla;
                string campo = "rubro";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }
        public int subrubroconproductos(string sigla, string clave)
        {
            try
            {
                string tabla = "mercaderias" + sigla;
                string campo = "subrubro";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }
        public int zonaconclientes(string sigla, string clave)
        {
            try
            {
                string tabla = "clientes" + sigla;
                string campo = "zona";
                if (ClaseBase.elemento_existente(tabla, campo, clave) == 0)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }
        public int insertar_producto(string sigla, string[] registro)
        {
            try
            {
                string tabla = "mercaderias" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref registro);
                ClaseBase.subir_id_rubro(sigla, registro[1], registro[0]);
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public int modificar_cliente(string sigla, string[] cliente)
        {
            try
            {
                string tabla = "clientes" + sigla;
                string[] campos;
                campos = new string[10];
                campos[0] = "codigo";
                campos[1] = "razon";
                campos[2] = "domicilio";
                campos[3] = "localidad";
                campos[4] = "zona";
                campos[5] = "telefono";
                campos[6] = "cuit";
                campos[7] = "situacion";
                campos[8] = "saldo";
                campos[9] = "bonificacion";
                ClaseBase.modificarRegistro(ref tabla, ref campos, ref cliente);
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public int modificar_rubro(string sigla, string[] rubro)
        {
            try
            {
                string tabla = "rubros" + sigla;
                string[] campos;
                campos = new string[2];
                campos[0] = "codigo";
                campos[1] = "rubro";
                ClaseBase.modificarRegistro(ref tabla, ref campos, ref rubro);
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public int modificar_subrubro(string sigla, string[] rubro)
        {
            try
            {
                string tabla = "subrubros" + sigla;
                string[] campos;
                campos = new string[3];
                campos[0] = "codigo";
                campos[1] = "descripcion";
                campos[2] = "rubro";
                ClaseBase.modificarRegistro(ref tabla, ref campos, ref rubro);
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public int modificar_zona(string sigla, string[] zona)
        {
            try
            {
                string tabla = "zonas" + sigla;
                string[] campos;
                campos = new string[2];
                campos[0] = "codigo";
                campos[1] = "denominacion";
                ClaseBase.modificarRegistro(ref tabla, ref campos, ref zona);
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public int modificar_producto(string sigla, string[] pdv)
        {
            try
            {
                string tabla = "mercaderias" + sigla;
                string[] campos;
                campos = new string[11];
                campos[0] = "CODIGO";
                campos[1] = "RUBRO";
                campos[2] = "DESCRIP";
                campos[3] = "PRECMIN";
                campos[4] = "PRECMAY";
                campos[5] = "IMPUESTOPR";
                campos[6] = "IMPUESTOPO";
                campos[7] = "STOCKACTUAL";
                campos[8] = "STOCKMIN";
                campos[9] = "SUBRUBRO";
                campos[10] = "TASA";
                ClaseBase.modificarRegistro(ref tabla, ref campos, ref pdv);
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public int insertar_cliente(string sigla, string[] registro)
        {
            try
            {
                string tabla = "clientes" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref registro);
                ClaseBase.subir_id("general" + sigla, "ULTCODCLIE");
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public int insertar_rubro(string sigla, string[] registro)
        {
            try
            {
                string tabla = "rubros" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref registro);
                ClaseBase.subir_id("general" + sigla, "ULTRUBRO");
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public int insertar_subrubro(string sigla, string[] registro)
        {
            try
            {
                string tabla = "subrubros" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref registro);
                ClaseBase.subir_id("general" + sigla, "ULTSUBRUBRO");
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public int insertar_zona(string sigla, string[] registro)
        {
            try
            {
                string tabla = "zonas" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref registro);
                ClaseBase.subir_id("general" + sigla, "ULTZONA");
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public int cargar_pago(string sigla, string[] registro)
        {
            try
            {
                string tabla = "pagos" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref registro);
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public int cargar_comprobante(string sigla, string[] comprobante)
        {
            try
            {
                string tabla = "facturas" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref comprobante);
                if (comprobante[1] == "A")
                {
                    ClaseBase.subir_id("general" + sigla, "ULTCOMPA");
                }
                else
                {
                    ClaseBase.subir_id("general" + sigla, "ULTCOMPB");
                }
                return (0);
            }
            catch
            {
                return (-1);
            }
        }
        public int cargar_renglon(string sigla, string[] comprobante)
        {
            try
            {
                string tabla = "renglones" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref comprobante);
                return (0);
            }
            catch
            {
                return (-1);
            }
        }
        public int cargar_cobro(string sigla, string[] cobro)
        {
            try
            {
                string tabla = "cobros" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref cobro);
                return (0);
            }
            catch
            {
                return (-1);
            }
        }
        public int cargar_cheque(string sigla, string[] cheque)
        {
            try
            {
                string tabla = "cheques" + sigla;
                ClaseBase.cargarRegistro(ref tabla, ref cheque);
                return (0);
            }
            catch
            {
                return (-1);
            }
        }
        public string[] obtener_cliente(string sigla, string codigo)
        {
            try
            {
                string tabla = "clientes" + sigla;
                string campo = "codigo";
                string[] registro;
                registro = new string[1];
                ClaseBase.obtenerRegistro(tabla, campo, codigo, ref registro);
                return (registro);
            }
            catch
            {
                return (null);
            }
        }
        public string[] obtener_rubro(string sigla, string codigo)
        {
            try
            {
                string tabla = "rubros" + sigla;
                string campo = "codigo";
                string[] registro;
                registro = new string[1];
                ClaseBase.obtenerRegistro(tabla, campo, codigo, ref registro);
                return (registro);
            }
            catch
            {
                return (null);
            }
        }
        public string[] obtener_subrubro(string sigla, string codigo)
        {
            try
            {
                string tabla = "subrubros" + sigla;
                string campo = "codigo";
                string[] registro;
                registro = new string[1];
                ClaseBase.obtenerRegistro(tabla, campo, codigo, ref registro);
                return (registro);
            }
            catch
            {
                return (null);
            }
        }
        public string[] obtener_zona(string sigla, string codigo)
        {
            try
            {
                string tabla = "zonas" + sigla;
                string campo = "codigo";
                string[] registro;
                registro = new string[1];
                ClaseBase.obtenerRegistro(tabla, campo, codigo, ref registro);
                return (registro);
            }
            catch
            {
                return (null);
            }
        }
        public string[] obtener_cheque(string sigla, string codigo)
        {
            try
            {
                string tabla = "cheques" + sigla;
                string campo = "numero";
                string[] registro;
                registro = new string[1];
                ClaseBase.obtenerRegistro(tabla, campo, codigo, ref registro);
                return (registro);
            }
            catch
            {
                return (null);
            }
        }
        public int marcar_cheque_pago(string sigla, string cheque)
        {
            try
            {
                ClaseBase.modificarValor("cheques" + sigla, "numero", cheque, "pago", "1");
                return (0);
            }
            catch
            {
                return (-1);
            }
        }
        public int marcar_cheque_cobro(string sigla, string cheque)
        {
            try
            {
                ClaseBase.modificarValor("cheques" + sigla, "numero", cheque, "canjeado", "1");
                return (0);
            }
            catch
            {
                return (-1);
            }
        }
        public DataTable obtener_rubros(string sigla)
        {
            try
            {
                string tabla = "rubros" + sigla;
                string fields = "*";
                return (ClaseBase.obtener_tabla(tabla, fields));
            }
            catch
            {
                return (null);
            }
        }
        public DataTable obtener_subrubros(string sigla)
        {
            try
            {
                string tabla = "subrubros" + sigla;
                string fields = "*";
                return (ClaseBase.obtener_tabla(tabla, fields));
            }
            catch
            {
                return (null);
            }
        }
        public DataTable obtener_sitivas(string sigla)
        {
            try
            {
                string tabla = "sitivas" + sigla;
                string fields = "*";
                return (ClaseBase.obtener_tabla(tabla, fields));
            }
            catch
            {
                return (null);
            }
        }
        public DataTable obtener_zonas(string sigla)
        {
            try
            {
                string tabla = "zonas" + sigla;
                string fields = "*";
                return (ClaseBase.obtener_tabla(tabla, fields));
            }
            catch
            {
                return (null);
            }
        }
        public string obtener_nro_producto(string sigla,string rubro)
        {
            try
            {
                string tabla = "rubros" + sigla;
                string campo = rubro;
                string comp = ClaseBase.obtener_id_alto_rubro(tabla, campo);
                int nro;
                if (comp != null)
                {
                    nro = Convert.ToInt16(comp);
                    return (nro.ToString());
                }
                else
                {
                    return ("1");
                }
            }
            catch
            {
                return (null);
            }
            /*try
            {
                string campo = "ULTCODMERCA";
                string comp = ClaseBase.obtener_id_alto(sigla, campo);
                int nro;
                if (comp != null)
                {
                    nro = Convert.ToInt16(comp) + 1;
                    return (nro.ToString());
                }
                else
                {
                    return ("1");
                }
            }
            catch
            {
                return (null);
            }*/
        }
        public string obtener_nro_cliente(string sigla)
        {
            try
            {
                string campo = "ULTCODCLIE";
                string comp = ClaseBase.obtener_id_alto(sigla, campo);
                int nro;
                if (comp != null)
                {
                    nro = Convert.ToInt16(comp) + 1;
                    return (nro.ToString());
                }
                else
                {
                    return ("1");
                }
            }
            catch
            {
                return (null);
            }
        }
        public string obtener_nro_rubro(string sigla)
        {
            try
            {
                string campo = "ULTRUBRO";
                string comp = ClaseBase.obtener_id_alto(sigla, campo);
                int nro;
                if (comp != null)
                {
                    nro = Convert.ToInt16(comp) + 1;
                    return (nro.ToString());
                }
                else
                {
                    return ("1");
                }
            }
            catch
            {
                return (null);
            }
        }
        public string obtener_nro_subrubro(string sigla)
        {
            try
            {
                string campo = "ULTSUBRUBRO";
                string comp = ClaseBase.obtener_id_alto(sigla, campo);
                int nro;
                if (comp != null)
                {
                    nro = Convert.ToInt16(comp) + 1;
                    return (nro.ToString());
                }
                else
                {
                    return ("1");
                }
            }
            catch
            {
                return (null);
            }
        }
        public string obtener_nro_zona(string sigla)
        {
            try
            {
                string campo = "ULTZONA";
                string comp = ClaseBase.obtener_id_alto(sigla, campo);
                int nro;
                if (comp != null)
                {
                    nro = Convert.ToInt16(comp) + 1;
                    return (nro.ToString());
                }
                else
                {
                    return ("1");
                }
            }
            catch
            {
                return (null);
            }
        }
        public string obtener_nro_facta(string sigla)
        {
            try
            {
                string campo = "ULTCOMPA";
                string comp = ClaseBase.obtener_id_alto(sigla, campo);
                int nro;
                if (comp != null)
                {
                    nro = Convert.ToInt16(comp) + 1;
                    return (nro.ToString());
                }
                else
                {
                    return ("1");
                }
            }
            catch
            {
                return (null);
            }
        }
        public string obtener_nro_factb(string sigla)
        {
            try
            {
                string campo = "ULTCOMPB";
                string comp = ClaseBase.obtener_id_alto(sigla, campo);
                int nro;
                if (comp != null)
                {
                    nro = Convert.ToInt16(comp) + 1;
                    return (nro.ToString());
                }
                else
                {
                    return ("1");
                }
            }
            catch
            {
                return (null);
            }
        }
        public int actualizar_stock(string sigla, DataTable renglones)
        {
            try
            {
                string tabla = "mercaderias" + sigla;
                string campo = "CODIGO";
                string campo2 = "STOCKACTUAL";
                int nuevo;
                double cantidad;
                string codigo = "";
                for (int i = 0; i < renglones.Rows.Count; i++)
                {
                    codigo = renglones.Rows[i][2].ToString();
                    if (codigo != "-")
                    {
                        string[] producto = this.obtener_producto(sigla, codigo);
                        cantidad = Convert.ToDouble(renglones.Rows[i][1]);
                        nuevo = ((Convert.ToInt16(producto[7])) - (Convert.ToInt16(cantidad)));
                        ClaseBase.modificarValor(tabla, campo, renglones.Rows[i][2].ToString(), campo2, nuevo.ToString());
                    }
                }
                return (0);
            }
            catch
            {
                return (-1);
            }
        }
        public int actualizar_saldo(string sigla, string cliente, string monto,int sumaresta)
        {
            try
            {
                string tabla = "clientes" + sigla;
                string campo = "SALDO";
                string clave = "CODIGO";
                string[] regcliente = this.obtener_cliente(sigla, cliente);
                double saldo;
                if (sumaresta == 1)//en este caso sumo
                {
                    saldo = Convert.ToDouble(regcliente[8]) + (Convert.ToDouble(monto));
                }
                else
                {
                    saldo = Convert.ToDouble(regcliente[8]) - (Convert.ToDouble(monto));
                }
                ClaseBase.modificarValor(tabla, clave, cliente, campo, saldo.ToString());
                return (0);
            }
            catch
            {
                return (-1);
            }
        }
        public DataTable obtener_tasas(string sigla)
        {
            try
            {
                string tabla = "tasas" + sigla;
                string fields = "*";
                return (ClaseBase.obtener_tabla(tabla, fields));
            }
            catch
            {
                return (null);
            }
        }
        public DataTable obtener_comprobantes_arepartir(string sigla, string desde, string hasta, string cliente)
        {
            try
            {

                return (ClaseBase.obtener_tabla_listado_arepartir(sigla, desde, hasta, cliente));
            }
            catch
            {
                return (null);
            }
        }
        public DataTable obtener_cobros(string sigla, string cliente)
        {
            return (ClaseBase.obtener_tabla_segun("cobros" + sigla, "codigo,cliente,monto,tipopago,cheque,descripcion,Date_format(fecha,'%Y-%m-%d')", "cliente", cliente, "fecha", "desc"));
        }
        public DataTable obtener_precios(string sigla, string rubro, bool todos)
        {
            try
            {
                string consulta = "";
                if (todos == true)
                {
                    consulta = "select mercaderias" + sigla + ".codigo,rubros" + sigla + ".rubro,mercaderias" + sigla + ".descrip,mercaderias" + sigla + ".precmay,subrubros" + sigla + ".descripcion,tasas" + sigla + ".valor,mercaderias"+sigla+".stockactual from mercaderias" + sigla + ",rubros" + sigla + ",subrubros" + sigla + ",tasas" + sigla + " where mercaderias" + sigla + ".rubro=rubros" + sigla + ".codigo and mercaderias" + sigla + ".subrubro=subrubros" + sigla + ".codigo and mercaderias" + sigla + ".tasa=tasas" + sigla + ".codigo order by mercaderias" + sigla + ".rubro,mercaderias" + sigla + ".subrubro,mercaderias" + sigla + ".descrip";
                }
                else
                {
                    consulta = "select mercaderias" + sigla + ".codigo,rubros" + sigla + ".rubro,mercaderias" + sigla + ".descrip,mercaderias" + sigla + ".precmay,subrubros" + sigla + ".descripcion,tasas" + sigla + ".valor,mercaderias" + sigla + ".stockactual from mercaderias" + sigla + ",rubros" + sigla + ",subrubros" + sigla + ",tasas" + sigla + " where mercaderias" + sigla + ".rubro=rubros" + sigla + ".codigo and mercaderias" + sigla + ".subrubro=subrubros" + sigla + ".codigo and mercaderias" + sigla + ".rubro=" + rubro + " and mercaderias" + sigla + ".tasa=tasas" + sigla + ".codigo order by mercaderias" + sigla + ".rubro,mercaderias" + sigla + ".subrubro,mercaderias" + sigla + ".descrip";
                }
                return (ClaseBase.obtener_tabla_especial(consulta));
            }
            catch
            {
                return (null);
            }

        }
        public DataTable obtener_pagos(string sigla, string desde, string hasta)
        {
            try
            {
                string consulta = "select descripcion,monto,Date_format(fecha,'%Y-%m-%d'),cheque from pagos" + sigla + " where (fecha BETWEEN '"+desde+"' and '"+hasta+"') order by pagos" + sigla + ".fecha";
                return (ClaseBase.obtener_tabla_especial(consulta));
            }
            catch
            {
                return (null);
            }

        }
        public DataTable obtener_cheques_actuales(string sigla)
        {
            try
            {
                string consulta = "select banco,numero,Date_format(fechacobro,'%Y-%m-%d'),Date_format(fechaemision,'%Y-%m-%d'),emisor,monto from cheques" + sigla + " where pago=0 and canjeado=0 order by banco DESC";
                return (ClaseBase.obtener_tabla_especial(consulta));
            }
            catch
            {
                return (null);
            }

        }
        public DataTable obtener_cheques_pagados(string sigla)
        {
            try
            {
                string consulta = "select banco,numero,Date_format(fechacobro,'%Y-%m-%d'),Date_format(fechaemision,'%Y-%m-%d'),emisor,monto from cheques" + sigla + " where pago=1 order by banco DESC";
                return (ClaseBase.obtener_tabla_especial(consulta));
            }
            catch
            {
                return (null);
            }

        }
        public DataTable obtener_cheques_cobrados(string sigla)
        {
            try
            {
                string consulta = "select banco,numero,Date_format(fechacobro,'%Y-%m-%d'),Date_format(fechaemision,'%Y-%m-%d'),emisor,monto from cheques" + sigla + " where canjeado=1 order by banco DESC";
                return (ClaseBase.obtener_tabla_especial(consulta));
            }
            catch
            {
                return (null);
            }

        }
        public DataTable obtener_cobros2(string sigla, string desde, string hasta)
        {
            try
            {
                string consulta = "select cliente,monto,tipopago,cheque,descripcion,Date_format(fecha,'%Y-%m-%d'),comprobante from cobros" + sigla + " where (fecha BETWEEN '" + desde + "' and '" + hasta + "') order by cobros" + sigla + ".fecha";
                return (ClaseBase.obtener_tabla_especial(consulta));
            }
            catch
            {
                return (null);
            }

        }
        public DataTable obtener_comprobantes(string sigla,string desde, string hasta, string tipo, bool impresas, string zona, bool todas,bool solonegro)
        {
            string consulta = "";
            if (solonegro == true)
            {
                consulta = "select id,tipo,documento,cliente,razon,subtotal,iva,total,Date_format(fecha,'%Y-%m-%d'),impresa,ivano from facturas" + sigla + " where (fecha between '" + desde + "' and '" + hasta + "') and (anulada = 0) and (impresa = 0)";
            }
            else
            {
                if ((tipo == "A") && (impresas == true))
                {
                    if (todas == true)
                    {
                        consulta = "select id,tipo,documento,cliente,razon,subtotal,iva,total,Date_format(fecha,'%Y-%m-%d'),impresa,ivano from facturas" + sigla + " where (fecha between '" + desde + "' and '" + hasta + "') and (tipo = 'A') and (anulada = 0) order by fecha desc";
                    }
                    else
                    {
                        consulta = "select id,tipo,documento,cliente,razon,subtotal,iva,total,Date_format(fecha,'%Y-%m-%d'),impresa,ivano from facturas" + sigla + " where (fecha between '" + desde + "' and '" + hasta + "') and (tipo = 'A') and (zona = " + zona + ") and (anulada = 0) order by fecha desc";
                    }
                }
                else if ((tipo == "A") && (impresas == false))
                {
                    if (todas == true)
                    {
                        consulta = "select id,tipo,documento,cliente,razon,subtotal,iva,total,Date_format(fecha,'%Y-%m-%d'),impresa,ivano from facturas" + sigla + " where (fecha between '" + desde + "' and '" + hasta + "') and (tipo = 'A') and (impresa = 1) and (anulada = 0) order by fecha desc";
                    }
                    else
                    {
                        consulta = "select id,tipo,documento,cliente,razon,subtotal,iva,total,Date_format(fecha,'%Y-%m-%d'),impresa,ivano from facturas" + sigla + " where (fecha between '" + desde + "' and '" + hasta + "') and (tipo = 'A') and (zona = " + zona + ") and (impresa = 1) and (anulada = 0) order by fecha desc";
                    }
                }
                else if ((tipo == "B") && (impresas == false))
                {
                    if (todas == true)
                    {
                        consulta = "select id,tipo,documento,cliente,razon,subtotal,iva,total,Date_format(fecha,'%Y-%m-%d'),impresa,ivano from facturas" + sigla + " where (fecha between '" + desde + "' and '" + hasta + "') and (tipo = 'B') and (impresa = 1) and (anulada = 0) order by fecha desc";
                    }
                    else
                    {
                        consulta = "select id,tipo,documento,cliente,razon,subtotal,iva,total,Date_format(fecha,'%Y-%m-%d'),impresa,ivano from facturas" + sigla + " where (fecha between '" + desde + "' and '" + hasta + "') and (tipo = 'B') and (zona = " + zona + ") and (impresa = 1) and (anulada = 0) order by fecha desc";
                    }
                }
                else if ((tipo == "B") && (impresas == true))
                {
                    if (todas == true)
                    {
                        consulta = "select id,tipo,documento,cliente,razon,subtotal,iva,total,Date_format(fecha,'%Y-%m-%d'),impresa,ivano from facturas" + sigla + " where (fecha between '" + desde + "' and '" + hasta + "') and (tipo = 'B') and (anulada = 0) order by fecha desc";
                    }
                    else
                    {
                        consulta = "select id,tipo,documento,cliente,razon,subtotal,iva,total,Date_format(fecha,'%Y-%m-%d'),impresa,ivano from facturas" + sigla + " where (fecha between '" + desde + "' and '" + hasta + "') and (tipo = 'B') and (zona = " + zona + ") and (anulada = 0)order by fecha desc";
                    }
                }
                else if ((tipo == "TODOS") && (impresas == false))
                {
                    if (todas == true)
                    {
                        consulta = "select id,tipo,documento,cliente,razon,subtotal,iva,total,Date_format(fecha,'%Y-%m-%d'),impresa,ivano from facturas" + sigla + " where (fecha between '" + desde + "' and '" + hasta + "') and (impresa = 1) and (anulada = 0) order by fecha desc";
                    }
                    else
                    {
                        consulta = "select id,tipo,documento,cliente,razon,subtotal,iva,total,Date_format(fecha,'%Y-%m-%d'),impresa,ivano from facturas" + sigla + " where (fecha between '" + desde + "' and '" + hasta + "') and (zona = " + zona + ") and (impresa = 1) and (anulada = 0) order by fecha desc";
                    }
                }
                else
                {
                    if (todas == true)
                    {
                        consulta = "select id,tipo,documento,cliente,razon,subtotal,iva,total,Date_format(fecha,'%Y-%m-%d'),impresa,ivano from facturas" + sigla + " where (fecha between '" + desde + "' and '" + hasta + "') and (anulada = 0)";
                    }
                    else
                    {
                        consulta = "select id,tipo,documento,cliente,razon,subtotal,iva,total,Date_format(fecha,'%Y-%m-%d'),impresa,ivano from facturas" + sigla + " where (fecha between '" + desde + "' and '" + hasta + "') and (zona = " + zona + ") and (anulada = 0)";
                    }
                }
            }
            return(ClaseBase.obtener_tabla_especial(consulta));
        }
        public DataTable obtener_productos(string sigla)
        {
            try
            {
                string tabla = "mercaderias" + sigla;
                string fields = "*";
                return (ClaseBase.obtener_tabla(tabla, fields));
            }
            catch
            {
                return (null);
            }
        }
        public DataTable obtener_clientes(string sigla)
        {
            try
            {
                string tabla = "clientes" + sigla;
                string fields = "*";
                return (ClaseBase.obtener_tabla(tabla, fields));
            }
            catch
            {
                return (null);
            }
        }
        public DataTable obtener_productos_segunrubro(string sigla, string rubro)
        {
            return ClaseBase.obtener_tabla_segun("mercaderias" + sigla, "rubro", rubro);
        }
        public DataTable obtener_productos_segunnombre(string sigla, string nombre)
        {
            return ClaseBase.obtener_tabla_segun("mercaderias" + sigla, "descrip", nombre);
        }
        public DataTable obtener_clientes_segunzona(string sigla, string zona)
        {
            return ClaseBase.obtener_tabla_segun("clientes" + sigla, "zona", zona);
        }
        public DataTable obtener_clientes_segunnombre(string sigla, string nombre)
        {
            return ClaseBase.obtener_tabla_segun("clientes" + sigla, "razon", nombre);
        }
        public string obtener_valor_tasa(string sigla, string tasa)
        {
            string clave = "codigo";
            string[] registro;
            registro = new string[1];
            ClaseBase.obtenerRegistro("tasas"+sigla, clave, tasa, ref registro);
            return (registro[2]);
        }
        public DataTable obtener_comprobantes_cliente_adeuda(string sigla, string cliente)
        {
            string consulta = "SELECT Date_format(fecha,'%Y-%m-%d'),id,documento,subtotal,iva,total,impresa,pagado FROM facturas" + sigla + " where (cliente='" + cliente + "') and (facturas" + sigla + ".pagado<facturas" + sigla + ".total) and (facturas" + sigla + ".anulada = 0) order by facturas" + sigla + ".fecha DESC";
            return(ClaseBase.obtener_tabla_especial(consulta));
        }
        public DataTable obtener_comprobantes_cliente(string sigla, string cliente)
        {
            string consulta = "SELECT Date_format(fecha,'%Y-%m-%d'),id,documento,subtotal,iva,total,impresa,pagado FROM facturas" + sigla + " where (cliente='" + cliente + "') and (facturas"+sigla+".anulada = 0)order by facturas" + sigla + ".fecha DESC";
            return (ClaseBase.obtener_tabla_especial(consulta));
        }
        public string[] obtener_factura(string sigla, string factura)
        {
            try
            {
                string[] registro; registro = new string[1];
                ClaseBase.obtenerRegistro("facturas" + sigla, "id", factura, ref registro);
                return (registro);
            }
            catch
            {
                return (null);
            }
        }
        public int actualizar_saldo_factura(string sigla, string[] factura)
        {
            try
            {
                string tabla = "facturas" + sigla;
                string[] campos;
                campos = new string[2];
                campos[0] = "id";
                campos[1] = "pagado";
                ClaseBase.modificarRegistro(ref tabla, ref campos, ref factura);
                return (0);
            }
            catch
            {
                return (1);
            }
        }
        public DataTable obtener_clientes_listado_arepartir(string sigla, string desde, string hasta, string zona, bool todas)
        {          
            try
            {
                if (todas == false)
                {
                    string consulta = "select clientes" + sigla + ".codigo,clientes" + sigla + ".razon,clientes" + sigla + ".saldo from facturas" + sigla + ",clientesdtb where (facturas" + sigla + ".fecha between '" + desde + "' and '" + hasta + "') and (facturas" + sigla + ".cliente = clientes" + sigla + ".codigo) and (clientes" + sigla + ".zona = '" + zona + "') and (clientes" + sigla + ".codigo <> 50000) and (facturas"+sigla+".anulada = 0) group by facturas" + sigla + ".cliente order by clientes" + sigla + ".razon ASC";
                    return ClaseBase.obtener_tabla_especial(consulta);
                }
                else
                {
                    string consulta = "select clientes" + sigla + ".codigo,clientes" + sigla + ".razon,clientes" + sigla + ".saldo from facturas" + sigla + ",clientesdtb where (facturas" + sigla + ".fecha between '" + desde + "' and '" + hasta + "') and (facturas" + sigla + ".cliente = clientes" + sigla + ".codigo) and (clientes" + sigla + ".codigo <> 50000)  and (facturas" + sigla + ".anulada = 0) group by facturas" + sigla + ".cliente order by clientes" + sigla + ".razon ASC";
                    return ClaseBase.obtener_tabla_especial(consulta);
                }
            }
            catch
            {
                return (null);
            }
        }
        public int anular_factura(string sigla, string nro, string documento)
        {
            try
            {
                DataTable renglones = ClaseBase.obtener_tabla_segun("renglones" + sigla, "factura", nro);
                string[] registro;
                registro = new string[1];
                ClaseBase.obtenerRegistro("comprobantes" + sigla, "documento", documento, ref registro);
                string[] producto;
                producto = new string[1];
                double cantidad;
                foreach (DataRow renglon in renglones.Rows)
                {
                    ClaseBase.obtenerRegistro("mercaderias" + sigla, "codigo", renglon[2].ToString(), ref producto);
                    if (registro[2] == "True")
                    {
                        cantidad = Convert.ToDouble(producto[7]) - Convert.ToDouble(renglon[4]);
                    }
                    else
                    {
                        cantidad = Convert.ToDouble(producto[7]) + Convert.ToDouble(renglon[4]);   
                    }
                    ClaseBase.modificarValor("mercaderias" + sigla, "codigo", renglon[2].ToString(), "stockactual", cantidad.ToString());
                }
                ClaseBase.modificarValor("facturas" + sigla, "id", nro, "anulada", "1");
                return (0);
            }
            catch
            {
                return (-1);
            }
        }

        /// ////////////////////////
        /// ///////////////////////
        /// NUEVOS
        /// ///////////////////////////
        /// /////////////////////////////

        public DataTable obtener_tipo_comprobantes(string sigla)
        {
            try
            {
                string tabla = "tipo_comprobantes" + sigla;
                string fields = "*";
                return (ClaseBase.obtener_tabla(tabla, fields));
            }
            catch
            {
                return (null);
            }
        }

        public DataTable obtener_pdv_tipocomp(string sigla)
        {
            try
            {
                string tabla = "pdv_tipocomp" + sigla;
                string fields = "*";
                return (ClaseBase.obtener_tabla(tabla, fields));
            }
            catch
            {
                return (null);
            }
        }

        public DataTable obtener_pdvs(string sigla)
        {
            try
            {
                string tabla = "pdv" + sigla;
                string fields = "*";
                return (ClaseBase.obtener_tabla(tabla, fields));
            }
            catch
            {
                return (null);
            }
        }

        public DataTable obtener_codPostales(string sigla)
        {
            try
            {
                string tabla = "codpostal" + sigla;
                string fields = "*";
                return (ClaseBase.obtener_tabla(tabla, fields));
            }
            catch
            {
                return (null);
            }
        }
        

        public DataTable obtener_clientes_listado(string sigla)
        {
            try
            {
                string tabla = "clientes" + sigla;
                string fields = "id, cuit, nombre, apellido, sitiva, direccion,idcodpostal,saldo";
                return (ClaseBase.obtener_tabla(tabla, fields));
            }
            catch
            {
                return (null);
            }
        }

        

        public DataTable obtener_valoresTasas(string sigla)
        {
            try
            {
                string tabla = "valores_tasa" + sigla;
                string fields = "*";
                return (ClaseBase.obtener_tabla(tabla, fields));
            }
            catch
            {
                return (null);
            }
        }

        public string seleccionar_id(string cadena)
        {
            try
            {
                char j;
                string retorno = "";
                for (int i = 0; i < cadena.Length; i++)
                {
                    j = cadena[i];
                    if (j.ToString() != " ")
                    {
                        retorno = retorno + j;
                    }
                    else
                    {
                        break;
                    }
                }
                return (retorno);
            }
            catch
            {
                return (null);
            }
        }

        

      

        public int elemento_en_datatable(string elem, DataTable tabla, int columna)
        {
            try
            {
                int retorno = 1;
                for (int i = 0; i <= tabla.Rows.Count; i++)
                {
                    if (tabla.Rows[i][columna].ToString() == elem)
                    {
                        retorno = 0;
                        break;
                    }
                }
                return (retorno);
            }
            catch
            {
                return (1);
            }
        }

        public string[] obtener_nro_comp(string sigla, string pdv, string tipocomp)
        {
            try
            {
                string tabla = "pdv_tipocomp" + sigla;
                string clave = "idpdv";
                string clave2 = "idtipo_comp";
                string[] registro;
                registro = new string[1];
                ClaseBase.obtenerRegistro_segun(tabla, clave, pdv, clave2, tipocomp, ref registro);
                return (registro);
            }
            catch
            {
                return (null);
            }
        }

        public int actualizar_nro_comp(string sigla, string valor, string nro)
        {
            try
            {
                string tabla = "pdv_tipocomp" + sigla;
                string clave = "id";
                string campo = "ultimo_emitido";
                return (ClaseBase.modificarValor(tabla, clave, valor, campo, nro));
            }
            catch
            {
                return (-1);
            }
        }

        public int actualizar_fechas_comp(string sigla, string fecha)
        {
            try
            {
                string tabla = "pdv_tipocomp" + sigla;
                string campo = "fecha_ult";
                ClaseBase.actualizar_columna(tabla, campo, fecha);
                return (0);
            }
            catch
            {
                return (-1);
            }
        }

        public string obtener_nro_comprobante(string sigla)
        {
            try
            {
                string tabla = "comprobantes" + sigla;
                string campo = "id";
                string comp = ClaseBase.obtener_id_alto(tabla, campo);
                int nro;
                if (comp != null)
                {
                    nro = Convert.ToInt16(comp) + 1;
                    return (nro.ToString());
                }
                else
                {
                    return ("1");
                }
            }
            catch
            {
                return (null);
            }
        }

        public string obtener_tipotasa(string sigla, string tasa)
        {
            try
            {
                string tabla = "tipos_tasa" + sigla;
                string[] registro;
                registro = new string[1];
                ClaseBase.obtenerRegistro(tabla, "id", tasa, ref registro);
                return (registro[1]);
            }
            catch
            {
                return (null);
            }
        }

        public int guardar_renglones(string sigla, string idcomp, DataTable renglones)
        {
            try
            {
                string tabla = "renglonescomp" + sigla;
                string[] valores;
                valores = new string[10];
                valores[0] = "";
                valores[1] = idcomp;
                for (int i = 0; i <= renglones.Rows.Count; i++)
                {

                    valores[2] = (i + 1).ToString();
                    valores[3] = renglones.Rows[i][0].ToString();
                    valores[4] = renglones.Rows[i][1].ToString();
                    valores[5] = renglones.Rows[i][2].ToString();
                    valores[6] = renglones.Rows[i][3].ToString();
                    valores[7] = renglones.Rows[i][4].ToString();
                    valores[8] = renglones.Rows[i][5].ToString();
                    valores[9] = renglones.Rows[i][6].ToString();
                    ClaseBase.cargarRegistroparche(ref tabla, ref valores);
                }
                return (0);
            }
            catch
            {
                return (-1);
            }
        }

        public int fecha_comprobante_mayorigual(string sigla, string fecha)
        {
            try
            {
                string tabla = "pdv_tipocomp" + sigla;
                string[] registro;
                registro = new string[1];
                string clave = "id";
                string valor = "1";
                ClaseBase.obtenerRegistro(tabla, clave, valor, ref registro);
                if (string.Compare(this.girar_mes_dia(this.acomodar_fecha(registro[4])), fecha) != 1)
                {
                    return (0);
                }
                else
                {
                    return (1);
                }
            }
            catch
            {
                return (-1);
            }
        }

        public int anular_comprobante(string sigla, string comprobante)
        {
            try
            {
                string tabla = "comprobantes" + sigla;
                ClaseBase.modificarValor(tabla, "idpdvtipocomp", comprobante, "estado", "0");
                return (0);
            }
            catch
            {
                return (-1);
            }
        }

    }
}
