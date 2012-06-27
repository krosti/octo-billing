using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using MySql.Data.MySqlClient;
using System.Data;

namespace Base
{
    public class Basededatos
    {
        public MySqlConnection connection;

        public Basededatos(ref string basedd)
        {
            Initialize(ref basedd);
        }

        public void Initialize(ref string basedd)
        {
            string conectionString = "SERVER=localhost; UID=root; database =" + basedd + ";PASSWORD=;";
            connection = new MySqlConnection(conectionString);
        }

        public static int crearBase(ref string nombre)
        {
            string MyConString = "SERVER=localhost; UID=root; PASSWORD=;";
            MySqlConnection connection = new MySqlConnection(MyConString);
            try
            {
                MySqlCommand command = connection.CreateCommand();
                command.CommandText = "create database" + nombre;
                connection.Open();
                command.ExecuteNonQuery();
                connection.Close();
                return (0); // Ejecución correcta
            }
            catch
            {
                return (1); //Error de conexión
            }
        }

        public bool OpenConnection()
        {
            try
            {
                connection.Open();
                return true;
            }
            catch (Exception ex)
            {
                if (ex is MySqlException)
                {
                    switch (((MySqlException)ex).Number)
                    {
                        case 0:
                            Console.WriteLine("Cannot connect to server.  Contact administrator");
                            break;

                        case 1045:
                            Console.WriteLine("Invalid username/password, please try again");
                            break;
                    }
                    return false;
                }
                else if (ex is InvalidOperationException)
                {
                    return true;
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("Otra excepcion");
                    return false;
                }
            }
        }

        public bool CloseConnection()
        {
            try
            {
                connection.Close();
                return true;
            }
            catch (MySqlException ex)
            {
                Console.WriteLine(ex.Message);
                return false;
            }
        }


        public static int conectarBase(ref string nombre)
        {
            string MyConString = "SERVER=localhost; DATABASE=" + nombre + "; UID=root; PASSWORD=;";
            MySqlConnection connection = new MySqlConnection(MyConString);
            try
            {
                connection.Open();
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public static int desonectarBase(ref string nombre)
        {
            string MyConString = "SERVER=localhost; DATABASE=" + nombre + "; UID=root; PASSWORD=;";
            MySqlConnection connection = new MySqlConnection(MyConString);
            try
            {
                connection.Close();
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public static int crearCampo(ref string tabla, ref string campo, ref string tipo, ref string largo, ref bool clave)
        {
            try
            {
                MySqlCommand command = new MySqlCommand();
                if (clave)
                {
                    command.CommandText = "alter table" + tabla + "add primary key(" + campo + ") " + tipo + "(" + largo + ")";
                }
                else
                {
                    command.CommandText = "alter table" + tabla + "add " + campo + " " + tipo + "(" + largo + ")";
                }
                command.ExecuteNonQuery();
                return (0);
            }
            catch
            {
                return (1);
            }

        }

        public int crearTabla(ref string nombre)
        {
            this.OpenConnection();
            string consulta = "CREATE TABLE IF NOT EXISTS " + nombre;
            MySqlCommand command = new MySqlCommand(consulta, connection);
            command.ExecuteNonQuery();
            try
            {
                this.CloseConnection();
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int vaciarTabla(ref string nombre)
        {
            this.OpenConnection();
            string consulta = "truncate " + nombre;
            MySqlCommand command = new MySqlCommand(consulta, connection);
            command.ExecuteNonQuery();
            try
            {
                this.CloseConnection();
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int copiarTabla(ref string nombre, ref string tabla)
        {
            this.OpenConnection();
            string consulta = "CREATE TABLE " + nombre + " AS (SELECT * FROM " + tabla + ")";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            command.ExecuteNonQuery();
            try
            {
                this.CloseConnection();
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int cargarRegistro(ref string tabla, ref string[] valores)
        {
            string palabra = "'";
            for (int i = 0; i < ((valores.Length) - 1); i++)
            {
                palabra += valores[i] + "','";
            }
            palabra += valores[((valores.Length) - 1)] + "'";
            this.CloseConnection();
            this.OpenConnection();
            string consulta = "insert into " + tabla + " values (" + palabra + ")";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            command.ExecuteNonQuery();
            try
            {
                this.CloseConnection();
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int cargarRegistroparche(ref string tabla, ref string[] valores)
        {
            string palabra = "'";
            for (int i = 0; i < ((valores.Length) - 1); i++)
            {
                palabra += valores[i] + "','";
            }
            palabra += valores[((valores.Length) - 1)] + "'";
            this.CloseConnection();
            this.OpenConnection();
            string consulta = "insert into " + tabla + " values (" + palabra + ")";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            command.ExecuteNonQuery();
            try
            {
                this.CloseConnection();
                return (0);
            }
            catch
            {
                return (1);
            }
        }


        public int elemento_existente(string tabla, string campo, string valor)
        {
            this.CloseConnection();
            this.OpenConnection();
            string consulta = "select count(*) from " + tabla + " where " + campo + "='" + valor + "'";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            MySqlDataReader Reader;
            try
            {
                Reader = command.ExecuteReader();
                int cant = 0;
                if (Reader.Read())
                {
                    cant = int.Parse(Reader.GetValue(0).ToString());
                }
                this.CloseConnection();
                if (cant > 0)
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

        public int elemento_existente2(string tabla, string cuit, string id)
        {
            this.CloseConnection();
            this.OpenConnection();
            string consulta = "select count(*) from " + tabla + " where (cuit ='" + cuit + "') and (id not like " + id + ")";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            MySqlDataReader Reader;
            try
            {
                Reader = command.ExecuteReader();
                int cant = 0;
                if (Reader.Read())
                {
                    cant = int.Parse(Reader.GetValue(0).ToString());
                }
                this.CloseConnection();
                if (cant > 0)
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


        public int modificarValor(string tabla, string clave, string valor, string campo, string nuevo)
        {

            this.OpenConnection();
            string consulta = "update " + tabla + " set " + campo + "='" + nuevo + "' where " + clave + "='" + valor + "'";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            command.ExecuteNonQuery();
            try
            {
                this.CloseConnection();
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public string subir_id(string tabla, string campo)
        {
            this.OpenConnection();
            string consulta = "update " + tabla + " set " + campo + " = " + campo + "+1";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            MySqlDataReader Reader;
            try
            {
                Reader = command.ExecuteReader();
                if (Reader.Read())
                {
                    string resultado = Reader.GetValue(0).ToString();
                    this.CloseConnection();
                    return (resultado);
                }
                else
                {
                    return (null);
                }
            }
            catch
            {
                return (null);
            }
        }
        public string subir_id_rubro(string sigla, string rubro, string codigo)
        {
            this.OpenConnection();
            string consulta = "update rubros" + sigla + " set contador="+codigo+" WHERE codigo = " + rubro;
            MySqlCommand command = new MySqlCommand(consulta, connection);
            MySqlDataReader Reader;
            try
            {
                Reader = command.ExecuteReader();
                if (Reader.Read())
                {
                    string resultado = Reader.GetValue(0).ToString();
                    this.CloseConnection();
                    return (resultado);
                }
                else
                {
                    return (null);
                }
            }
            catch
            {
                return (null);
            }
        }
        public string obtener_id_alto(string tabla, string campo)
        {
            this.CloseConnection();
            this.OpenConnection();
            string consulta = "select " + campo + " from general" + tabla + " where id=1";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            MySqlDataReader Reader;
            try
            {
                Reader = command.ExecuteReader();
                if (Reader.Read())
                {
                    string resultado = Reader.GetValue(0).ToString();
                    this.CloseConnection();
                    return (resultado);
                }
                else
                {
                    return (null);
                }
            }
            catch
            {
                return (null);
            }
        }
        public string obtener_id_alto_rubro(string tabla, string campo)
        {
            this.CloseConnection();
            this.OpenConnection();
            string consulta = "select contador from " + tabla + " where codigo='" + campo + "'";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            MySqlDataReader Reader;
            try
            {
                Reader = command.ExecuteReader();
                if (Reader.Read())
                {
                    string resultado = Reader.GetValue(0).ToString();
                    this.CloseConnection();
                    return (resultado);
                }
                else
                {
                    return (null);
                }
            }
            catch
            {
                return (null);
            }
        }

        public int actualizar_columna(string tabla, string campo, string valor)
        {

            this.OpenConnection();
            string consulta = "update " + tabla + " set " + campo + "='" + valor + "'";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            command.ExecuteNonQuery();
            try
            {
                this.CloseConnection();
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int borrarregistrio(ref string tabla, ref string clave, ref string valor)
        {
            this.OpenConnection();
            string consulta = "delete from " + tabla + " where " + clave + "='" + valor + "'";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            command.ExecuteNonQuery();
            try
            {
                this.CloseConnection();
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int borrarregistrio_foraneo(ref string tabla, ref string clave, ref string valor, ref string clave2, ref string valor2)
        {
            this.OpenConnection();
            string consulta = "delete from " + tabla + " where " + clave + "='" + valor + "' and " + clave2 + "='" + valor2 + "'";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            command.ExecuteNonQuery();
            try
            {
                this.CloseConnection();
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int obtenerRegistro(string tabla, string clave, string valor, ref string[] registro)
        {
            this.CloseConnection();
            this.OpenConnection();
            string consulta = "select * from " + tabla + " where " + clave + "='" + valor + "'";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            MySqlDataReader Reader;
            try
            {
                Reader = command.ExecuteReader();
                if (Reader.Read())
                {
                    int cant = 0;
                    for (int i = 0; i < Reader.FieldCount; i++)
                    {
                        registro[i] = Reader.GetValue(i).ToString();
                        cant++;
                        Array.Resize(ref registro, cant + 1);
                    }
                    Array.Resize(ref registro, cant);
                }
                this.CloseConnection();
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public DataTable obtenerColumna(ref string tabla, ref string campo)
        {
            this.OpenConnection();
            string query = "select " + campo + " from " + tabla;
            MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
            DataSet DS = new DataSet();
            adapter.Fill(DS);
            this.CloseConnection();
            return DS.Tables[0];
        }

        public static int claveExistente(ref string tabla, ref string clave, ref string valor)
        {
            MySqlCommand command = new MySqlCommand();
            command.CommandText = "select * from" + tabla + " where " + clave + "=" + valor;
            try
            {
                command.ExecuteNonQuery();
                MySqlDataReader Reader;
                Reader = command.ExecuteReader();
                if (Reader.Read())
                {
                    return (0);
                }
                else
                {
                    return (1);
                };
            }
            catch
            {
                return (2);
            }
        }

        public int modificarRegistro_foraneo(ref string tabla, ref string[] campos, ref string[] valores)
        {
            string palabra = "";
            for (int i = 2; i < valores.Length - 1; i++)
            {
                if (valores[i] != "")
                    palabra += campos[i] + "='" + valores[i] + "',";
            }
            if (valores[valores.Length - 1] != "")
                palabra += campos[campos.Length - 1] + "='" + valores[valores.Length - 1] + "'";
            this.OpenConnection();
            string consulta = "update " + tabla + " set " + palabra + " where " + campos[0] + "=" + valores[0] + " and " + campos[1] + "=" + valores[1];
            MySqlCommand command = new MySqlCommand(consulta, connection);
            try
            {
                command.ExecuteNonQuery();
                this.CloseConnection();
                return (0);
            }
            catch
            {
                return (1);
            }
        }


        public DataTable obtener_tabla_entre(string table, string fields, string campo, string desde, string hasta, string orden)
        {
            string query = "select " + fields + " from " + table + " where " + campo + " between '" + desde + "' and '" + hasta + "' order by " + orden;
            MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
            DataSet DS = new DataSet();
            adapter.Fill(DS);
            return DS.Tables[0];
        }

        public DataTable obtener_tabla_especial(string consulta)
        {
            string query = consulta;
            MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
            DataSet DS = new DataSet();
            adapter.Fill(DS);
            return DS.Tables[0];
        }

        public DataTable obtener_tabla_listado_arepartir(string sigla, string desde, string hasta, string cliente)
        {
            string query = "SELECT facturas" + sigla + ".id,facturas" + sigla + ".tipo,facturas" + sigla + ".documento,facturas" + sigla + ".cliente,facturas" + sigla + ".razon,facturas" + sigla + ".total,Date_format(facturas" + sigla + ".fecha,'%Y-%m-%d'),facturas" + sigla + ".pagado FROM facturas" + sigla + ",clientes" + sigla + " WHERE (facturas" + sigla + ".fecha between '" + desde + "' and '" + hasta + "') and (clientes" + sigla + ".codigo='" + cliente + "') and (facturas" + sigla + ".cliente = clientes" + sigla + ".codigo) and (facturas"+sigla+".anulada = 0) order by clientes" + sigla + ".razon asc";
            MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
            DataSet DS = new DataSet();
            adapter.Fill(DS);
            return DS.Tables[0];
        }

        public DataTable obtener_tabla_segun(string table, string campo, string valor)
        {
            string query = "select * from " + table + " where (" + campo + " = '"+valor+"')";
            MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
            DataSet DS = new DataSet();
            adapter.Fill(DS);
            return DS.Tables[0];
        }




        //done
        /// <summary>
        /// Trae a un DataTable todos los registros de la tabla 
        /// </summary>
        /// <param name="table">Nombre de la tabla</param>
        /// <param name="fields">Campos</param>
        /// <param name="fields">Desde</param>
        /// <param name="fields">Hasta</param>
        /// <returns></returns>
        public DataTable obtener_tabla_segun(string table, string fields, string campo, string valor, string campoorden, string orden)
        {
            string query = "select "+fields+" from " + table + " where " + campo + " = '" + valor + "' order by "+ campoorden+" "+orden;
            MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
            DataSet DS = new DataSet();
            adapter.Fill(DS);
            return DS.Tables[0];
        }

        public DataTable obtener_tabla_entre_orden(string table, string fields, string campo, string desde, string hasta, string orden)
        {
            string query = "select " + fields + " from " + table + " where " + campo + " between '" + desde + "' and '" + hasta + "' order by " + orden + " asc";
            MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
            DataSet DS = new DataSet();
            adapter.Fill(DS);
            return DS.Tables[0];
        }
        public DataTable obtener_tabla_entre_orden_segun(string table, string fields, string campo, string desde, string hasta, string orden, string segun, string valor)
        {
            string query = "select " + fields + " from " + table + " where " + campo + " between '" + desde + "' and '" + hasta + "' where "+segun +"="+valor+" AND order by " + orden + " asc";
            MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
            DataSet DS = new DataSet();
            adapter.Fill(DS);
            return DS.Tables[0];
        }

        public DataTable obtener_tabla(string table, string fields)
        {
            string query = "select " + fields + " from " + table + " order by codigo asc";
            MySqlDataAdapter adapter = new MySqlDataAdapter(query, connection);
            DataSet DS = new DataSet();
            adapter.Fill(DS);
            return DS.Tables[0];
        }


        public static int obtenerCampo(ref string tabla, ref string clave, ref string valor, ref string campo, ref string valorcampo)
        {
            MySqlCommand command = new MySqlCommand();
            command.CommandText = "select " + campo + " from" + tabla + " where " + clave + "=" + valor;
            try
            {
                command.ExecuteNonQuery();
                MySqlDataReader Reader;
                Reader = command.ExecuteReader();
                if (Reader.Read())
                {
                    valorcampo = Reader.GetValue(0).ToString();
                }
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public double obtener_valor_tasa(string sigla, string fecha, string tasa)
        {
            this.OpenConnection();
            string consulta = "select valor from valores_tasaj23 where valores_tasaj23.desde <= '" + fecha + "' and valores_tasaj23.hasta >= '" + fecha + "' and tasa_id='" + tasa + "'";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            MySqlDataReader Reader;
            try
            {
                Reader = command.ExecuteReader();
                string valor = "";
                if (Reader.Read())
                {
                    valor = Reader.GetValue(0).ToString();
                }
                this.CloseConnection();
                return (Convert.ToDouble(valor));
            }
            catch
            {
                return (0);
            }
        }

        public int concepto_a_calcular(string sigla, string emirec, string concepto)
        {
            this.OpenConnection();
            string consulta = "select count(*) from calculos" + sigla + " where emirec_id = '" + emirec + "' and concepto = '" + concepto + "'";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            MySqlDataReader Reader;
            try
            {
                Reader = command.ExecuteReader();
                string valor = "";
                if (Reader.Read())
                {
                    valor = Reader.GetValue(0).ToString();
                }
                this.CloseConnection();
                return (Convert.ToInt16(valor));
            }
            catch
            {
                return (0);
            }
        }



        public int obtenerRegistro_segun(string tabla, string clave, string valor, string clave2, string valor2, ref string[] registro)
        {
            this.OpenConnection();
            string consulta = "select * from " + tabla + " where " + clave + "=" + valor + " and " + clave2 + "=" + valor2;
            MySqlCommand command = new MySqlCommand(consulta, connection);
            MySqlDataReader Reader;
            try
            {
                Reader = command.ExecuteReader();
                if (Reader.Read())
                {
                    int cant = 0;
                    for (int i = 0; i < Reader.FieldCount; i++)
                    {
                        registro[i] = Reader.GetValue(i).ToString();
                        cant++;
                        Array.Resize(ref registro, cant + 1);
                    }
                    Array.Resize(ref registro, cant);
                }
                this.CloseConnection();
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int obtener_campo(ref string table, ref string fields, ref string campo, ref string desde, ref string hasta, ref string[] registro)
        {
            this.OpenConnection();
            string consulta = "select " + fields + " from " + table + " where " + campo + " between '" + desde + "' and '" + hasta + "'";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            MySqlDataReader Reader;
            try
            {
                Reader = command.ExecuteReader();
                if (Reader.Read())
                {
                    int cant = 0;
                    for (int i = 0; i < Reader.FieldCount; i++)
                    {
                        registro[i] = Reader.GetValue(i).ToString();
                        cant++;
                        Array.Resize(ref registro, cant + 1);
                    }
                }
                this.CloseConnection();
                return (0);
            }
            catch
            {
                return (-1);
            }
        }

        public int obtener_fecha_alta_tasa(string sigla, string id, string tasa, ref string[] registro)
        {
            this.OpenConnection();
            string consulta = "select * from valores_tasa" + sigla + " where tasa_id =" + tasa + " and id<>'" + id + "' order by hasta desc limit 1";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            MySqlDataReader Reader;
            try
            {
                Reader = command.ExecuteReader();
                if (Reader.Read())
                {
                    int cant = 0;
                    for (int i = 0; i < Reader.FieldCount; i++)
                    {
                        registro[i] = Reader.GetValue(i).ToString();
                        cant++;
                        Array.Resize(ref registro, cant + 1);
                    }
                    Array.Resize(ref registro, cant);
                }
                this.CloseConnection();
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public int modificarRegistro(ref string tabla, ref string[] campos, ref string[] valores)
        {
            string palabra = "";
            for (int i = 1; i < valores.Length - 1; i++)
            {
                if (valores[i] != "")
                    palabra += campos[i] + "='" + valores[i] + "',";
            }
            if (valores[valores.Length - 1] != "")
                palabra += campos[campos.Length - 1] + "='" + valores[valores.Length - 1] + "'";
            this.OpenConnection();
            string consulta = "update " + tabla + " set " + palabra + " where " + campos[0] + "='" + valores[0] + "'";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            try
            {
                command.ExecuteNonQuery();
                this.CloseConnection();
                return (0);
            }
            catch
            {
                return (1);
            }
        }

        public string fecha_valor_tasa_hasta(string tabla, string clave, string campo, string valor)
        {
            this.OpenConnection();
            string consulta = "select " + campo + " from " + tabla + " where " + clave + " =' " + valor + "' order by " + campo + " DESC";
            MySqlCommand command = new MySqlCommand(consulta, connection);
            MySqlDataReader Reader;
            try
            {
                Reader = command.ExecuteReader();

                if (Reader.Read())
                {
                    return (Reader.GetValue(0).ToString());
                }
                else
                {
                    return (Convert.ToString(1));
                }

            }
            catch
            {
                return (Convert.ToString(-1));
            }
        }

    }
}
