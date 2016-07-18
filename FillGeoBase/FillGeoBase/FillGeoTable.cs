using System;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Npgsql;
using System.Windows.Forms;

namespace FillGeoBase
{
    class FillGeoTable
    {
        public NpgsqlConnection connection {get; set;}

        public FillGeoTable(string constring)
        {
            this.connection = new NpgsqlConnection(constring);
        }

        NpgsqlCommand command = new NpgsqlCommand();

        public void creatRawDataTable(string tablename)  //Создание таблицы с "сырыми" данными из Excel
        {
            try
            {
                connection.Open();
                command.Connection = connection;
                String createtable = "CREATE TABLE " + tablename + " ( id varchar, area varchar, coordinates varchar, echelon varchar, zone varchar, note varchar);";
                new NpgsqlCommand(createtable, connection).ExecuteNonQuery();
            }
            catch
            {
                MessageBox.Show("Таблица с таким именем не была создана, возможно она уже существует");
            }
            finally
            {
                connection.Close();
            }

        }

        public void creatGeoDataTable(string tablename)  //Создание таблицы с геоданными из таблицы с "сырыми данными"
        {
            try
            {
                connection.Open();
                command.Connection = connection;
                String createtable = "CREATE TABLE geo" + tablename + " ( id int4, area varchar(200), echelon varchar(200), zone varchar(200), note varchar(200) );" +
                                     "SELECT AddGeometryColumn('', 'geo"+tablename+"','geom',-1,'POLYGON',2)";
                new NpgsqlCommand(createtable, connection).ExecuteNonQuery();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message.ToString());
            }
            finally
            {
                connection.Close();
            }

        }

        private void clearTable(string tablename)    // метод очистки выбранной таблицы
        {
            command.Connection = connection;
            String tabletrunc = "TRUNCATE "+tablename+";";
            new NpgsqlCommand(tabletrunc, connection).ExecuteNonQuery();
        }

        public void fillGeoTable(string tablename)  // метод заполнения геотаблицы геоданными из таблицы с "сырыми" данными
        {
            string coordinatesbuf;
            string coordinatesbuff;
            try
            {
                connection.Open();
                clearTable("geo"+tablename);

                String sqlcom = "SELECT*FROM "+tablename+";";
                DataTable dt = new DataTable();
                NpgsqlDataAdapter da = new NpgsqlDataAdapter(sqlcom, connection);
                da.Fill(dt);
                System.Data.DataTableReader tablereader = dt.CreateDataReader();
                while (tablereader.Read())
                {
                    Object id = tablereader.GetValue(0);
                    Object area = tablereader.GetValue(1);
                    Object coordinates = tablereader.GetValue(2); ;
                    Object echelon = tablereader.GetValue(3);
                    Object zone = tablereader.GetValue(4);
                    Object note = tablereader.GetValue(5);

                    coordinatesbuf = coordinates.ToString();
                    coordinatesbuff = coordinatesbuf.Remove(coordinatesbuf.Length-2);
                    //MessageBox.Show(coordinatesbuff);
                    try
                    {
                        command.CommandText = "insert into geo"+tablename+" (id, area, echelon, zone, note, geom) values (" + System.Int16.Parse(id.ToString()) + ",'" + area.ToString() + "','" +
                                           echelon.ToString() + "','" + zone.ToString() + "','" + note.ToString() + "', ST_GeomFromText('POLYGON((" + coordinatesbuff + "))') );";
                        command.ExecuteNonQuery();
                    }
                    catch (Exception ex)
                    {
                       //MessageBox.Show(ex.Message.ToString());
                    }
                    
                }
            }
            finally
            {
                connection.Close();                  
            }           
        }
    }
}
