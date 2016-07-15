using System;
using System.Collections.Generic;
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

        /* Возможно пригодится для создания таблицы из C#
            CREATE TABLE geotable ( id int4, area varchar(200), echelon varchar(200), zone varchar(200), note varchar(200) );  
            SELECT AddGeometryColumn('', 'geotable','geom',-1,'POLYGON',2); 
         */
        public void fillGeoTable()
        {
            string coordinatesbuf;
            string coordinatesbuff;
            try
            {
                connection.Open();
                command.Connection = connection;
                String sqlcom = "SELECT*FROM rawdata2;";
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
                    MessageBox.Show(coordinatesbuff);
                    command.CommandText = "insert into geotable (id, area, echelon, zone, note, geom) values (" + System.Int16.Parse(id.ToString()) + ",'" + area.ToString() + "','" +
                                           echelon.ToString() + "','" + zone.ToString() + "','" + note.ToString() + "', ST_GeomFromText('POLYGON((0 0,0 1,1 1,1 0,0 0))') );";
                    command.ExecuteNonQuery();
                }
            }
            finally
            {
                connection.Close();                  
            }           
        }
    }
}
