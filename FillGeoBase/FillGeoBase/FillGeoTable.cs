using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using Npgsql;

namespace FillGeoBase
{
    class FillGeoTable
    {
        public NpgsqlConnection connection {get; set;}

        public FillGeoTable(string constring)
        {
           // this.constring = constring;
            this.connection = new NpgsqlConnection(constring);
        }

        //NpgsqlConnection connection = new NpgsqlConnection(constring);
        NpgsqlCommand command = new NpgsqlCommand();

        public void fillGeoTable(/*NpgsqlConnection conn*/)
        {
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
                    command.CommandText = "insert into geotable (id, area, echelon, zone, note, geom) values (" + System.Int16.Parse(id.ToString()) + ",'" + area.ToString() + "','" +
                                           echelon.ToString() + "','" + zone.ToString() + "','" + note.ToString() + "', ST_GeomFromText('POLYGON((0 0,0 1,1 1,1 0,0 0))') );";
                    //command.CommandText = "insert into geotable (id, area, echelon) values (" + System.Int16.Parse(id.ToString()) + ",'" + area.ToString() + "','" +
                    //                       echelon.ToString() + "');";
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
