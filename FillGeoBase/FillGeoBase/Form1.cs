﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Npgsql;
using System.Reflection;
using WorkBook = Microsoft.Office.Interop.Excel.Workbook;
using Word = Microsoft.Office.Interop.Word;
using Application = Microsoft.Office.Interop.Excel.Application;


using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace FillGeoBase
{
    public partial class Form1 : Form
    {
        public string Pathfile { get; set; }
        public Form1()
        {
            InitializeComponent();
        }

        private static object[,] loadCellByCell(int row, int maxColNum, _Worksheet osheet)
        {
            var list = new object[2, maxColNum + 1];
            for (int i = 1; i <= maxColNum; i++)
            {
                var RealExcelRangeLoc = osheet.Range[(object)osheet.Cells[row, i], (object)osheet.Cells[row, i]];
                object valarrCheck;
                try
                {
                    valarrCheck = RealExcelRangeLoc.Value[XlRangeValueDataType.xlRangeValueDefault];
                }
                catch
                {
                    valarrCheck = (object)RealExcelRangeLoc.Value2;
                }
                list[1, i] = valarrCheck;
            }
            return list;
        }

        private static void fillGeoTable(NpgsqlConnection conn)
        {
            NpgsqlCommand command = new NpgsqlCommand();
            command.Connection = conn;
            //String sqlcomfill = "insert into test (id) values (1);";
            String sqlcom = "SELECT*FROM rawdata;";
            System.Data.DataTable dt = new System.Data.DataTable();
            NpgsqlDataAdapter da = new NpgsqlDataAdapter(sqlcom, conn);
            da.Fill(dt);
            System.Data.DataTableReader tablereader = dt.CreateDataReader();
            while (tablereader.Read())
            {
                Object id = tablereader.GetValue(0); ;
                Console.WriteLine(System.Int16.Parse(id.ToString()));
                command.CommandText = "insert into test (id) values (" +System.Int16.Parse(id.ToString())+ ");";
                command.ExecuteNonQuery();
                Console.ReadLine();
            }

        }

        private void button1_Click(object sender, EventArgs e)
        {  

            using (OpenFileDialog dialog = new OpenFileDialog())
            {
                //dialog.Filter = "Текстовые файлы|*.txt";
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    //textBox1.Text = File.ReadAllText(dialog.FileName);
                    Pathfile = dialog.FileName;
                    textBox1.Text = Pathfile;
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {      

            string connectionString = "Server=127.0.0.1;Port=5432;User Id=postgres;Password=supervisor;Database=gisdb;";
            NpgsqlConnection conn = new NpgsqlConnection(connectionString);
            conn.Open();
            NpgsqlCommand comm = new NpgsqlCommand();
            comm.Connection = conn;

            fillGeoTable(conn);
            
            Application ExcelObj = null;
            WorkBook excelbook = null;
            try{
                //Word.Application application = new Word.Application();
                //Word.Document document;
                ExcelObj = new Application();
                ExcelObj.DisplayAlerts = false;
                /*const*/ string f = Pathfile;//@"C:\book.xlsx";
                excelbook = ExcelObj.Workbooks.Open(f, 0, true, 5, "", "", false, XlPlatform.xlWindows);

                var sheets = excelbook.Sheets;
                var maxNumSheet = sheets.Count;

                for (int i = 1; i <= maxNumSheet; i++)
                {
                    var osheet = (_Worksheet) excelbook.Sheets[i];
                    Range excelRange = osheet.UsedRange;
                    int maxColNum;
                    int lastRow;
                    try
                    {
                        maxColNum = excelRange.SpecialCells(XlCellType.xlCellTypeLastCell).Column;
                        lastRow = excelRange.SpecialCells(XlCellType.xlCellTypeLastCell).Row;
                    }
                    catch
                    {
                        maxColNum = excelRange.Columns.Count;
                        lastRow = excelRange.Rows.Count;
                    }

                    for (int l = 1; l <= lastRow; l++)
                    {
                        Range RealExcelRangeLoc = osheet.Range[(object) osheet.Cells[l, 1], (object) osheet.Cells[l, maxColNum]];
                        object[,] valarr = null;
                        try
                        {
                            var valarrCheck = RealExcelRangeLoc.Value[XlRangeValueDataType.xlRangeValueDefault];
                            if (valarrCheck is object[,] || valarrCheck == null)
                                valarr = (object[,]) RealExcelRangeLoc.Value[XlRangeValueDataType.xlRangeValueDefault];

                         // Console.WriteLine(valarr[1, 1] + " " + valarr[1, 3]);

                            string sql = "insert into rawdata (id, polygon, clss) values ('" + valarr[1, 1] + "', '" + valarr[1, 3] + "', '" + valarr[1, 5] + "');";
                            comm.CommandText = sql;
                            comm.ExecuteNonQuery();//.ExecuteScalar().ToString(); //Выполняем нашу команду.
                            comm.Dispose();    
                        }
                        catch
                        {
                            valarr = loadCellByCell(l, maxColNum, osheet);
                        }
                    }
                }
            }
            finally
            {
                conn.Close();
                if (excelbook != null)
                {
                    excelbook.Close();
                    Marshal.ReleaseComObject(excelbook);
                }
                if (ExcelObj != null) ExcelObj.Quit();
            }
        }
    }   
}