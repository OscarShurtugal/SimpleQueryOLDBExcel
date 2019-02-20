using System;
using System.Collections.Generic;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SimpleQueryOLDBExcel
{
    public class MethodsSimpleQueryOLDBExcel
    {

        /// <summary>
        /// This method will perform a Query to the Excel File to certain columns given as parameters.
        /// It will receive both the column to perform the where operation and the column where the result is as well as value you're looking for
        /// </summary>
        /// <param name="excelSheetName"></param>
        /// <param name="filePath"></param>
        /// <param name="columnaASeleccionar"></param>
        /// <param name="columnaDondeBuscar"></param>
        /// <param name="valorBuscado"></param>
        /// <returns></returns>
        public string oledbExcelColumnQuery(string excelSheetName, string filePath, string columnWithTheResult, string columnToPerformWhereOperation, string valueLookedFor)
        {

            String excelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0;HDR=YES\"", filePath);

            //Create Connection to Excel workbook
            try
            {
                using (OleDbConnection excelConnection = new OleDbConnection(excelConnString))
                {
                    string query = "Select ["+columnWithTheResult+"] from ["+excelSheetName+"$] Where ["+columnToPerformWhereOperation+"] ='"+valueLookedFor+"'";

                    //Create OleDbCommand to fetch data from Excel 
                    //using (OleDbCommand cmd = new OleDbCommand("Select [First Name] from ["+excelSheetName+"$] Where [Last Name]='Fontana'", excelConnection))
                    using (OleDbCommand cmd = new OleDbCommand(query, excelConnection))
                    {
                        excelConnection.Open();
                        using (OleDbDataReader dReader = cmd.ExecuteReader())
                        {
                            if (dReader.Read())
                            {
                                return dReader.GetValue(0).ToString();

                            }
                            else
                            {
                                return "NO VALUE FOUND";
                            }
                        }

                    }
                }
            }

            catch (Exception e)
            {
                return e.Message;
            }

        }
    }
}
