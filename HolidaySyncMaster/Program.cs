using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data.SqlClient;

namespace HolidaySyncMaster
{
    class Program
    {
        static void Main(string[] args)
        {

            string holidaySheet = @"\\designsvr1\apps\Design and Supply MS EXCEL\Capacity Related\ShopFloorHolidayLinkMacro.xlsm";

            //open the sheet


            Process[] processesBefore = Process.GetProcessesByName("excel");
            // Open the file in Excel.
            var xlApp = new Excel.Application();
            var xlWorkbooks = xlApp.Workbooks;
            var xlWorkbook = xlWorkbooks.Open(holidaySheet,3,Type.Missing,Type.Missing,"music");
            var xlWorksheet = xlWorkbook.Sheets[1]; // assume it is the first sheet
                                                    // Get Excel processes after opening the file.


            Process[] processesAfter = Process.GetProcessesByName("excel");
            //force update of links
      

        //




        Excel.Range xlRange;
            xlRange = xlWorksheet.UsedRange;

            List<int> staff_id_list = new List<int>();
            List<int> staff_column_list = new List<int>();
            List<string> date_column = new List<string>();


            int nRows = xlRange.Rows.Count + 1;
            int nCols = xlRange.Columns.Count + 1;


            //get each staff id (and the int of this column) into lists
            for (int column = 3; column < nCols; column++)
            {
                xlRange = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[2, column];
                if (Convert.ToInt32(xlRange.Text) > 0)
                {
                    staff_id_list.Add(Convert.ToInt32(xlRange.Text));
                    staff_column_list.Add(column);
                }
            }

            //get each date and insert it into the date_column list
            for (int row = 3; row <= nRows; row++)
            {
                string temp = "";
                for (int column = 1; column < 3; column++)
                {
                    xlRange = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[row, column];
                    temp = temp + xlRange.Text;
                    if (column == 1)
                        temp = temp + "_";

                }
                date_column.Add(temp);
            }


            //make a datatable and store everything in it 
            //staff id / date 1 / date 2 / date 3   etc

            DataTable dt = new DataTable();
            DataColumn staff_column_insert = dt.Columns.Add("staff_id", typeof(Int32)); //always the first column
            // add a column for each unique date (in date_column)
            for (int i = 0; i < date_column.Count; i++)
                dt.Columns.Add(date_column[i].ToString(), typeof(float));
            //add the max number of rows we will need here
            for (int staff_loop = 0; staff_loop < staff_column_list.Count; staff_loop++)
            {
                DataRow newRow;
                newRow = dt.NewRow();
                dt.Rows.Add(newRow);
            }

            //start adding the data for each staff id




            for (int dtRow = 0; dtRow < dt.Rows.Count; dtRow++)//row of each staffs data
            {

                for (int dtCol = 0; dtCol < dt.Columns.Count; dtCol++) // staff id > aug value > may value
                {
                    //loop throug excel rows and get each value for the staff
                    Excel.Range range;
                    range = xlWorksheet.UsedRange;
                    range = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[dtCol + 2, staff_column_list[dtRow]];
                    dt.Rows[dtRow][dtCol] = Convert.ToDouble(range.Text);


                    Console.WriteLine(dt.Rows[dtRow][dtCol]);
                }

            }


            using (SqlConnection conn = new SqlConnection(CONNECT.ConnectionString))
            {
                conn.Open();
                string sql = "DROP TABLE [dbo].[aaa_holiday_master]";
                try
                {
                    using (SqlCommand cmd = new SqlCommand(sql, conn))
                        cmd.ExecuteNonQuery();
                }
                catch { }

                sql = "CREATE TABLE [dbo].[aaa_holiday_master](" +
                "[staff_id][int] NULL,";

                //loop through date_column
                for (int date_loop = 0; date_loop < date_column.Count; date_loop++)
                    sql = sql + " [" + date_column[date_loop].ToString() + "] [float] NULL,";

                sql = sql.Remove(sql.Length - 1, 1);
                sql = sql + ") ON[PRIMARY]";

                using (SqlCommand cmd = new SqlCommand(sql, conn))
                    cmd.ExecuteNonQuery();




                //now we add everything into the new table 


                sql = "INSERT INTO [dbo].[aaa_holiday_master] (staff_id,";
                for (int date_loop = 0; date_loop < date_column.Count; date_loop++)
                    sql = sql + date_column[date_loop].ToString() + ",";
                sql = sql.Remove(sql.Length - 1, 1);
                sql = sql + ") VALUES (";

                string insert_sql = "";

                for (int dtRow = 0; dtRow < dt.Rows.Count; dtRow++)//row of each staffs data
                {
                    for (int dtCol = 0; dtCol < dt.Columns.Count; dtCol++) // staff id > aug value > may value                    
                        insert_sql = insert_sql + dt.Rows[dtRow][dtCol].ToString() + ",";
                    insert_sql = insert_sql.Remove(insert_sql.Length - 1, 1);
                    insert_sql =  sql + insert_sql + ")";

                    using (SqlCommand cmd = new SqlCommand(insert_sql, conn))
                    cmd.ExecuteNonQuery();

                    insert_sql = "";
                }

                conn.Close();
            }



            ////////foreach (DataRow row in dt.Rows)
            ////////{
            ////////    foreach (DataColumn col in dt.Columns)
            ////////    {
            ////////        if (col.ColumnName == "staff_id")
            ////////            row[col] = staff_id_list[row.in]
            ////////        else
            ////////        {
            ////////            for (int holiday_row = 3; holiday_row <= nRows; holiday_row++)
            ////////            {
            ////////                xlRange = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[holiday_row, staff_column_list[staff_loop]];
            ////////                // xlRange.Text;
            ////////                newRow[dt_col] = Convert.ToDouble(xlRange.Text);
            ////////            }
            ////////        }

            ////////    }
            ////////}






            //for (int row = 3; row <= nRows; row++)
            //{
            //    for (int col = 0; col < staff_column.Count; col++)
            //    {
            //        xlRange = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[3, staff_column[col]];
            //        Console.WriteLine(xlRange.Text);
            //    }
            //}



            //for (int iRow = 1; iRow <= nRows + 1; iRow++)
            //{
            //    for (int iCount = 1; iCount <= nCols; iCount++)
            //    {
            //        xlRange = (Microsoft.Office.Interop.Excel.Range)xlWorksheet.Cells[iRow, iCount];
            //        Console.WriteLine(xlRange.Text);
            //    }
            //}




            //close the sheet and use task manager to close
            //xlWorkbook.Close(false); //close the excel sheet without saving
            //// xlApp.Quit();
            //// Manual disposal because of COM
            //xlApp.Quit();
            // Now find the process id that was created, and store it.
            int processID = 0;
            foreach (Process process in processesAfter)
            {
                if (!processesBefore.Select(p => p.Id).Contains(process.Id))
                {
                    processID = process.Id;
                }
            }

            // And now kill the process.
            if (processID != 0)
            {
                Process process = Process.GetProcessById(processID);
                process.Kill();
            }

        }
    }
}
