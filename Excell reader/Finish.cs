using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel=Microsoft.Office.Interop.Excel;
namespace Excell_reader
{
    class Finish
    {
        private Excel.Application app = null;
        private Excel.Workbook workbook = null;
        private Excel.Worksheet worksheet = null;
        private Excel.Range worksheet_range = null;
        

        public void CreateFinishDoc(double workTime)
        {
            try
            {
                app = new Excel.Application();
                app.Visible = true;
                workbook = app.Workbooks.Add(1);
                worksheet = (Excel.Worksheet)workbook.Sheets[1];

                Excel.Range cells1 = (Excel.Range)worksheet.get_Range("A6", "A17").Cells;
                cells1.Merge(Type.Missing);
                Excel.Borders border = worksheet.Range["A6", "N27"].Borders;
                border.LineStyle = Excel.XlLineStyle.xlContinuous;
                
                

                
               
                border.LineStyle = Excel.XlLineStyle.xlContinuous;


            }
            catch (Exception e)
            {
                Console.Write("Error");
            }
            finally
            {
            }
        }
    }
}
