using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Application = Microsoft.Office.Interop.Excel;
namespace Excell_reader
{
    class Logic

    {
        
        public void read(string FirstFile, string TeacherFio, double SemestrInsert)
        {


            Application.Application app;
            Application.Workbooks wrbks = null;
            Application.Workbook wrbk = null;
            Application.Worksheet wrsh;
            
            double workTime = 0;


            app = new Application.Application { DisplayAlerts = true };
            try
            {





                wrbks = app.Workbooks;
                wrbk = wrbks.Open(Path.Combine(Environment.CurrentDirectory, FirstFile));

                long ShCount = wrbk.Worksheets.Count;
                long SheetCount = 1;
                long NumSecond = 5;
                bool cil = true;



                do
                {


                    wrsh = wrbk.Worksheets[SheetCount++];
                    do
                    {
                        NumSecond++;        //листаем строки
                        string TeacherExcell = wrsh.Range["P" + NumSecond].Value;
                        if (TeacherExcell != null)
                        {
                            TeacherExcell = TeacherExcell.Remove(0, 1);
                            TeacherExcell = TeacherExcell.Replace(" ", "").Replace(".", "");
                            TeacherFio = TeacherFio.Replace(" ", "").Replace(".", "");

                        }

                        if (TeacherFio == TeacherExcell)
                        {
                            // System.Windows.MessageBox.Show("Гыыы");


                            double SemestrTwo = wrsh.Range["F" + NumSecond].Value;
                            if (SemestrInsert == SemestrTwo)
                            {
                                string Discpline1 = wrsh.Range["E" + NumSecond].Value;
                                Discpline1 = Discpline1.Replace(" ", "").Replace("(Экзаменатор)", "");
                                string Discpline2 = wrsh.Range["E" + (NumSecond - 1)].Value;

                                if (Discpline2 != null)
                                {

                                    Discpline2 =Discpline2.Replace(" ", "").Replace("(Экзаменатор)", "");
                                }
                                if ((Discpline1 == Discpline2) || (Discpline2 == null))
                                {
                                    workTime += wrsh.Range["S" + (NumSecond)].Value;
                                    System.Windows.MessageBox.Show("+");
                                }
                                
                                
                                //cil = false;
                                // System.Windows.MessageBox.Show("Гыыы");
                            }

                        }
                        //else NumSecond += 1;
                    } while (cil);
                    Marshal.ReleaseComObject(wrsh);
                } while (SheetCount <= ShCount);




            }
            catch(Exception E)
            {
                System.Windows.MessageBox.Show("Eror"+E.Message);


            }
            finally
            {
                app.Workbooks.Close();
                app.Quit();
                
                
                Marshal.ReleaseComObject(wrbk);
                Marshal.ReleaseComObject(wrbks);

                System.Windows.MessageBox.Show(workTime.ToString());


            }
        }
    }
}
