using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Application = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Excel_reader
{
    internal class WorkHours 
    { 
        public double budget;
		public double nonBudget;

        public double TotalHours
		{
			get { return budget + nonBudget; }
		}
    }

    internal class Logic
    {
        //Константы для работы с таблицой
        private const string IdColumn = "A";
        private const string DisciplineColumn = "E";
        private const string SemesterColumn = "F";
        private const string TeacherColumn = "P";

        private const string TotalBudgetHoursColumn = "V";
        private const string TotalNonBudgetHoursColumn = "W";
        private const string GroupColumIndex = "G";


        private const long StartingRow = 6;

        private Application.Worksheet sheet;
        private Application.Workbook wrbk;

        private string selectedTeacher;
        private int selectedSemester;
        private double totalHours = 0;

        //    private Dictionary<string, WorkHours> hoursPerDiscipline = new Dictionary<string, WorkHours>();
        private Dictionary<string, Dictionary<string, WorkHours>> hoursPerDiscipline = new Dictionary<string, Dictionary<string, WorkHours>>();
        public void Read(string FirstFile, string TeacherFio, int SemestrInsert)
        {
            Stopwatch stopwatch = Stopwatch.StartNew();

            selectedTeacher = TeacherFio.Replace(" ", "").Replace(".", "");
            selectedSemester = SemestrInsert;

            Application.Application app = new Application.Application { DisplayAlerts = true };

            Application.Workbooks wrbks = app.Workbooks;
            wrbk = app.Workbooks.Open(Path.Combine(Environment.CurrentDirectory, FirstFile));

            try
            {
                for (int sheetIndex = 1; sheetIndex <= wrbk.Worksheets.Count; sheetIndex++)
                {
                    ProcessSheet(sheetIndex);
                }
            }
            catch (Exception E)
            {
                System.Windows.MessageBox.Show("Error: " + E.Message + ". Elapsed time: " + stopwatch.Elapsed);
            }
            finally
            {
                app.Workbooks.Close();
                app.Quit();

                Marshal.ReleaseComObject(wrbk);
                Marshal.ReleaseComObject(wrbks);

                string output = "";

                /* foreach (var disciplineName in hoursPerDiscipline.Keys)
                 {
                     foreach (var groupName in hoursPerDiscipline[disciplineName].Keys)
                     {
                         output += groupName + " Часов по предмету " + disciplineName +
                             ": Бюджет: " + hoursPerDiscipline[disciplineName][groupName].budget +
                             ", Не бюджет: " + hoursPerDiscipline[disciplineName][groupName].nonBudget +
                             ", Всего: " + hoursPerDiscipline[disciplineName][groupName].TotalHours + "\n";
                     }
                 }*/
                FillTemplate();
                //System.Windows.MessageBox.Show(output + "\nElapsed time: " + stopwatch.Elapsed + "\nDiscipline count = " + hoursPerDiscipline.Count);
            }
        }

        private void ProcessSheet(long index)
        {
            sheet = wrbk.Worksheets[index];
            long RowIndex = StartingRow;

            while (GetCell(IdColumn, RowIndex).Value != null)
            {
                ProcessRow(RowIndex);
                RowIndex++;
            }

            Marshal.ReleaseComObject(sheet);
        }

        private void ProcessRow(long index)
        {
            double Semester = GetCell(SemesterColumn, index).Value;
            if (selectedSemester != Semester) return;

            string teacher = GetCell(TeacherColumn, index).Value;
            if (teacher == null) return;

            teacher = teacher.Remove(0, 1).Replace(" ", "").Replace(".", "");

            if (selectedTeacher == teacher)
            {
                AddHoursAtRow(index);
            }
        }

        private void AddHoursAtRow(long index)
        {
            string discipline = GetCell(DisciplineColumn, index).Value;
            discipline = discipline.Replace(" ", "").Replace("(Экзаменатор)", "");
            string groupname = GetCell(GroupColumIndex, index).Value;

            if (!hoursPerDiscipline.ContainsKey(discipline)) // есть ли дисциплина в словаре
            {
                hoursPerDiscipline.Add(discipline, new Dictionary<string, WorkHours>() { { groupname, new WorkHours() } }); // если нету то добавляем с значением часов равным 0 


            }
            if (!hoursPerDiscipline[discipline].ContainsKey(groupname))
            {
                hoursPerDiscipline[discipline].Add(groupname, new WorkHours());
                //  if (hoursPerDiscipline[discipline].ContainsKey(groupname)) { System.Windows.MessageBox.Show("Я РАБОТАЮ"); }
            }

            WorkHours hours = hoursPerDiscipline[discipline][groupname];

            foreach (var gropName in hoursPerDiscipline[discipline].Keys) // вероятно здесь проблема
            {


                var budgetValue = GetCell(TotalBudgetHoursColumn, index).Value;
                var nonBudgetValue = GetCell(TotalNonBudgetHoursColumn, index).Value;

                if (budgetValue != null)
                {
                    hours.budget += budgetValue;
                }

                if (nonBudgetValue != null)
                {
                    hours.nonBudget += nonBudgetValue;
                }

            }


        }
        private void FillTemplate()
        {
            Application.Application exl = new Application.Application();

            if (exl == null)
            {
                System.Windows.MessageBox.Show("Проверьте инсталляцию MS Excel");
                return;
            }

            Application.Worksheet xlSheet;
            object misValue = System.Reflection.Missing.Value;

            Application.Workbook xlBook = exl.Workbooks.Add(misValue);
            xlSheet = (Application.Worksheet)xlBook.Sheets[1];
            xlSheet.Cells[1, 1] = "Предмет";
            xlSheet.Cells[1, 2] = "Группа";
            xlSheet.Cells[1, 3] = "Буджет";
            xlSheet.Cells[1, 4] = "ВнеБ";
            xlSheet.Cells[1, 5] = "Всего";

            int i = 2;
            foreach (var disciplineName in hoursPerDiscipline.Keys) 
            { foreach (var groupname in hoursPerDiscipline[disciplineName].Keys)
                {
                    xlSheet.Cells[i, 1] = disciplineName; //"Часов по предмету " + disciplineName +
                    xlSheet.Cells[i, 2] = groupname;
                    xlSheet.Cells[i, 3] = hoursPerDiscipline[disciplineName][groupname].budget;// ": Бюджет: " + hoursPerDiscipline[disciplineName].budget + 
                    xlSheet.Cells[i, 4] = hoursPerDiscipline[disciplineName][groupname].nonBudget; //", Не бюджет: " + hoursPerDiscipline[disciplineName].nonBudget + 
                    xlSheet.Cells[i, 5] = hoursPerDiscipline[disciplineName][groupname].TotalHours; // ", Всего: " + hoursPerDiscipline[disciplineName].TotalHours + "\n"; 
                    i++;
                }
            } 

            xlBook.SaveAs("e:\\test.xls", Application.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Application.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlBook.Close(true, misValue, misValue);
            exl.Quit();

            Marshal.ReleaseComObject(xlSheet);
            Marshal.ReleaseComObject(xlBook);
            Marshal.ReleaseComObject(exl);
        }

        
		private Range GetCell(string row, long column)
		{
			return sheet.Range[String.Format("{0}{1}", row, column)];
		}
	}
        
}