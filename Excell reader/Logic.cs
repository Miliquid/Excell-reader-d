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
        public string group;

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
		//private const string TotalHoursColumn = "S";
        private const string TotalBudgetHoursColumn = "V";
        private const string TotalNonBudgetHoursColumn = "W";


		private const long StartingRow = 6;
       
		private Application.Worksheet sheet;
        private Application.Workbook wrbk;

		private string selectedTeacher;
        private int selectedSemester;
        private double totalHours = 0;

        private Dictionary<string, WorkHours> hoursPerDiscipline = new Dictionary<string, WorkHours>();

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

                foreach(var disciplineName in hoursPerDiscipline.Keys)
				{
					output += "Часов по предмету " + disciplineName + 
                        ": Бюджет: " + hoursPerDiscipline[disciplineName].budget +
                        ", Не бюджет: " + hoursPerDiscipline[disciplineName].nonBudget +
						", Всего: " + hoursPerDiscipline[disciplineName].TotalHours + "\n";
				}
                
                System.Windows.MessageBox.Show(output + "\nElapsed time: " + stopwatch.Elapsed + "\nDiscipline count = " + hoursPerDiscipline.Count);
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

			if (!hoursPerDiscipline.ContainsKey(discipline)) // есть ли дисциплина в словаре
			{
				hoursPerDiscipline.Add(discipline, new WorkHours()); // если нету то добавляем с значением часов равным 0
			}

            WorkHours hours = hoursPerDiscipline[discipline];

            var budgetValue = GetCell(TotalBudgetHoursColumn, index).Value;
            var nonBudgetValue = GetCell(TotalNonBudgetHoursColumn, index).Value;

            if(budgetValue != null)
			{
				hours.budget += budgetValue;
			}
			if(nonBudgetValue != null)
			{
				hours.nonBudget += nonBudgetValue; 
            }

			//hoursPerDiscipline[discipline] += GetCell(TotalHoursColumn, index).Value;
	    }

		private Range GetCell(string row, long column)
		{
			return sheet.Range[String.Format("{0}{1}", row, column)];
		}
        public void FillTemplat(string Template)
        {
            Application.Application app = new Application.Application { DisplayAlerts = true };

            Application.Workbooks wrbks = app.Workbooks;
            wrbk = app.Workbooks.Open(Path.Combine(Environment.CurrentDirectory, Template));
            foreach (var discipline in hoursPerDiscipline.Keys)
            {
              //  sheet.(discipline, hoursPerDiscipline[discipline].budget, hoursPerDiscipline[discipline]nonBbudget);
            }
        }

    }
}