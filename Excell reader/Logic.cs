using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Application = Microsoft.Office.Interop.Excel;
using Range = Microsoft.Office.Interop.Excel.Range;

namespace Excel_reader
{
	class Logic
	{
		private const string DisciplineColumn = "E";
		private const string SemesterColumn = "F";
		private const string TeacherColumn = "P";
		private const string TotalHoursColumn = "S";
		private const long StartingRow = 6;

		private Application.Worksheet sheet;
        private Application.Workbook wrbk;

		private string selectedTeacher;
        private int selectedSemester;
        private double totalHours = 0;

		public void Read(string FirstFile, string TeacherFio, int SemestrInsert)
		{
			selectedTeacher = TeacherFio.Replace(" ", "").Replace(".", "");
			selectedSemester = SemestrInsert;

			Application.Application app = new Application.Application { DisplayAlerts = true };
			Application.Workbooks wrbks = app.Workbooks;
			wrbk = app.Workbooks.Open(Path.Combine(Environment.CurrentDirectory, FirstFile));

			try
			{
				for (int sheetIndex = 1; sheetIndex <= wrbk.Worksheets.Count; sheetIndex++)
                {
					sheet = wrbk.Worksheets[sheetIndex];
					long RowIndex = StartingRow;

					while (GetCell("A", RowIndex).Value != null)
					{
						ProcessRow(RowIndex);
						RowIndex++;
					}

					Marshal.ReleaseComObject(sheet);
				}   
			}
			catch (Exception E)
			{
				System.Windows.MessageBox.Show("Error" + E.Message);
			}
			finally
			{
				app.Workbooks.Close();
				app.Quit();

				Marshal.ReleaseComObject(wrbk);
				Marshal.ReleaseComObject(wrbks);

				System.Windows.MessageBox.Show(totalHours.ToString());
			}
        }

        private Range GetCell(string row, long column)
		{
            return sheet.Range[String.Format("{0}{1}", row, column)];
		}

        private void ProcessRow(long index)
		{
			string teacher = GetCell(TeacherColumn, index).Value;

			if (teacher == null) return;

			teacher = teacher.Remove(0, 1).Replace(" ", "").Replace(".", "");

			if (selectedTeacher == teacher)
			{
				double SemestrTwo = GetCell(SemesterColumn, index).Value;
				if (selectedSemester == SemestrTwo)
				{
					string Discpline1 =  GetCell(DisciplineColumn, index).Value;
					Discpline1 = Discpline1.Replace(" ", "").Replace("(Экзаменатор)", "");
					string Discpline2 = GetCell(DisciplineColumn, index - 1).Value;

					if (Discpline2 != null)
					{
						Discpline2 = Discpline2.Replace(" ", "").Replace("(Экзаменатор)", "");
					}
					if ((Discpline1 == Discpline2) || (Discpline2 == null))
					{
						totalHours += GetCell(TotalHoursColumn, index).Value;
					}
				}
			}
		}

        private void ProcessSheet(long index)
		{

		}
	}
}