using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Application = Microsoft.Office.Interop.Excel;
namespace Excel_reader
{
	class Logic
	{
		private const string DisciplineColumn = "E";
		private const string SemesterColumn = "F";
		private const string TeacherColumn = "P";
		private const string TotalHoursColumn = "S";
		private const long StartingRow = 6;

		private Application.Application app;
		private Application.Workbooks wrbks;
		private Application.Workbook wrbk;
		private Application.Worksheet wrsh;

		public void read(string FirstFile, string TeacherFio, double SemestrInsert)
		{
			double totalHours = 0;

			app = new Application.Application { DisplayAlerts = true };
			try
			{
				wrbks = app.Workbooks;
				wrbk = wrbks.Open(Path.Combine(Environment.CurrentDirectory, FirstFile));

				long SheetCount = wrbk.Worksheets.Count;
				long CurrentSheetIndex = 1;
				long RowIndex = StartingRow;

				bool finished = false;

				do
				{
					wrsh = wrbk.Worksheets[CurrentSheetIndex++];

					while (!finished)
					{
						object end = wrsh.Range[String.Format("A{0}", RowIndex)].Value;
						if (end == null)
							finished = true;

						string TeacherExcell = wrsh.Range[String.Format("{0}{1}", TeacherColumn, RowIndex)].Value;
						if (TeacherExcell != null)
						{
							TeacherExcell = TeacherExcell.Remove(0, 1);
							TeacherExcell = TeacherExcell.Replace(" ", "").Replace(".", "");
							TeacherFio = TeacherFio.Replace(" ", "").Replace(".", "");

						}

						if (TeacherFio == TeacherExcell)
						{
							double SemestrTwo = wrsh.Range[String.Format("{0}{1}", SemesterColumn, RowIndex)].Value;
							if (SemestrInsert == SemestrTwo)
							{
								string Discpline1 = wrsh.Range[String.Format("{0}{1}", DisciplineColumn, RowIndex)].Value;
								Discpline1 = Discpline1.Replace(" ", "").Replace("(Экзаменатор)", "");
								string Discpline2 = wrsh.Range[String.Format("{0}{1}", DisciplineColumn, RowIndex - 1)].Value;

								if (Discpline2 != null)
								{

									Discpline2 = Discpline2.Replace(" ", "").Replace("(Экзаменатор)", "");
								}
								if ((Discpline1 == Discpline2) || (Discpline2 == null))
								{
									totalHours += wrsh.Range[String.Format("{0}{1}", TotalHoursColumn, RowIndex)].Value;
									//System.Windows.MessageBox.Show("+");
								}

								//   double end = wrsh.Range[String.Format("A{0}", RowIndex)].Value;
								// if (end == null)
								//     Finish = false;

							}
						}
						RowIndex++;

					}
					Marshal.ReleaseComObject(wrsh);

				} while (CurrentSheetIndex <= SheetCount);
			}
			catch (Exception E)
			{
				System.Windows.MessageBox.Show("Eror" + E.Message);
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
    }
}
