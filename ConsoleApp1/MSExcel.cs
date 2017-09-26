using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace MSPDF {
	class MSExcel : IMSConvert {
		/// <summary>
		/// This field is not exposed by MS as Excel to PDF conversion takes significant time. If you are
		/// to use interops to convert Excel to PDF expect the process to be CPU intensive and time consuming.
		/// </summary>
		public static readonly int FILEFORMAT_EXCEL_TO_PDF = 57;

		private Boolean overrideExisting;
		private String inputFile;
		private String outputFile;
		private Excel.Application app = null;
		private Boolean convertToPDF = true;
		private Boolean canConvert = false;
		public MSExcel(String inputFile, String outputFile, Boolean overrideExisting) {
			this.inputFile = inputFile;
			this.outputFile = outputFile;
			FileInfo inputInfo = new FileInfo(this.inputFile);
			FileInfo outputInfo = new FileInfo(this.outputFile);
			if (inputInfo.Extension.Equals(".PDF", StringComparison.CurrentCultureIgnoreCase) ^ outputInfo.Extension.Equals(".PDF", StringComparison.CurrentCultureIgnoreCase)) {
				// Input and output format can not be the same
				this.canConvert = true;
				if (inputInfo.Extension.Equals(".PDF", StringComparison.CurrentCultureIgnoreCase)) {
					this.convertToPDF = false;
				} 
			}
			


			this.overrideExisting = overrideExisting;

			// Clean output if they exist
			if (overrideExisting && File.Exists(outputFile)) File.Delete(outputFile);

			//Throw Error if input files does not exist
			if (!File.Exists(inputFile)) throw new IOException("Input file does not exist");

			app = new Excel.Application();
		}
		/// <summary>
		/// Only support Excel to PDF
		/// </summary>
		/// <returns></returns>
		public bool Convert() {
			if (!this.canConvert) {
				// Unsupported file types. One of file types has to be PDF
				return false;
			}
			if (this.convertToPDF) {
				return this.ConvertToPDF();
			}
			return false;
		}

		private bool ConvertToPDF() {
			Excel.Workbook xcel = null;
			try {
				app = new Excel.Application();

				xcel = app.Workbooks.Open(this.inputFile);
				xcel.SaveAs(this.outputFile, FILEFORMAT_EXCEL_TO_PDF);
			} catch (Exception ex) {
				Console.WriteLine(ex.Message + "" + ex.StackTrace);
			}
			try {
				// Exit without prompting the save dialog
				xcel.Close(Excel.XlSaveAction.xlDoNotSaveChanges);
			} catch (Exception ex) {
				Console.WriteLine(ex.Message + "" + ex.StackTrace);
			}
			

			if (File.Exists(this.outputFile)) {
				return true;
			} else {
				return false;
			}
		}

		public void Close() {
			try {
				// Exit without prompting the save dialog
				app.Quit();
			} catch (Exception ex) {
				throw new ApplicationException("Could not cleanly shutdown MS Word. Error message: " + ex.Message);
			}
		}
	}
}
