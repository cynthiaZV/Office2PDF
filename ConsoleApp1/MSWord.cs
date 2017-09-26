using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace MSPDF {
	class MSWord : IMSConvert {
		private Boolean overrideExisting;
		private String inputFile;
		private String outputFile;
		private Boolean convertToPDF = true;
		private Boolean canConvert = false;
		private Word.Application app = null;
		public MSWord(String inputFile, String outputFile, Boolean overrideExisting) {
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

			app = new Word.Application();
		}

		private Boolean ConvertToPDF() {
			
			Word.Document doc = null;
			try {
				doc = app.Documents.Open(this.inputFile);
				doc.SaveAs2(this.outputFile, Word.WdSaveFormat.wdFormatPDF);
			} catch (Exception ex) {
				Console.WriteLine(ex.Message + "" + ex.StackTrace);
			}
			try {
				// Exit without prompting the save dialog
				doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
			} catch (Exception ex) {
				Console.WriteLine(ex.Message + "" + ex.StackTrace);
			}

			if (File.Exists(this.outputFile)) {
				return true;
			} else {
				return false;
			}
		}
		public Boolean Convert() {
			if (!this.canConvert) {
				// Unsupported file types. One of file types has to be PDF
				return false;
			}
			if (this.convertToPDF) {
				return this.ConvertToPDF();
			} else {
				return this.ConvertToWord();
			}
		}
		private Boolean ConvertToWord() {
			Word.Document doc = null;
			try {
				app = new Word.Application();
				app.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
				doc = app.Documents.Open(this.inputFile);
				doc.SaveAs2(this.outputFile, Word.WdSaveFormat.wdFormatDocumentDefault);

			} catch (Exception ex) {
				Console.WriteLine(ex.Message + "" + ex.StackTrace);
				return false;
			}

			try {
				// Exit without prompting the save dialog
				doc.Close(Word.WdSaveOptions.wdDoNotSaveChanges);
			} catch (Exception ex) {
				// This is not fatal. MS Word just want us to press save changes before.
				Console.WriteLine(ex.Message + "" + ex.StackTrace);
			}
			
			if (File.Exists(this.outputFile)) {
				return true;
			} else {
				return false;
			}

		}
		/// <summary>
		///		Close MS Application. Throws ApplicationException upon failure.  
		/// </summary>
		public void Close() {
			try {
				// Exit without prompting the save dialog
				if(app != null) {
					app.Quit(Word.WdSaveOptions.wdDoNotSaveChanges);
				}
			} catch (Exception ex) {
				throw new ApplicationException("Could not cleanly shutdown MS Word. Error message: " + ex.Message);
			}
		}
	}
}
