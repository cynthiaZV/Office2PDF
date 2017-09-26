using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
namespace MSPDF {
	class MSPowerPoint : IMSConvert {
		private Boolean overrideExisting;
		private String inputFile;
		private String outputFile;
		private Boolean convertToPDF = true;
		private Boolean canConvert = false;
		private PowerPoint.Application app = null;
		public MSPowerPoint(String inputFile, String outputFile, Boolean overrideExisting) {
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

			app = new PowerPoint.Application();
		}
		public Boolean Convert() {
			if (!this.canConvert) {
				// Unsupported file types. One of file types has to be PDF
				return false;
			}
			if (this.convertToPDF) {
				return this.ConvertToPDF();
			} else {
				return this.ConvertToPP();
			}
		}
		private Boolean ConvertToPDF() {
			PowerPoint.Presentation presentation = null;
			try {
				app = new PowerPoint.Application();

				presentation = app.Presentations.Open(this.inputFile);

				presentation.ExportAsFixedFormat2(this.outputFile, PowerPoint.PpFixedFormatType.ppFixedFormatTypePDF);//SaveCopyAs(outFilePDF , PowerPoint.PpSaveAsFileType.ppSaveAsPNG);

			} catch (Exception ex) {
				Console.WriteLine(ex.Message + "" + ex.StackTrace);
				return false;
			}
			try {
				// Exit without prompting the save dialog
				presentation.Close();
			} catch (Exception ex) {
				Console.WriteLine(ex.Message + "" + ex.StackTrace);
			}
			

			if (File.Exists(this.outputFile)) {
				return true;
			} else {
				return false;
			}
		}
		
		private Boolean ConvertToPP() {
			PowerPoint.Presentation presentation = null;
			try {
				app = new PowerPoint.Application();

				presentation = app.Presentations.Open(this.inputFile);

				presentation.Convert2(this.outputFile);//SaveCopyAs(outFilePDF , PowerPoint.PpSaveAsFileType.ppSaveAsPNG);

			} catch (Exception ex) {
				Console.WriteLine(ex.Message + "" + ex.StackTrace);
				return false;
			}
			try {
				// Exit without prompting the save dialog
				presentation.Close();
			} catch (Exception ex) {
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
				app.Quit();
			} catch (Exception ex) {
				throw new ApplicationException("Could not cleanly shutdown MS Word. Error message: " + ex.Message);
			}
		}
	}
}

