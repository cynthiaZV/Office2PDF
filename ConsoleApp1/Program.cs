using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace MSPDF {
	class Program {
		
		static void TestMS() {
			String inputDirectory = ".\\input";
			String outputDirectory = ".\\output";

			String inputWordFile = Path.Combine(inputDirectory, "input.docx");
			String inputPdfFile = Path.Combine(inputDirectory, "input.pdf");
			String inputExcellFile = Path.Combine(inputDirectory, "input.xlsx");
			String inputPowerPointFile = Path.Combine(inputDirectory, "input.pptx");

			String outputWordToPdfFile = Path.Combine(outputDirectory, "docx.pdf"); 
			String outputPdfToWordFile = Path.Combine(outputDirectory, "pdf.docx");
			String outputExcelToPdfFile = Path.Combine(outputDirectory, "xlsx.pdf");
			String outputPowerPointToPdfFile = Path.Combine(outputDirectory, "ppt.pdf");



			MSWord word = new MSWord(inputWordFile, outputWordToPdfFile, true);
			word.Convert();
			word.Close();

			word = new MSWord(inputPdfFile, outputPdfToWordFile, true);
			word.Convert();
			word.Close();

			MSPowerPoint pp = new MSPowerPoint(inputPowerPointFile, outputPowerPointToPdfFile, true);
			pp.Convert();
			pp.Close();
			
			MSExcel excel = new MSExcel(inputExcellFile, outputExcelToPdfFile, true);
			excel.Convert();
			excel.Close();

			return;
		}
		static void Main(string[] args) {


			if (args.Length < 2) {
				Console.WriteLine("Missing arguments. Number of inputs: " + args.Length);
				return;
			}
			String inputFile = args[0];
			String outputFile = args[1];
			
			Console.WriteLine("Input: " + inputFile);
			Console.WriteLine("Output: " + outputFile);

			IMSConvert app = null;
			if (inputFile.Contains(".xls")) {
				app = new MSExcel(inputFile, outputFile, true);
			} else if (inputFile.Contains(".ppt")) {
				app = new MSPowerPoint(inputFile, outputFile, true);
			} else if (inputFile.Contains(".doc")) {
				app = new MSWord(inputFile, outputFile, true);
			} else if (inputFile.Contains(".pdf") && outputFile.Contains(".doc")) {
				app = new MSWord(inputFile, outputFile, true);
			}

			app.Convert();
			app.Close();
			GC.Collect();
			Console.WriteLine("Ending programme.");
		}
	}
}
