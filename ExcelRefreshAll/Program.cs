using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace ExcelRefreshAll
{
	class Program
	{
		static void Main(string[] args)
		{
			if (args.Length == 0 || String.IsNullOrEmpty(args[0]))
				throw new Exception("Missing argument specifying folder to process.");

			var folderPath = args[0];
			if(!Directory.Exists(folderPath))
				throw new Exception("Folder does not exist.");

			Console.WriteLine("Opening Excel...");
			var excel = new Application
			{
				ShowWindowsInTaskbar = true
			};

			var folder = Directory.GetFiles(folderPath);
			foreach (var file in folder)
			{
				if (!file.EndsWith(".xlsx")) continue;

				Console.Write("Opening '{0}'...", file);
				var workbook = excel.Workbooks.Open(file);
				Console.WriteLine(" done.");

				var dataConnections = workbook.Connections;
				foreach (WorkbookConnection connection in dataConnections)
				{
					Console.Write("Refreshing data connection '{0}'...", connection.Name);
					try
					{
						var originalBackgroundQuery = connection.OLEDBConnection.BackgroundQuery;
						connection.OLEDBConnection.BackgroundQuery = false;
						connection.Refresh();
						connection.OLEDBConnection.BackgroundQuery = originalBackgroundQuery;
						Console.WriteLine(" done.");
					}
					catch (Exception exception)
					{
						Console.WriteLine(" error ({0}).", exception.Message);
					}
				}

				Console.Write("Saving workbook...");
				workbook.Save();
				Console.WriteLine(" done.");

				Console.Write("Closing workbook...");
				workbook.Close();
				Console.WriteLine(" done.");
			}

			Console.Write("Closing Excel...");
			excel.Quit();
			Console.WriteLine(" done.");
		}
	}
}
