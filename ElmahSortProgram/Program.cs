using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Xml;
using System.Xml.Linq;
using System.Reflection;
using Microsoft.Office.Interop.Excel;
using GemBox.Spreadsheet;



namespace ExcelSortingAutomation
{

	public class Program
	{
		[STAThread]
		public static void Main(string[] args)
		{

			SpreadsheetInfo.SetLicense("EL6O-2685-AZHM-Z6GQ");
			ExcelFile ef = new ExcelFile();
			ExcelWorksheet ws = ef.Worksheets.Add("Hello World");

			DirectoryInfo Exceptions = new DirectoryInfo("C:\\ErrorsMay2017");
			FileInfo[] ExceptionFiles = Exceptions.GetFiles("*.xml");

			var checkTime = DateTime.Now;
			int i = 0;
			
			// generates table head 
			ws.Cells[0, 0].Value = "Error Id:";
			ws.Cells[0, 1].Value = "type";
			ws.Cells[0, 2].Value = "message";
			ws.Cells[0, 3].Value = "time";
			ws.Cells[0, 4].Value = "Http_Host";
			ws.Cells[0, 5].Value = "Path_Info";
			string host = "";
			foreach (var exception in ExceptionFiles.OrderByDescending(x => x.CreationTime).ToList())
			{
				string xml = File.ReadAllText(exception.FullName);
				XmlDocument doc = new XmlDocument();
				doc.LoadXml(xml);
				foreach (XmlNode node in doc.SelectNodes("//error"))
				{
					string errorId = node.Attributes["errorId"].Value;
					string type = node.Attributes["type"].Value;
					string message = node.Attributes["message"].Value;
					string time = node.Attributes["time"].Value;
					ws.Cells[i, 0].Value = errorId;
					ws.Cells[i, 1].Value = type;
					ws.Cells[i, 2].Value = message;
					ws.Cells[i, 3].Value = time;
					i++;
				}

				foreach (XmlNode node in doc.SelectNodes("//item[@name=]/value"))
				{
					string nodeValue = node.Attributes["string"].Value;
					//if (node.Attributes["string"].Value.Substring(0, 5) == "HOST:")
					//{
					//	host = node.Attributes["string"].Value;
					//}			
					Console.WriteLine("node Values {0}", nodeValue);
					Console.ReadLine();
				}
			}
			
			//host.ToString();

			ef.Save("C:\\ErrorsMay2017\\errorlog " + checkTime.ToString("MM-dd-yyyy") + ".xls");
		}
	}
}
