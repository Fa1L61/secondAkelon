using System;
using System.IO;
using ClosedXML;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Spreadsheet;
using secondAkelon.Managers;
using secondAkelon.Models;

namespace ConsoleApp
{
    class Program
    {
        static void Main(string[] args)
        {


            ExcelDataManager excelDataManager = new ExcelDataManager("C:\\Users\\fa1l7\\Downloads\\copy.xlsx");
            excelDataManager.CustomerInfo("Молоко");


        }
    }
}