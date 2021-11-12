using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using NPOI.OpenXml4Net.OPC;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelParser
{
    class Program
    {
        static void Main(string[] args)
        {
            List<AccountInfo> listAccountInfos = new List<AccountInfo>();

            using (FileStream file = new FileStream(@"C:\...\Workbook1.xlsx", FileMode.Open, FileAccess.Read))
            {
                
                OPCPackage pkg = OPCPackage.Open(file);
                XSSFWorkbook wb = new XSSFWorkbook(pkg);
                ISheet sheet = wb.GetSheet("Sheet1");

                //foreach (var x in sheet)
                //{
                //    AccountInfo accountInfo = new AccountInfo();
                //    var r = 0;
                //    accountInfo.AccountNumber = sheet.GetRow(r).GetCell(0).NumericCellValue;
                //    listAccountInfos.Add(accountInfo);
                //    //Console.WriteLine(sheet.GetRow(r).GetCell(0));
                //    r++;
                //}

                for (int row = 0; row <= sheet.LastRowNum; row++)
                {
                    if (sheet.GetRow(row) != null)
                    {
                        AccountInfo accountInfo = new AccountInfo();
                        accountInfo.AccountNumber = sheet.GetRow(row).GetCell(0).NumericCellValue;
                        listAccountInfos.Add(accountInfo);
                        //Console.WriteLine(string.Format("Row {0} = {1}", row, sheet.GetRow(row).GetCell(0).NumericCellValue));
                    }
                }
            }

            foreach (var i in listAccountInfos)
            {
                Console.WriteLine(i.AccountNumber);
            }

            string stringOfAccountNumbers = "('" + string.Join("','", listAccountInfos.Select(l => l.AccountNumber)) + "')";
            Console.Write(stringOfAccountNumbers);
            Console.ReadKey();
        }
    }

    public class AccountInfo
    {
        public double AccountNumber { get; set; }
    }
}
