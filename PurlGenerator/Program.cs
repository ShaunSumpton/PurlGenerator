using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PurlGenerator
{
    class Program
    {
        static void Main(string[] args)
        {

            Console.Write("Paste File Location:"); // Past in location of file
            string fileLoc = Console.ReadLine(); // Read location of file

            string[] filePaths = Directory.GetFiles(fileLoc, "*.xlsx"); // get files in directory with .xlsx extention to a string 

            int a = 1;

            foreach (string i in filePaths) // loop through files and display in console window
            { 
                Console.WriteLine(a+")"+i); // display full file location in console window
                a++;
            }

            Console.Write("Select File Number:"); // ask user to select file number "Each file listed will be prefixed with a number"
            var fn  = Console.ReadLine();

            Application application = new Application(); // Create Excel Instance

            Workbook exceldoc = application.Workbooks.Open(filePaths[Convert.ToInt32(fn)-1]); // create workbook
            Worksheet ws; // create worksheet

            ws = (Worksheet)exceldoc.Worksheets[1]; // worksheet assigned to 1st sheet in workbook


            int LastRow = ws.UsedRange.Rows.Count;    // find last row and last column of sheet

            Range last = ws.Cells.SpecialCells(XlCellType.xlCellTypeLastCell, Type.Missing);
            _ = ws.get_Range("A1", last);

            Console.Write("Surname Comlumn Letter (e.g A) : "); // ask for column where surname will be to generate PURL
            string Fname = Console.ReadLine();

            Console.Write("Purl Column Letter (e.g B):"); // ask for column where Generated PURL will be 
            String Purl = Console.ReadLine();

            for (int i = 2; i < LastRow + 1; i++)
            {

                string temp = ws.Range[Fname + i].Value.ToString(); // check length of surname
                int iTemp = temp.Length;

                if (iTemp > 7)
                {
                    iTemp = 8;

                }


                // build perl

                string Sname = temp.Substring(0, iTemp); // if surname over 8 chars trim to 8

                Sname = Sname.Replace("@", "").Replace(" ", "").Replace("/", "").Replace(".", "").Replace(",", "").Replace("'", "")
                .Replace("&", "").Replace("(", "").Replace(")", "").Replace("\"", "").Replace("-", "").Replace(@"\", "").Replace("+", ""); // Replace chars in purl

                char l1 = RandomLetter.GetLetter(); // Generate random number or letters for purl
                int n1 = RandomNumber.GetNumber();
                char l2 = RandomLetter.GetLetter();
                int n2 = RandomNumber.GetNumber();
                char l3 = RandomLetter.GetLetter();
                int n3 = RandomNumber.GetNumber();

                ws.Range[Purl + i].Value = Sname + l1 + n1 + l2 + n2 + l3 + n3; // copy to column


            }

            var desktopFolder = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            string FileLoc = desktopFolder + @"\PURL.xlsx"; // save to users desktop
            


            if (File.Exists(FileLoc)) // check if file already exsists, if so replace file
            {
                File.Delete(FileLoc);
            }

            exceldoc.SaveAs(FileLoc, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlExclusive, Type.Missing, Type.Missing, Type.Missing,
                        Type.Missing, Type.Missing);
            Console.WriteLine("File Saved");
            exceldoc.Close(false);
           
        }
    }
    class RandomLetter
    {
        static Random _random = new Random();
        public static char GetLetter()
        {
            // This method returns a random lowercase letter.
            // ... Between 'a' and 'z' inclusize.
            int num = _random.Next(0, 26); // Zero to 25
            char let = (char)('A' + num);
            return let;
        }
    }

    class RandomNumber
    {
        // ... Create new Random object.
        static Random random = new Random();

        public static int GetNumber()
        {
        
         
            int num = random.Next(1, 10);
            return num;
        }

    }
}

