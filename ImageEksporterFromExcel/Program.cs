using System;
using System.IO;
using Spire.Xls;
using System.Drawing.Imaging;
using System.Drawing;

namespace ImageEksporterFromExcel
{
    class Program
    {
        static void Main(string[] args)
        {
            String filePath;
            string[] dirs;
            String FileName;
            Workbook workbook = new Workbook();
            int[] tab = new int[1000];

            MemoryStream ms;
            FileStream fs;

            Console.WriteLine("Podaj lokalizację pliku:");
            filePath = Console.ReadLine();
            if (filePath[filePath.Length - 1] != '\\')
            {
                filePath = filePath + '\\';
            }
            dirs = Directory.GetFiles(filePath, "*.xls*");

            if (dirs.Length > 1)
            {
                for (int i = 0; i < dirs.Length; i++)
                {
                    Console.WriteLine(i + ". " + dirs[i]);
                }
                Console.WriteLine("Podaj numer pliku");

                FileName = dirs[int.Parse(Console.ReadLine())];
            }
            else
            {
                Console.WriteLine(dirs[0]);
                FileName = dirs[0];
            }

            workbook.LoadFromFile(FileName);
            Worksheet sheet = workbook.Worksheets[0];


            Console.WriteLine(sheet.Pictures.Count);
            for (int i = 0; i < sheet.Pictures.Count; i++)
            {
                sheet.Pictures[i].Top += sheet.Pictures[i].Height / 2;
                sheet.Pictures[i].Left += sheet.Pictures[i].Width / 2;
                sheet.Pictures[i].Scale(20, 20);
            }

            string tmp;
            for (int i = 0; i < sheet.Pictures.Count; i++)
            {
                try
                {
                    Console.WriteLine(sheet.GetFormulaStringValue(sheet.Pictures[i].TopRow, 1));
                    Console.WriteLine(sheet.GetText(sheet.Pictures[i].TopRow, 1));
                    Console.WriteLine(sheet.GetNumber(sheet.Pictures[i].TopRow, 1));
                    Console.WriteLine(sheet.GetCaculateValue(sheet.Pictures[i].TopRow, 1).ToString());
                    if (sheet.GetText(sheet.Pictures[i].TopRow, 1) != null)
                    {
                        tmp = sheet.GetText(sheet.Pictures[i].TopRow, 1);
                    }
                    else if (!Double.IsNaN(sheet.GetNumber(sheet.Pictures[i].TopRow, 1)))
                    {
                        tmp = sheet.GetNumber(sheet.Pictures[i].TopRow, 1).ToString();
                    }
                    else
                    {
                        tmp = sheet.GetFormulaStringValue(sheet.Pictures[i].TopRow, 1);
                    }

                    Console.WriteLine("text: " + tmp + "row: " + sheet.Pictures[i].TopRow + " column: " + sheet.Pictures[i].LeftColumn);
                    if (tab[sheet.Pictures[i].TopRow] == 0)
                    {
                        fs = File.Create(filePath + tmp + ".jpg");
                        //fs = File.Create(filePath + sheet.Pictures[i].TopRow + ".jpg");
                    }
                    else
                    {
                        fs = File.Create(filePath + tmp + "_" + tab[sheet.Pictures[i].TopRow] + ".jpg");
                        //fs = File.Create(filePath + sheet.Pictures[i].TopRow + "_" + tab[sheet.Pictures[i].TopRow] + ".jpg");
                    }



                    tab[sheet.Pictures[i].TopRow]++;

                    Console.WriteLine(sheet.Pictures[i].Picture.Width + "x" + sheet.Pictures[i].Picture.Height);
                    ms = new MemoryStream();
                    Console.WriteLine();


                    
;
                    sheet.Pictures[i].Rotation = 0;
                    sheet.Pictures[i].Picture.Reset();


                    sheet.Pictures[i].Width = sheet.Pictures[i].Picture.Width;
                    sheet.Pictures[i].Height = sheet.Pictures[i].Picture.Height;

                    sheet.Pictures[i].SaveToImage(ms);    
                    
                    Image.FromStream(ms).Save(fs, ImageFormat.Jpeg);
                }
                catch (NullReferenceException ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
        }
    }
}
