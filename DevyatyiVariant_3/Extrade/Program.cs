using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Aspose.Cells;

namespace Aksel
{
    class Program
    {
        static void Main(string[] args)
        {

            string Path;
            Console.WriteLine("Тип файла xml?\n(1)-Нет\n(2)-Да\n");
            int FyleType = Convert.ToInt32(Console.ReadLine());
            while (FyleType < 2)
            {
                switch (FyleType)
                {
                    case 1:
                        {
                            FyleType = FyleType + 10;
                            break;
                        }
                    case 2:
                        {
                            Console.WriteLine("Введите путь к xml-файлу, по типу:\t C:/Users/User/Desktop/FileToCombine.xml\t:\n");
                            Path = Console.ReadLine();
                            var workbook = new Workbook(Path);
                            Console.WriteLine("Где сохранить xml файл?\n");
                            string PathToSave; PathToSave = Console.ReadLine();
                            workbook.Save(PathToSave);
                            Console.WriteLine("Новый файл xlsx создан:\n");
                            FyleType = FyleType + 10;
                            break;
                        }
                }
            }
            Console.WriteLine("Введите путь к файлу:\n");
            Path = Console.ReadLine();
            Workbook WorkBook = new Workbook(Path);

            int nn = 0;
            string[] nameOfSourceWorksheets = {""};
            while (WorkBook.Worksheets[nn] != null)
            {
                nameOfSourceWorksheets[nn] = WorkBook.Worksheets[nn].Name;
                nn++;
            }
            //string[] nameOfSourceWorksheets = { "1. Статьи в журналах", "2. Публ. в научн. сборниках", "3. Монографии", "4. Иные научные произведения", "6. Учебники и уч. пособия"};


            
            int Vybor = 0;
            while (Vybor != 9) 
            {
                Console.WriteLine("Выберите действие:\n(1)-Выбрать лист\n(2)-Обратить файл в xml\n(9)-Выход\n");
                Vybor = Convert.ToInt32(Console.ReadLine());
                switch (Vybor)
                {
                    case 1:
                        {

                            Console.WriteLine("Выберите лист:\n");
                            for (int i = 0; i < nameOfSourceWorksheets.Length; i++)
                            {
                                Console.Write("("); Console.Write(i + 1); Console.Write(")\t=\t");
                                Console.Write(nameOfSourceWorksheets[i]);
                                Console.WriteLine("\n");
                            }
                            int k = Convert.ToInt32(Console.ReadLine());
                            k = k - 1;
                            Worksheet WSFirst = WorkBook.Worksheets[k];
                            if (k < 0 || k > nameOfSourceWorksheets.Length)
                            {
                                Console.WriteLine("Некорректный ввод, такого листа не существует\n");
                            }
                            string sheetName = nameOfSourceWorksheets[k]; //Лист выбран 
                            /*Console.Write("Выбран лист, откуда будет перенос:\t");*/ Console.WriteLine(sheetName);
                            Console.WriteLine("Выберите номер колонны для переноса:\n");
                            int I = Convert.ToInt32(Console.ReadLine());
                            Console.WriteLine("Сколько строк содержат заголовкии?");
                            int Zagolovki = Convert.ToInt32(Console.ReadLine());



                            Console.WriteLine("Введите лист куда перенести:");
                            for (int i = 0; i < nameOfSourceWorksheets.Length; i++)
                            {
                                Console.Write("("); Console.Write(i + 1); Console.Write(")\t=\t");
                                Console.Write(nameOfSourceWorksheets[i]);
                                Console.WriteLine("\n");
                            }
                            int n = Convert.ToInt32(Console.ReadLine());
                            n = n - 1;
                            int m = n;
                            if (n < 0 || n > nameOfSourceWorksheets.Length)
                            {
                                Console.WriteLine("Некорректный ввод, такого листа не существует\n");
                            }
                                /*Console.WriteLine(WorkBook.Worksheets[n].Cells[0, 0].Value);*/  Console.WriteLine('\n');

                                Worksheet WSSecond = WorkBook.Worksheets[m];
                                Console.WriteLine("СоздаЛОСЬ");

                            Console.WriteLine("Выберите колонну, в которую перенести:");
                            int J = Convert.ToInt32(Console.ReadLine());

                            // WSFirst.Cells.CopyColumn(WSFirst.Cells,/*Источник переноса*/ WSFirst.Cells.Columns[I-1].Index,/*Цель переноса*/ WSSecond.Cells.Columns[J-1].Index);

                            //I -- Переносимая колонна, J -- Та в которую

                            int Counter_Rows_Second = WorkBook.Worksheets[m].Cells.Rows.Count; //Создаём счётчик, принимающий значение количества СТРОК
                            int Counter_Columns_Second = WorkBook.Worksheets[m].Cells.Columns.Count; //Создаём счётчик, принимающий значение колиечества КОЛОНН
                            //WSSecond -- ТА, В КОТОРУЮ ПЕРЕНОСИМ, //WSFirst -- ТА, КОТОРУЮ ПЕРЕНОСИМ
                            //Console.WriteLine(Counter_Rows_Second);
                            //Console.WriteLine(Counter_Columns_Second);
                            int Counter_Rows_First = WorkBook.Worksheets[k].Cells.Rows.Count - Zagolovki;
                            int Counter_Columns_First = WorkBook.Worksheets[k].Cells.Columns.Count;

                            //Начало переноса:
                            for (int i = Counter_Rows_Second; i < Counter_Rows_Second + Counter_Rows_First ; i++)
                            {
                                WSSecond.Cells[i,J].Value = WSFirst.Cells[Zagolovki + 1,I].Value;
                            }
                            break;
                        }

                    case 2:
                        {
                            var workbook = new Workbook(Path);
                            Console.WriteLine("Где сохранить xml файл?\n");
                            string PathToSave; PathToSave = Console.ReadLine();
                            workbook.Save(PathToSave);
                            break;
                        }

                    case 9: 
                        {
                            break;
                        }
                }
            }




            // Сохранение документа
            Console.WriteLine("Введите путь для сохранения нового документа:\n");
            string PathToEnd; PathToEnd = Console.ReadLine();
            WorkBook.Save(PathToEnd);

            //Перевод в XML
            //var workbook = new Workbook("C:/Users/Bonolenov/Desktop/Items.xml");
            //workbook.Save("C:/Users/Bonolenov/Desktop/Items.xlsx");
        }
    }
}