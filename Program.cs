using System;
using System.Collections.Generic;
using System.IO;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Spreadsheet;

class Program
{
    static void Main(string[] args)
    {
        //Ścieżki plików. Plik do edycji wyciągamy dokładną ścieżkę z powodu różnych nazw pliku tylko końcówka taka sama
        ProgressBar(0, 100);
        string mainPath = "H01w.xlsx";
        string folder = Directory.GetCurrentDirectory();
        string[] files = Directory.GetFiles(folder, "*-H01w.xlsx");
        string checkedPath = files[0];
        string outputPath = "H01w_gotowy.xlsx";
        ProgressBar(10, 100);

        //Otwarcie dokumentów excelowych 
        using var wbMain = new XLWorkbook(mainPath);
        using var wbChecked = new XLWorkbook(checkedPath);
        ProgressBar(20, 100);

        //Otwarcie skoroszytów i przypisanie danych do zmiennej
        var sheetMain = wbMain.Worksheet(1);
        var sheetChecked = wbChecked.Worksheet(1);
        ProgressBar(30, 100);

        sheetMain.Unprotect();
        sheetChecked.Unprotect();

        //Słownik na dane które znalazły się w excelu głównym 
        var dict =  new Dictionary<string, IXLRow>();
        ProgressBar(40, 100);

        //Pobranie numer końcowej kolumny i wiersza
        int lastRowMain = sheetMain.LastRowUsed().RowNumber();
        int lastColMain = sheetMain.LastColumnUsed().ColumnNumber();
        ProgressBar(50, 100);

        for (int row = 2; row <= lastRowMain; row++)
        {
            //Stworzenie klucza po którym będą wyszukiwane dane z excela z wszystkimi danymi, który składa sie z "regon|numer_firmy"
            string key = sheetMain.Cell(row, 1).GetString().Trim() + "|" + sheetMain.Cell(row, 2).GetString().Trim();
            
            //Stworzenie wiersza z danymi jeśli nie istnieje on w słowniku
            if (!dict.ContainsKey(key))
            {
                dict[key] = sheetMain.Row(row);
            }
        }
        ProgressBar(60, 100);

        //Pobranie numer końcowego wiersza
        int lastRowChecked = sheetChecked.LastRowUsed().RowNumber();

        for (int row = 2; row <= lastRowChecked; row++)
        {
            //Stworzenie klucza dla excela sprawdzanego
            string key = sheetChecked.Cell(row, 1).GetString().Trim() + "|" + sheetChecked.Cell(row, 2).GetString().Trim();

            //Wyszukuje w słowniku po kluczu, wartości i przypisuje je do zmiennej main
            if (dict.TryGetValue(key, out var main))
            {
                for (int col = 3; col <= 7; col++)
                {
                    //przypisanie wartości z danej kolumny do zmiennej 
                    string adresMain = main.Cell(col).GetString();
                    string adresChecked = sheetChecked.Cell(row, col).GetString();

                    //Jeśli nie są takie same podmienia wartość w kolumnie oraz zmienia kolor na żółty
                    if (adresMain != adresChecked)
                    {
                        sheetChecked.Cell(row, col).Value = adresMain;
                        sheetChecked.Cell(row, col).Style.Fill.BackgroundColor = XLColor.Yellow;
                    }
                }

                //Przypisanie wartości z danej kolumny do zmiennej
                string detaluMain = main.Cell(17).GetString();
                string detaluChecked = sheetChecked.Cell(row, 17).GetString();

                //Jeśli kolumna 17 jest pusta to wartość z kolumny z excela wzorcowego jest przenoszona
                if (String.IsNullOrEmpty(detaluChecked))
                {
                    sheetChecked.Cell(row, 17).Value = detaluMain;
                }

                //Sprawdzanie czy sp_b i sp_u oraz fo_b i fo_u nie są puste oraz czy są takie same
                HighlightDifferent(sheetChecked, row, 8, 9);
                HighlightDifferent(sheetChecked, row, 14, 15);
            }
        }
        ProgressBar(70, 100);

        //Tworzenie nowych kolumn oraz nazwanie nagłówków
        sheetChecked.Column(11).InsertColumnsAfter(2).Clear(XLClearOptions.AllFormats | XLClearOptions.Contents);
        sheetChecked.Cell(1, 12).Value = "roznica_pow_b_pow_u";
        sheetChecked.Cell(1, 13).Value = "dynamika_pow";
        sheetChecked.Column(15).InsertColumnsAfter(2).Clear(XLClearOptions.AllFormats | XLClearOptions.Contents);
        sheetChecked.Cell(1, 16).Value = "roznica_lpr_b_lpr_u";
        sheetChecked.Cell(1, 17).Value = "dynamika_lpr";
        sheetChecked.Column(21).InsertColumnsAfter(2).Clear(XLClearOptions.AllFormats | XLClearOptions.Contents);
        sheetChecked.Cell(1, 22).Value = "roznica_detal_detal_u";
        sheetChecked.Cell(1, 23).Value = "dynamika_detal";
        ProgressBar(80, 100);
        
        //Petla wykonuje działania na kolumnach
        for (int row = 2; row <= lastRowChecked; row++)
        {
            //Zaokrąglenie kolumn
            sheetChecked.Cell(row, 10).Value = Round(Read(sheetChecked.Cell(row, 10)), 0);
            sheetChecked.Cell(row, 14).Value = Round(Read(sheetChecked.Cell(row, 14)), 0);
            sheetChecked.Cell(row, 20).Value = Round(Read(sheetChecked.Cell(row, 20)), 1);

            sheetChecked.Cell(row, 10).Style.NumberFormat.Format = "0";
            sheetChecked.Cell(row, 14).Style.NumberFormat.Format = "0";
            sheetChecked.Cell(row, 20).Style.NumberFormat.Format = "0.0";

            //Wywołanie funkcji liczącej roznice i dynamikę
            Calc(sheetChecked, row, 10, 11, 12, 13, 0);
            Calc(sheetChecked, row, 14, 15, 16, 17, 0);
            Calc(sheetChecked, row, 20, 21, 22, 23, 1);

            //Kolorowanie jeśli dany rekord spełnia warunek
            if (Read(sheetChecked.Cell(row, 13)) >= 150 || Read(sheetChecked.Cell(row, 13)) <= 50)
            {
                sheetChecked.Cell(row, 13).Style.Fill.BackgroundColor = XLColor.Pumpkin;
            }

            double? detal = Read(sheetChecked.Cell(row, 20));
            double? detal_u = Read(sheetChecked.Cell(row, 21));

            if ((Read(sheetChecked.Cell(row, 23)) >= 160 || Read(sheetChecked.Cell(row, 23)) <=  40) && (Math.Abs((detal ?? 0) - (detal_u ?? 0)) > 2000))
            {
                sheetChecked.Cell(row, 23).Style.Fill.BackgroundColor = XLColor.Pumpkin;
            }
        }
        ProgressBar(90, 100);

        //Funkcja kolorująca kolumny w zależności czy UW jest puste lub wynosi 5
        MarkMissingData(sheetChecked, lastRowChecked);
        ProgressBar(95, 100);

        wbChecked.SaveAs(outputPath);
        ProgressBar(100, 100);
        Console.WriteLine("\n\nGotowe! Zapisano jako: " + outputPath);
    }
    
    //Funkcja czytająca dany wiersz i kolumne, spowodowane jest to tym, że mamy rekordy puste w excelu i jeśli chcemy użyć zmienno przecinkowego typu musi on być zabezpiecznony przed nullem, ? dlatego jest używany.
    static double? Read(IXLCell c)
    {
        return c.TryGetValue<double>(out var v) ? v : null;
    }

    //Taki sam powód powstania funkcji zaokrąglającej jak Read.
    static double? Round(double? v, int places)
    {
        return v.HasValue ? Math.Round(v.Value, places) : null;
    }

    //Sprawdza czy wartość istnieje a następnie porównuje obie wartości czy są równe
    static void HighlightDifferent(IXLWorksheet ws, int row, int col1, int col2)
    {
        double? v1 = Read(ws.Cell(row, col1));
        double? v2 = Read(ws.Cell(row, col2));

        if (v1.HasValue && v2.HasValue && v1 != v2)
        {
            ws.Cell(row, col1).Style.Fill.BackgroundColor = XLColor.Aqua;
            ws.Cell(row, col2).Style.Fill.BackgroundColor = XLColor.Aqua;
        }
    }

    //Funkcja obsługująca działania matematyczne
    static void Calc(IXLWorksheet ws, int row, int colB, int colU, int colDiff, int colDyn, int round)
    {
        double? b = Read(ws.Cell(row, colB));
        double? u = Read(ws.Cell(row, colU));

        double? diff = b - u;

        if (diff == 0)
        {
            ws.Cell(row, colDiff).Value = "";
        }
        else
        {

            ws.Cell(row, colDiff).Value = (round == 1) ? Round(diff, 1) : Round(diff, 0);
            ws.Cell(row, colDiff).Style.NumberFormat.Format = (round == 1) ? "0.0" : "0";
        }

        double? dyn = (u == 0 || u == null) ? null : (b / u) * 100;

        if (dyn == 100.0)
        {
            ws.Cell(row, colDyn).Value = "";
        }
        else
        {
            ws.Cell(row, colDyn).Value = Round(dyn, 1);
            ws.Cell(row, colDyn).Style.NumberFormat.Format = "0.0";
        }
    }

    //Funkcja rysująca pasek postępu
    static void ProgressBar(int current, int total)
    {
        int width = 100; 
        double percentage = (double)current / total;
        int filled = (int)(percentage * width);

        string bar = "[" + new string('#', filled) + new string('-', width - filled) + $"] {percentage:0.0%}";
        Console.Write("\r" + bar);
    }

    //Funkcja kolorująca brakujące dane jeśli kolumna UW jest pusta sprawdza czy kolumny sp_b, pow_b, lpr_b, fo_b lub detal są puste koloruje na czerwono
    //Jeśli UW ma wartość wpisaną 5 to koloruje kolumny powyżej jeśli są puste
    static void MarkMissingData(IXLWorksheet ws, int lastRow)
    {
        int[] colsToCheck = { 8, 10, 14, 18, 20 };

        for (int row = 2; row <= lastRow; row++)
        {
            if (Read(ws.Cell(row, 24)) == 5)
            {
                foreach (int col in colsToCheck)
                {
                    if (!ws.Cell(row, col).IsEmpty())
                    {
                        ws.Cell(row, col).Style.Fill.BackgroundColor = XLColor.Red;
                    }
                }
                continue;
            }
            
            if (!ws.Cell(row, 24).IsEmpty())
                continue;

            foreach (int col in colsToCheck)
            {
                if (ws.Cell(row, col).IsEmpty())
                {
                    ws.Cell(row, col).Style.Fill.BackgroundColor = XLColor.Red;
                }
            }
        }
    }
}