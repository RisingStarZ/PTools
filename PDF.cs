///Copyright (C) 2017  RisingStarZ
///This program is free software: you can redistribute it and/or modify it under the terms 
///of the GNU Affero General Public License as published by the Free Software Foundation, 
///either version 3 of the License, or(at your option) any later version.
///
///This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; 
///without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.
///See the GNU Affero General Public License for more details.
///
///You should have received a copy of the GNU Affero General Public License along with this program.
///If not, see<http://www.gnu.org/licenses/>.

using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;

namespace PTools
{
    public static class PDF
    {
        /// <summary>
        /// Рисует один штрихкод в файле PDF на определенной странице
        /// </summary>
        /// <param name="x">для ГИБДД кемерова значение по умолчанию 380</param>
        /// <param name="y">для ГИБДД кемерова значение по умолчанию 685</param>
        public static void AddBarcode(string file, int page, string barcode, int x, int y, bool replaceSource)
        {
            try
            {
                using (var ms = new MemoryStream())
                {
                    using (var reader = new PdfReader(file))
                    using (var stamper = new PdfStamper(reader, ms))
                    {
                        if (page < 1 || page > reader.NumberOfPages) throw new Exception("Номер страницы для вставки штрихкода не входит в диапазон страниц файла");
                        if (barcode.Length != 14) throw new Exception("Длина штрихкода должна быть 14 символов");

                        var stCb = stamper.GetOverContent(2);
                        var bc = new BarcodeInter25() { Code = barcode };
                        stCb.SetColorFill(BaseColor.WHITE);
                        stCb.SetColorStroke(BaseColor.WHITE);

                        stCb.MoveTo(x - 15, y + 45);
                        stCb.LineTo(x + 165, y + 45);
                        stCb.LineTo(x + 165, y);
                        stCb.LineTo(x - 15, y);
                        stCb.Fill();

                        var bc_image = bc.CreateImageWithBarcode(stCb, BaseColor.BLACK, BaseColor.BLACK);
                        bc_image.SetAbsolutePosition(x, y);
                        bc_image.ScalePercent(117F);
                        stCb.AddImage(bc_image);
                    }
                    ms.Flush();
                    if (replaceSource) File.WriteAllBytes(file, ms.ToArray());
                    else File.WriteAllBytes(file.Insert(file.Length - 4, "_out"), ms.ToArray());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error in PTools.dll", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Рисует один штрихкод в файле PDF на определенной странице
        /// </summary>
        /// <param name="x">для ГИБДД кемерова значение по умолчанию 380</param>
        /// <param name="y">для ГИБДД кемерова значение по умолчанию 685</param>
        /// <param name="step">Шаг вставки штрихкода. Так, если step = 1, то штрихкод быдет на каждой странице, step = 2 - через одну и.т.д.</param>
        /// <param name="y">для ГИБДД кемерова значение по умолчанию 685</param>
        /// <param name="barcodeList">Список штрихкодов</param>
        public static void AddBarcodeSeqence(string file, int step, List<string> barcodeList, int x, int y, bool replaceSource)
        {
            using (var ms = new MemoryStream())
            {
                using (var reader = new PdfReader(file))
                using (var stamper = new PdfStamper(reader, ms))
                {
                    for (int i = 1; i <= reader.NumberOfPages; i += step)
                    {
                        try
                        {
                            if (reader.NumberOfPages % step > 0) throw new Exception("Заданный интервал step не вмещается в кол-во страниц целое число раз");
                            if (barcodeList[i - 1].Length != 14) throw new Exception(string.Format("Длина штрихкода #{0} должна быть 14 символов", i + 1));

                            var stCb = stamper.GetOverContent(i);
                            var bc = new BarcodeInter25() { Code = barcodeList[i - 1] };
                            stCb.SetColorFill(BaseColor.WHITE);
                            stCb.SetColorStroke(BaseColor.WHITE);

                            stCb.MoveTo(x - 15, y + 45);
                            stCb.LineTo(x + 165, y + 45);
                            stCb.LineTo(x + 165, y);
                            stCb.LineTo(x - 15, y);
                            stCb.Fill();

                            var bc_image = bc.CreateImageWithBarcode(stCb, BaseColor.BLACK, BaseColor.BLACK);
                            bc_image.SetAbsolutePosition(x, y);
                            bc_image.ScalePercent(117F);
                            stCb.AddImage(bc_image);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "Error in PTools.dll", MessageBoxButton.OK, MessageBoxImage.Error);
                        }
                    }
                }
                ms.Flush();
                if (replaceSource) File.WriteAllBytes(file, ms.ToArray());
                else File.WriteAllBytes(file.Insert(file.Length - 4, "_out"), ms.ToArray());
            }
        }

        /// <summary>
        /// Insert single addressee block in PDF file
        /// </summary>
        /// <param name="file">input file</param>
        /// <param name="beforePage">block will be inserted before this page</param>
        /// <param name="doublesided">if true, blanc page will be added after addressee block</param>
        /// <param name="person">Addressee data</param>
        /// <param name="x">lower left corner x coordinate</param>
        /// <param name="y">lower left corner y coordinate</param>
        /// <param name="replaceSource">Replace original file by result</param>
        public static void AddAddresseBlock(string file, int beforePage, bool doublesided, Addressee person, int x, int y, bool replaceSource)
        {
            using (var ms = new MemoryStream())
            {
                using (var doc = new Document(iTextSharp.text.PageSize.A4))
                {
                    var reader = new PdfReader(file);
                    var writer = PdfWriter.GetInstance(doc, ms);
                    writer.SetPdfVersion(PdfWriter.PDF_VERSION_1_7);
                    writer.CompressionLevel = PdfStream.NO_COMPRESSION;
                    string ttf = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "TIMES.TTF");
                    var baseFont = BaseFont.CreateFont(ttf, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                    var timesN = new Font(baseFont, 9, Font.NORMAL);
                    var timesB = new Font(baseFont, 9, Font.BOLD);

                    doc.Open();
                    var cb = writer.DirectContent;
                    int n = 1;
                    if (beforePage > 1)
                    {
                        for (; n <= beforePage; n++)
                        {
                            cb.AddTemplate(writer.GetImportedPage(reader, n), 0, 0);
                            doc.NewPage();
                        }
                    }

                    if (person.Barcode != null)
                    {
                        var bc = new BarcodeInter25() { Code = person.Barcode };
                        var bc_image = bc.CreateImageWithBarcode(cb, BaseColor.BLACK, BaseColor.BLACK);

                        bc_image.SetAbsolutePosition(355, 665);
                        bc_image.ScalePercent(117F);
                        cb.AddImage(bc_image);
                    }
                    var ct = new ColumnText(cb);
                    ct.SetSimpleColumn(new Rectangle(345, 550, 550, 665));

                    if (person.Name != null) ct.AddElement(new Paragraph("Кому: " + person.Name, timesN));
                    if (person.Address != null) ct.AddElement(new Paragraph("Куда: " + person.Address, timesN));
                    ct.Go();

                    if (person.Index != 0)
                    {
                        cb.SetFontAndSize(BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 15);
                        cb.BeginText();
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, person.Index.ToString(), 465, 680, 0);
                        cb.EndText();
                    }
                    doc.NewPage();
                    if (doublesided)
                    {
                        doc.Add(new Chunk());
                        doc.NewPage();
                    }
                    for (; n <= reader.NumberOfPages; n++)
                    {
                        cb.AddTemplate(writer.GetImportedPage(reader, n), 0, 0);
                        doc.NewPage();
                    }
                }
                ms.Flush();
                if (replaceSource) File.WriteAllBytes(file, ms.ToArray());
                else File.WriteAllBytes(file.Insert(file.Length - 4, "_out"), ms.ToArray());
            }
        }

        public static void AddAddresseBlockSeqence(string file, int startBeforePage, int step, bool doublesided, List<Addressee> roll, int x, int y, bool replaceSource) //llx:345 lly:550 defaults
        {
            using (var ms = new MemoryStream())
            {
                using (var doc = new Document(iTextSharp.text.PageSize.A4))
                {
                    var reader = new PdfReader(file);
                    var writer = PdfWriter.GetInstance(doc, ms);
                    writer.SetPdfVersion(PdfWriter.PDF_VERSION_1_5);
                    string ttf = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "TIMES.TTF");
                    var baseFont = BaseFont.CreateFont(ttf, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                    var timesN = new Font(baseFont, 9, Font.NORMAL);
                    var timesB = new Font(baseFont, 9, Font.BOLD);

                    doc.Open();
                    var cb = writer.DirectContent;
                    int n = 0;
                    for (int i = startBeforePage; i <= reader.NumberOfPages; i++)
                    {
                        if (startBeforePage > 1)
                        {
                            for (int j = 1; j <= startBeforePage; j++)
                            {
                                cb.AddTemplate(writer.GetImportedPage(reader, j), 0, 0);
                                doc.NewPage();
                            }
                        }

                        if (roll[n].Barcode != null && roll[n].Barcode != "")
                        {
                            var bc = new BarcodeInter25() { Code = roll[n].Barcode };
                            var bc_image = bc.CreateImageWithBarcode(cb, BaseColor.BLACK, BaseColor.BLACK);

                            bc_image.SetAbsolutePosition(x + 10, y + 115);
                            bc_image.ScalePercent(117F);
                            cb.AddImage(bc_image);
                        }

                        var ct = new ColumnText(cb);
                        ct.SetSimpleColumn(new Rectangle(x, y, x + 205, y + 115));
                        ct.AddElement(new Paragraph("Кому: " + roll[n].Name, timesN));
                        ct.AddElement(new Paragraph("Куда: " + roll[n].Address, timesN));
                        ct.Go();

                        cb.SetFontAndSize(BaseFont.CreateFont(BaseFont.TIMES_BOLD, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 15); //llx:345 lly:550
                        if (roll[n].Index != 0)
                        {
                            cb.BeginText();
                            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, roll[n].Index.ToString(), x + 120, y + 130, 0);
                            cb.EndText();
                        }
                        cb.SetFontAndSize(BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, BaseFont.NOT_EMBEDDED), 5);
                        cb.BeginText();
                        cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, (n + 1).ToString(), x + 120, y + 120, 0);
                        cb.EndText();
                        doc.NewPage();
                        if (doublesided)
                        {
                            doc.Add(new Chunk());
                            doc.NewPage();
                        }
                        for (int j = 0; j < step; j++)
                        {
                            cb.AddTemplate(writer.GetImportedPage(reader, i), 0, 0);
                            doc.NewPage();
                        }
                        //if (n == roll.Count - 1) break;
                        n++;
                    }
                }
                ms.Flush();
                if (replaceSource) File.WriteAllBytes(file, ms.ToArray());
                else File.WriteAllBytes(file.Insert(file.Length - 4, "_out"), ms.ToArray());
            }
        }

        public static void SplitPdf(string pdfPath, int step, string outPath = "")
        {
            var sw = new Stopwatch();
            sw.Start();
            string path = outPath == "" ? "" : outPath + "\\";
            var files = Directory.EnumerateFiles(pdfPath, "*.pdf");
            if (files.Count() == 0) throw new Exception("По указанному пути PDF файлы не найдены.");
            foreach (var file in files)
            {
                using (var reader = new PdfReader(file))
                {
                    if (reader.NumberOfPages % 2 != 0) throw new Exception("В документе нечетное количество страниц");

                    for (int i = 1; i <= reader.NumberOfPages; i += step)
                    {
                        string barcode = string.Empty;
                        using (var fs = new FileStream($"{path}{i}.pdf", FileMode.Create))
                        {
                            var doc = new Document();
                            var copy = new PdfCopy(doc, fs);
                            doc.Open();
                            var matches = new List<string>();
                            for (int j = 0; j < step; j++)
                            {
                                var strategy = new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy();
                                //var parsedText = string.Join("", iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(reader, i + j).Split(default(string[]), StringSplitOptions.RemoveEmptyEntries));
                                var parsedText = iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(reader, i + j);

                                //var index = parsedText.IndexOf("999999");
                                //if (index >= 0) barcode = parsedText.Substring(index, 14);
                                //int index = 0;
                                //do
                                //{
                                //    index = parsedText.IndexOf("№ ", index + 2);
                                //    if (index >= 0) barcode = parsedText.Substring(index + 1, 20).Trim();
                                //    if (Regex.IsMatch(barcode, "^[0-9]{1,20}$")) matches.Add(barcode);
                                //} while (index != -1);
                                Match firstPageMatches;
                                MatchCollection secondPageMatch;
                                if (j == 1)
                                {
                                    secondPageMatch = Regex.Matches(parsedText, @"18810154\d{12,12}");
                                    foreach (Match item in secondPageMatch)
                                    {
                                        matches.Add(item.Value);
                                    }
                                }
                                else
                                {
                                    firstPageMatches = Regex.Match(parsedText, @"18810154\d{12,12}");
                                    if (firstPageMatches.Success)
                                    {
                                        barcode = firstPageMatches.Value;
                                        matches.Add(firstPageMatches.Value);
                                    }
                                }

                                copy.AddPage(copy.GetImportedPage(reader, i + j));
                            }
                            //if (barcode == string.Empty) throw new Exception($"Не найден штрихкод около страницы {i}.");
                            if (matches.Count < 2) throw new Exception($"В {Path.GetFileName(file)} около страницы {i} не найден один из штрихкодов.");
                            if (!matches.GetRange(1, matches.Count-1).Contains(matches.First()))
                            {
                                throw new Exception($"Файл: {Path.GetFileName(file)} Штрихкоды не совпадают около страницы {i}.");
                            }
                            doc.Close();
                            matches.Clear();
                        }
                        if (File.Exists($"{path}{i}.pdf"))
                        {
                            if (File.Exists($"{path}{barcode}.pdf")) File.Delete($"{path}{barcode}.pdf");
                            File.Move($"{path}{i}.pdf", $"{path}{barcode}.pdf");
                        }
                    }
                }
            }
            sw.Stop();
            Console.WriteLine($"Elapsed: {sw.ElapsedMilliseconds} ms.");
        }
    }
}


//380, 685