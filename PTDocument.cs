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
using System.IO;
using System.Windows;

namespace PTools
{
    public class PTDocument : IDisposable
    {
        private Document _doc;
        private MemoryStream _ms;
        private PdfWriter _writer;
        private PdfContentByte _cb;
        private BaseFont _bf;
        private bool _isClosed = false;

        public int Count { get; private set; }

        public PTDocument(PageSize size)
        {
            Count = 0;
            _ms = new MemoryStream();
            var fmt = size == PageSize.A4 ? iTextSharp.text.PageSize.A4 : iTextSharp.text.PageSize.A3;
            _doc = new Document(fmt);
            _writer = PdfWriter.GetInstance(_doc, _ms);
            _doc.Open();
            _cb = _writer.DirectContent;
            var timesPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Fonts), "arial.TTF");
            _bf = BaseFont.CreateFont(timesPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
        }

        public int AddPagesFromFile(string fileName, int pageForRecno, int x, int y, string recno)
        {
            try
            {
                PdfReader reader = new PdfReader(fileName);
                for (var i = 1; i <= reader.NumberOfPages; i++)
                {
                    PdfImportedPage page = _writer.GetImportedPage(reader, i);
                    _cb.AddTemplate(page, 0, 0);

                    if (pageForRecno == i) PutText(recno, _bf, 5, _cb, x, y);

                    _doc.NewPage();
                    Count++;
                }
                return reader.NumberOfPages;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error in PTools.dll", MessageBoxButton.OK, MessageBoxImage.Error);
                return -1;
            }
        }

        public int AddPagesFromFile(string fileName, int pageForRecno, 
            int x1, int y1, int x2, int y2, int fontSize1, int fontSize2, string recno1, string recno2)
        {
            try
            {
                PdfReader reader = new PdfReader(fileName);
                for (var i = 1; i <= reader.NumberOfPages; i++)
                {
                    PdfImportedPage page = _writer.GetImportedPage(reader, i);
                    _cb.AddTemplate(page, 0, 0);

                    if (pageForRecno == i)
                    {
                        PutText(recno1, _bf, fontSize1, _cb, x1, y1);
                        if (recno2 != "") PutText(recno2, _bf, fontSize2, _cb, x2, y2);
                    }

                    _doc.NewPage();
                    Count++;
                }
                return reader.NumberOfPages;
            }

            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error in PTools.dll", MessageBoxButton.OK, MessageBoxImage.Error);
                return -1;
            }
        }

        public void SaveToFile(string outPath)
        {
            if (_ms.Length != 0)
            {
                _doc.Close();
                _isClosed = true;
                _ms.Flush();
                File.WriteAllBytes(outPath, _ms.ToArray());

                this.Dispose();
            }
            else throw new Exception("MemoryStream пуст, нечего сохранять.");
        }

        public void Dispose()
        {
            if (!_isClosed)
            {
                if (_doc.PageNumber == 0) { _writer.PageEmpty = false; _doc.NewPage(); }
                _doc.Close();
            }
            _ms.Dispose();
            _ms.Close();
            _writer.Close();
            GC.SuppressFinalize(this);
        }

        private void PutText(string text, BaseFont font, int fontSize, PdfContentByte cb, float x, float y)
        {
            cb.SetColorFill(BaseColor.BLACK);
            cb.SetFontAndSize(font, fontSize);
            cb.BeginText();
            cb.ShowTextAligned(PdfContentByte.ALIGN_LEFT, text, x, y, 0);
            cb.EndText();
        }
    }

    public enum PageSize { A4 = 1, A3 = 2 };
}
