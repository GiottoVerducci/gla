using System;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace GLA
{
    //http://www.edkhoze.com/itextsharp
    public class PdfFooterHelper : PdfPageEventHelper
    {
        private readonly string _footer;
        private readonly string _headerLeft;
        private readonly string _headerRight;
        private readonly Image _watermarkImage;
        private readonly int _watermarkOpacity;

        // This is the contentbyte object of the writer
        private PdfContentByte _cb;

        // we will put the final number of pages in a template
        private PdfTemplate _template;

        // this is the BaseFont we are going to use for the header / footer
        private BaseFont _bf;

        public Font FooterFont { get; set; }

        public PdfFooterHelper(string footer, string headerLeft, string headerRight, Image watermarkImage, int watermarkOpacity)
        {
            _footer = footer;
            _headerLeft = headerLeft;
            _headerRight = headerRight;
            _watermarkImage = watermarkImage;
            _watermarkOpacity = watermarkOpacity;
        }

        public override void OnOpenDocument(PdfWriter writer, Document document)
        {
            _bf = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.CP1252, BaseFont.NOT_EMBEDDED);
            _cb = writer.DirectContent;
            _template = _cb.CreateTemplate(50, 50);
        }

        public override void OnCloseDocument(PdfWriter writer, Document document)
        {
            base.OnCloseDocument(writer, document);
        }

        public override void OnStartPage(PdfWriter writer, Document document)
        {
            base.OnStartPage(writer, document);
            //float fontSize = 80;
            //float xPosition = 300;
            //float yPosition = 400;
            //float angle = 45;

            if (_watermarkImage == null)
                return;

            PdfContentByte under = writer.DirectContentUnder;
            //under.AddImage(_watermarkImage, );

            _watermarkImage.SetAbsolutePosition((document.PageSize.Width - _watermarkImage.Width) / 2,
                (document.PageSize.Height - _watermarkImage.Height) / 2);

            //_watermarkImage.ScaleToFit(60, 60);
            //_watermarkImage.SetAbsolutePosition((document.PageSize.Width - 60) / 2,
            //    document.PageSize.Height - 90);

            //document.Add(_watermarkImage);
            //under.AddImage(_watermarkImage);

            under.SaveState();

            var graphicsState = new PdfGState { FillOpacity = _watermarkOpacity / 100f };
            under.SetGState(graphicsState);
            under.AddImage(_watermarkImage);
            
            under.RestoreState();

            //BaseFont baseFont = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.EMBEDDED);
            //under.BeginText();
            //under.SetColorFill(BaseColor.LIGHT_GRAY);
            //under.SetFontAndSize(baseFont, fontSize);
            //under.ShowTextAligned(PdfContentByte.ALIGN_CENTER, "test", xPosition, yPosition, angle);
            //under.EndText();
        }

        public override void OnEndPage(PdfWriter writer, Document document)
        {
            base.OnEndPage(writer, document);

            if (document.PageNumber != 1)
            {
                Rectangle pageSize = document.PageSize;
                WriteText(_footer, pageSize.GetLeft(30), pageSize.GetBottom(30));
                WriteText("Page " + document.PageNumber, pageSize.GetRight(30), pageSize.GetBottom(30), Alignment.Right);
                WriteText(_headerLeft, pageSize.GetLeft(30), pageSize.GetTop(30));
                WriteText(_headerRight, pageSize.GetRight(30), pageSize.GetTop(30), Alignment.Right);

                //_cb.AddTemplate(_template, pageSize.GetLeft(30), pageSize.GetBottom(30));
            }
        }

        private void WriteText(string footer, float x, float y, Alignment alignment = Alignment.Left)
        {
            float len = _bf.GetWidthPoint(footer, 8);
            if (alignment == Alignment.Right)
                x -= len;
            else if (alignment == Alignment.Center)
                x -= len / 2;

            _cb.SetRGBColorFill(100, 100, 100);

            _cb.BeginText();
            _cb.SetFontAndSize(_bf, 8);
            _cb.SetTextMatrix(x, y);
            _cb.ShowText(footer);
            _cb.EndText();
        }

        private enum Alignment
        {
            Left,
            Right,
            Center
        }
    }
}