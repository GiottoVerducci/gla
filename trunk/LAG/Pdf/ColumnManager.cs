using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;

namespace GLA
{
    public class UnbreakableElement
    {
        public bool DebuggerBreak;
        public List<IElement> Elements { get; private set; }
        public float Height { get; set; }

        public UnbreakableElement(params IElement[] elements)
        {
            Elements = elements.ToList();
        }
    }

    public class ColumnManager
    {
        private readonly Document _doc;
        private readonly PdfWriter _writer;
        private float _columnTop;
        private bool _newPageRequired;
        private float _gutter = 15f;

        private readonly List<UnbreakableElement> _unbreakableElements = new List<UnbreakableElement>();

        public ColumnManager(Document doc, PdfWriter writer)
        {
            _doc = doc;
            _writer = writer;
            _columnTop = doc.Top;
        }

        public void AddSpUnbreakable(params IElement[] elements)
        {
            //foreach (var element in elements)
            //    _unbreakableElements.Add(new UnbreakableElement(element));
            _unbreakableElements.Add(new UnbreakableElement(elements) { DebuggerBreak = true });
        }

        public void AddUnbreakable(params IElement[] elements)
        {
            //foreach (var element in elements)
            //    _unbreakableElements.Add(new UnbreakableElement(element));
            _unbreakableElements.Add(new UnbreakableElement(elements));
        }

        public void Render(int columnCount)
        {
            if (_newPageRequired)
            {
                _columnTop = _doc.Top;
                _doc.NewPage();
                _newPageRequired = false;
            }
            else if (_columnTop <= _doc.Bottom)
                _columnTop = _doc.Top;

            ComputeUnbreakableElementsHeights(columnCount);

            int status = 0;
            //Checks the value of status to determine if there is more text
            //If there is, status is 2, which is the value of NO_MORE_COLUMN

            var columnWidth = GetColumnWidth(columnCount);
            PdfContentByte cb = _writer.DirectContent;
            while (_unbreakableElements.Count > 0)
            {
                int fitCount; // the number of elements that fit
                float height;
                bool mustCreateNewUnbreakableElement = true;
                do
                {
                    height = GetOptimalHeight(columnCount, out fitCount);
                    if (fitCount == 0) // we must break the unbreakable element that doesn't fit on a page
                    {
                        var ue = _unbreakableElements[0];
                        var lastIndex = ue.Elements.Count - 1;
                        var lastElement = ue.Elements[lastIndex];
                        ue.Elements.RemoveAt(lastIndex);
                        if (mustCreateNewUnbreakableElement)
                        {
                            _unbreakableElements.Insert(1, new UnbreakableElement(lastElement));
                            mustCreateNewUnbreakableElement = false;
                        }
                        else
                            _unbreakableElements[1].Elements.Insert(0, lastElement);
                        ComputeUnbreakableElementsHeights(columnCount);
                    }
                } while (fitCount == 0);


                var columnIndex = 0;
                var currentTop = _columnTop;
                bool newColumn = true;
                var bottom = _columnTop - height;
                var ct = new ColumnText(cb);
                if (_unbreakableElements[0].DebuggerBreak)
                    Debugger.Break();
                while (fitCount > 0)
                {
                    var up = _unbreakableElements[0];

                    foreach (var element in up.Elements)
                        ct.AddElement(element);

                    _unbreakableElements.RemoveAt(0);
                    --fitCount;

                    //while (ColumnText.HasMoreText(status))
                    {
                    newCol:
                        if (newColumn && _columnTop < _doc.Top && up.Elements[0] is Paragraph)
                        {
                            currentTop -= ((Paragraph)up.Elements[0]).SpacingBefore;
                        }

                        if (currentTop - bottom < up.Height)
                        {
                            if (columnIndex == columnCount - 1)
                            {
                                columnIndex = 0;
                                _doc.NewPage();
                            }
                            else
                                ++columnIndex;
                            currentTop = _columnTop;
                            newColumn = true;
                            goto newCol;
                        }
                        else
                        {
                            newColumn = false;
                        }

                        //Writing the column
                        var rectangle = GetColumnRectangle(columnWidth, currentTop, bottom, columnIndex);
                        //var rectangle = GetColumnRectangle(columnWidth, _doc.Top, _doc.Top - height, columnIndex);
                        ct.SetSimpleColumn(rectangle);

                        //Needs to be here to prevent app from hanging
                        ct.YLine = currentTop;
                        //Commit the content of the ColumnText to the document
                        //ColumnText.Go() returns NO_MORE_TEXT (1) and/or NO_MORE_COLUMN (2)
                        //In other words, it fills the column until it has either run out of column, or text, or both
                        status = ct.Go();

                        //if (status == ColumnText.NO_MORE_COLUMN && i == columnCount - 1)
                        //{
                        //    i = 0;
                        //    doc.NewPage();
                        //}
                        //else
                        currentTop -= up.Height + (up.Elements.Last() is Paragraph ? ((Paragraph)up.Elements.Last()).SpacingAfter : 0f);
                    }
                }
                if (_unbreakableElements.Count > 0) // more text, needs a new page
                {
                    _columnTop = _doc.Top;
                    _doc.NewPage();
                }
                else
                {
                    _columnTop -= height;
                }
            }

            //bool requiresMoreSpace = true;
            //while (currentBottom > _doc.Bottom && requiresMoreSpace)
            //{
            //    var ctCopy = ColumnText.Duplicate(ct);
            //    //left[3] = right[3] = left2[3] = right2[3] = currentBottom;
            //    requiresMoreSpace = TryRenderColumns(doc, true, ctCopy, status, currentBottom, columnCount);
            //    if (requiresMoreSpace)
            //        currentBottom = Math.Max(doc.Bottom, currentBottom - 10f);
            //}
            //TryRenderColumns(doc, false, ct, status, currentBottom, columnCount);
            //_columnTop = currentBottom;
        }

        private float GetOptimalHeight(int columnCount, out int fitCount)
        {
            bool isTopPage = _columnTop == _doc.Top;
            float bottomSpace = 0f;
            var height = _columnTop - _doc.Bottom;
            if ((fitCount = GetFitCount(height, columnCount, isTopPage, bottomSpace)) == _unbreakableElements.Count)
            {
                do
                {
                    height -= 1f;
                    bottomSpace += 1f;
                } while (GetFitCount(height, columnCount, isTopPage, bottomSpace) == _unbreakableElements.Count);
                height += 1f;
            }
            return height;
        }

        private int GetFitCount(float height, int columnCount, bool isTopPage, float bottomSpace)
        {
            var columnIndex = 0;
            var remainingColumnHeight = height;
            int i = 0;
            var newColumn = true;
            while (i < _unbreakableElements.Count && columnIndex < columnCount)
            {
                var up = _unbreakableElements[i];
                if (up.Height == -1) // since the unbreakable element doesn't fit, we can stop here
                    return i;
                //if (up.Height + up.Paragraph.SpacingAfter > remainingColumnHeight)
                //{
                //    if (up.Height > remainingColumnHeight)
                //    {
                //        remainingColumnHeight = height;
                //        ++columnIndex;
                //    }
                //    else if (i < _unbreakableElements.Count - 1)
                //    {
                //        ++i;
                //        remainingColumnHeight = height;
                //        ++columnIndex;
                //    }
                //    else
                //        return i;
                //}
                var spacingBefore = up.Elements[0] is Paragraph ? ((Paragraph)up.Elements[0]).SpacingBefore : 0f;
                var spacingAfter = up.Elements.Last() is Paragraph ? ((Paragraph)up.Elements.Last()).SpacingAfter : 0f;

                if (up.Height + (isTopPage && newColumn ? 0f : spacingBefore) + (Math.Min(bottomSpace, spacingAfter)) > remainingColumnHeight)
                {
                    remainingColumnHeight = height;
                    ++columnIndex;
                    newColumn = true;
                }
                else
                {
                    //remainingColumnHeight -= up.Height + up.Paragraph.SpacingAfter;
                    remainingColumnHeight -= up.Height + (isTopPage && newColumn ? 0f : spacingBefore) + spacingAfter; // Math.Min(bottomSpace, up.Paragraph.SpacingAfter);
                    ++i;
                    newColumn = false;
                }
            }
            return columnIndex < columnCount
                ? _unbreakableElements.Count
                : i;
        }

        private static int __Count;

        private void ComputeUnbreakableElementsHeights(int columnCount)
        {
            var columnWidth = GetColumnWidth(columnCount);
            PdfContentByte cb = _writer.DirectContent;

            ++__Count;

            Debug.WriteLine("ComputeUnbreakableElementsHeights " + __Count);

            foreach (var up in _unbreakableElements)
            {
                var height = 0f;
                int status;
                bool biggerThanPage = false;
                if(up.DebuggerBreak)
                    Debugger.Break();

                if (up.Elements.Count == 0)
                {
                    up.Height = 0;
                    continue;
                }

                var dichotomicMode = true;
                var lastTryWasGood = false;
                var range = _doc.Top - _doc.Bottom;
                var maxDichotomicSteps = 7;
                var currentStep = 0;
                var previousHeight = 0f;
                do
                {
                    var ct = new ColumnText(cb);
                    foreach (var element in up.Elements)
                        ct.AddElement(element);

                    if (++currentStep == maxDichotomicSteps)
                        dichotomicMode = false;

                    if (dichotomicMode)
                    {
                        if (currentStep > 1 && !lastTryWasGood)
                            height = previousHeight;
                        previousHeight = height;
                        range = (int)(range / 2);
                        height += lastTryWasGood ? -range : range;
                    }
                    else
                    {
                        height += currentStep > 1 && lastTryWasGood ? -1f : 1f;
                        if (height > _doc.Top - _doc.Bottom)
                        {
                            biggerThanPage = true;
                            break;
                        }
                    }

                    //Writing the column
                    var rectangle = GetColumnRectangle(columnWidth, _doc.Top, _doc.Top - height, 0);
                    ct.SetSimpleColumn(rectangle);

                    //Needs to be here to prevent app from hanging
                    ct.YLine = _doc.Top;
                    //Commit the content of the ColumnText to the document
                    //ColumnText.Go() returns NO_MORE_TEXT (1) and/or NO_MORE_COLUMN (2)
                    //In other words, it fills the column until it has either run out of column, or text, or both
                    status = ct.Go(true);

                    if (dichotomicMode)
                        lastTryWasGood = status != 2;

                } while (dichotomicMode || (lastTryWasGood && status != 2) || (!lastTryWasGood && status == 2));

                if (biggerThanPage)
                    up.Height = -1;
                else
                    up.Height = height;
                Debug.WriteLine("  " + up.Height);
            }
        }

        //private List<float[]> GetColumnBorders(float columnWidth, float top, float bottom, int columnIndex)
        //{
        //    var left = _doc.Left + columnIndex * (_gutter + columnWidth);
        //    var right = left + columnWidth;
        //    float[] leftBorder = { left, top, left, bottom };
        //    float[] rightBorder = { right, top, right, bottom };
        //    return new List<float[]> { leftBorder, rightBorder };
        //}

        private Rectangle GetColumnRectangle(float columnWidth, float top, float bottom, int columnIndex)
        {
            var left = _doc.Left + columnIndex * (_gutter + columnWidth);
            var right = left + columnWidth;
            return new Rectangle(left, bottom, right, top);
        }

        public float GetColumnWidth(int columnCount)
        {
            return (_doc.Right - _doc.Left - _gutter * (columnCount - 1)) / columnCount;
        }

        public void EndPage()
        {
            _newPageRequired = true;
        }
    }
}