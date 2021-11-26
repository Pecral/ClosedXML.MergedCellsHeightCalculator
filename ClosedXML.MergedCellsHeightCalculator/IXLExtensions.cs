using System;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using System.Drawing;

namespace ClosedXML.MergedCellsHeightCalculator
{
    public static class IXLExtensions
    {
        /// <summary>
        /// Set the perfect row height..
        /// </summary>
        /// <param name="row"></param>
        /// <param name="allowHeightDecrease">Specify whether the row should be allowed to be shrinked lower than the current height</param>
        /// <param name="calculateSingleCellHeight">If you want use the calculated height even for single cells (rather than letting Excel calculate it), set it to true.
        /// This is useful if the row height already has a custom height (Excel wouldn't override this height)</param>
        public static void HeightAutoFit(this IXLRow row, bool allowHeightDecrease = false, bool calculateSingleCellHeight = false)
        {
            //get worksheet
            IXLWorksheet sheet = row.Worksheet;

            int rowNumber = row.RowNumber();
            //get all the merged cells of this row
            IEnumerable<IXLRange> mergedCells = sheet.MergedRanges.Where(x => x.FirstRow().RowNumber() == rowNumber && x.LastRow().RowNumber() == rowNumber);

            double maxHeight = 0;

            foreach (var range in mergedCells)
            {
                //check whether wordwrap mode is activated for either the first cell or the entire range
                bool calculateHeight = range.Style.Alignment.WrapText || range.FirstCell().Style.Alignment.WrapText;

                if (calculateHeight)
                {
                    double neededRowHeight = range.CalculateMergedCellWordWrapHeight();

                    if (neededRowHeight > maxHeight)
                    {
                        maxHeight = neededRowHeight;
                    }
                }
            }

            //check whether any single cell with activated wordwrap that needs more height than
            //the merged cells exists
            var cellsWithWordWrap = row.CellsUsed().Where(c => c.IsMerged() == false && c.Style.Alignment.WrapText);
            foreach (var cell in cellsWithWordWrap)
            {
                //create range for cell (a little bit dirty)
                var cellRange = cell.AsRange();
                double neededRowHeight = cellRange.CalculateMergedCellWordWrapHeight();

                //We could return here if the calculated height for the single cell is higher than the height for the merged cells,
                //because Excel's WordWrap would be more precise than our function. The problem is that once a custom row height is set,
                //Excel won't override this height. We can't change this behaviour with an extension,
                //it's possible to set the bool "CustomHeight" in ClosedXML's source code though.
                //Because of that we will use our own height here.
                if (neededRowHeight > maxHeight)
                {
                    if (calculateSingleCellHeight)
                    {
                        maxHeight = neededRowHeight;
                    }
                    else
                    {
                        //let excel calculate it
                        return;
                    }
                }
            }

            //set calculated height to current height if its 0
            maxHeight = maxHeight > 0 ? maxHeight : row.Height;

            //if the new height is lower than the current height and the user doesn't allow a decrease, we will do nothing
            if (allowHeightDecrease || maxHeight > row.Height)
            {
                row.Height = maxHeight;
            }
        }

        /// <summary>
        /// Calculate the row height that is needed to display the content of a merged cell with activated wordwrap
        /// </summary>
        /// <param name="range">The range across a merged cell within a single row</param>
        /// <returns></returns>
        public static double CalculateMergedCellWordWrapHeight(this IXLRange range)
        {
            //the range should only be across a single row - return the height of the first row if it's across multiple rows
            if (range.FirstRow().RowNumber() != range.LastRow().RowNumber())
            {
                return range.FirstRow().WorksheetRow().Height;
            }

            //the content will be in the first of the merged cells
            var firstCell = range.FirstCell();

            // calculate the perfect height that would be needed to display the content in one row (without word-wrap)
            // -- this is a refactored and changed version of XLRow's AdjustToContents()-method
            double neededHeightForOneRow = firstCell.CalculateContentHeightWithoutWrap();

            // calculate the perfect width that would be needed to display the content
            // -- refactored and changed version of XLColumn's AdjustToContents()-method
            double neededWidth = firstCell.CalculateContentWidth();

            //accumulated width of all cells
            double widthOfAllCells = 0;
            range.Columns().ForEach(c => widthOfAllCells += c.WorksheetColumn().Width);

            //how many times should we multiply the height
            double heightMultiplier = neededWidth / widthOfAllCells;

            //the number is rounded up because we can only use a row as a whole (of course)
            //we're adding 0.9 instead of 1 because the heightMultiplier is always a little bit to heigh
            int roundedMultiplier = (int)(heightMultiplier + 0.9);

            //multiply the needed height with the multiplier - the multiplier should be at least 1 though
            return neededHeightForOneRow * (roundedMultiplier >= 1 ? roundedMultiplier : 1);
        }

        /// <summary>
        /// Calculate the perfect height that would be needed to display all text in one row (without word-wrap)
        /// This function will respect different font sizes as well as font types.
        /// --- refactored and changed version of XLRow's AdjustToContents()-method
        /// </summary>
        /// <param name="cell">The cell in which the content is placed</param>
        /// <returns>The height that is needed to display the whole content (without wordwrap)</returns>
        public static double CalculateContentHeightWithoutWrap(this IXLCell cell)
        {
            var fontCache = new Dictionary<IXLFontBase, Font>();
            Double rowMaxHeight = 0;

            Int32 textRotation = cell.Style.Alignment.TextRotation;
            if (cell.HasRichText || textRotation != 0 || cell.Value.ToString().Contains(Environment.NewLine))
            {
                //fonts with their content line by line
                var fontContentList = new List<KeyValuePair<IXLFontBase, string>>();

                //fonts with their content as a whole, newlines not considered
                var contentToIterate = new List<KeyValuePair<IXLFontBase, string>>();

                //if the value is a rich text, we have to iterate through the different rich texts because they could have different fonts
                if (cell.HasRichText)
                {
                    cell.RichText.ForEach(rt => contentToIterate.Add(new KeyValuePair<IXLFontBase, String>(rt, rt.Text)));
                }
                else
                {
                    contentToIterate.Add(new KeyValuePair<IXLFontBase, String>(cell.Style.Font, cell.GetFormattedString()));
                }

                //iterate through the content, divide it by line and add the font to the font content list
                contentToIterate.ForEach(content =>
                {
                    IXLFontBase font = content.Key;
                    //content that could contain multiple lines
                    string text = content.Value;

                    //split by new line
                    var arr = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
                    Int32 arrCount = arr.Count();
                    for (Int32 i = 0; i < arrCount; i++)
                    {
                        String s = arr[i];
                        if (i < arrCount - 1)
                        {
                            s += Environment.NewLine;
                        }
                        fontContentList.Add(new KeyValuePair<IXLFontBase, String>(font, s));
                    }
                });

                //get font with the highest height
                Double maxLongCol = fontContentList.Max(kp => kp.Value.Length);
                Double maxHeightCol = fontContentList.Max(kp => kp.Key.GetHeight(fontCache));
                Int32 lineCount = fontContentList.Count(kp => kp.Value.Contains(Environment.NewLine)) + 1;

                //true when rotation is horizontal
                if (textRotation == 0)
                    rowMaxHeight = maxHeightCol * lineCount;
                else
                {
                    if (textRotation == 255)
                        rowMaxHeight = maxLongCol * maxHeightCol;
                    else
                    {
                        Double rotation;
                        if (textRotation == 90 || textRotation == 180 || textRotation == 255)
                            rotation = 90;
                        else
                            rotation = textRotation % 90;

                        rowMaxHeight = (rotation / 90.0) * maxHeightCol * maxLongCol * 0.5;
                    }
                }
            }
            //we can use the font's height if the text is horizontal and does not contain any rich text or new line
            else
            {
                rowMaxHeight = cell.Style.Font.GetHeight(fontCache);
            }

            //set height to current height if something went wrong
            if (rowMaxHeight <= 0)
            {
                rowMaxHeight = cell.WorksheetRow().Height;
            }

            //dispose fonts
            foreach (IDisposable font in fontCache.Values)
            {
                font.Dispose();
            }

            return rowMaxHeight;
        }

        /// <summary>
        /// Calculate the perfect width that would be needed to display all text
        /// This function will respect different font sizes as well as font types.
        /// --- refactored and changed version of XLColumn's AdjustToContents()-method
        /// </summary>
        /// <param name="cell">The cell in which the content is placed</param>
        /// <returns>The width that is needed to display the whole content</returns>
        public static double CalculateContentWidth(this IXLCell cell)
        {
            var fontCache = new Dictionary<IXLFontBase, Font>();
            Double colMaxWidth = 0;

            Int32 textRotation = cell.Style.Alignment.TextRotation;
            if (cell.HasRichText || textRotation != 0 || cell.Value.ToString().Contains(Environment.NewLine))
            {
                //fonts with their content line by line
                var fontContentList = new List<KeyValuePair<IXLFontBase, string>>();

                //fonts with their content as a whole, newlines not considered
                var contentToIterate = new List<KeyValuePair<IXLFontBase, string>>();

                //if the value is a rich text, we have to iterate through the different rich texts because they could have different fonts
                if (cell.HasRichText)
                {
                    cell.RichText.ForEach(rt => contentToIterate.Add(new KeyValuePair<IXLFontBase, String>(rt, rt.Text)));
                }
                else
                {
                    contentToIterate.Add(new KeyValuePair<IXLFontBase, String>(cell.Style.Font, cell.GetFormattedString()));
                }

                //iterate through the content, divide it by line and add the font to the font content list
                contentToIterate.ForEach(content =>
                {
                    IXLFontBase font = content.Key;
                    //content that could contain multiple lines
                    string text = content.Value;

                    //split by new line
                    var arr = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

                    Int32 arrCount = arr.Count();
                    for (Int32 i = 0; i < arrCount; i++)
                    {
                        String s = arr[i];
                        if (i < arrCount - 1)
                        {
                            s += Environment.NewLine;
                        }

                        fontContentList.Add(new KeyValuePair<IXLFontBase, String>(font, s));
                    }
                });

                #region foreach (var fontContent in fontContentList)

                Double runningWidth = 0;
                Boolean rotated = false;
                Double maxLineWidth = 0;
                Int32 lineCount = 1;
                foreach (KeyValuePair<IXLFontBase, string> fontContent in fontContentList)
                {
                    var f = fontContent.Key;
                    String formattedString = fontContent.Value;

                    Int32 newLinePosition = formattedString.IndexOf(Environment.NewLine);
                    if (textRotation == 0)
                    {
                        #region if (newLinePosition >= 0)

                        if (newLinePosition >= 0)
                        {
                            if (newLinePosition > 0)
                                runningWidth += f.GetWidth(formattedString.Substring(0, newLinePosition), fontCache);

                            if (runningWidth > colMaxWidth)
                                colMaxWidth = runningWidth;

                            runningWidth = newLinePosition < formattedString.Length - 2
                                               ? f.GetWidth(formattedString.Substring(newLinePosition + 2), fontCache)
                                               : 0;
                        }
                        else
                            runningWidth += f.GetWidth(formattedString, fontCache);

                        #endregion
                    }
                    else
                    {
                        #region if (textRotation == 255)

                        if (textRotation == 255)
                        {
                            if (runningWidth <= 0)
                                runningWidth = f.GetWidth("X", fontCache);

                            if (newLinePosition >= 0)
                                runningWidth += f.GetWidth("X", fontCache);
                        }
                        else
                        {
                            rotated = true;
                            Double vWidth = f.GetWidth("X", fontCache);
                            if (vWidth > maxLineWidth)
                                maxLineWidth = vWidth;

                            if (newLinePosition >= 0)
                            {
                                lineCount++;

                                if (newLinePosition > 0)
                                    runningWidth += f.GetWidth(formattedString.Substring(0, newLinePosition), fontCache);

                                if (runningWidth > colMaxWidth)
                                    colMaxWidth = runningWidth;

                                runningWidth = newLinePosition < formattedString.Length - 2
                                                   ? f.GetWidth(formattedString.Substring(newLinePosition + 2), fontCache)
                                                   : 0;
                            }
                            else
                                runningWidth += f.GetWidth(formattedString, fontCache);
                        }

                        #endregion
                    }
                }

                #endregion

                if (runningWidth > colMaxWidth)
                    colMaxWidth = runningWidth;

                #region if (rotated)

                if (rotated)
                {
                    Int32 rotation;
                    if (textRotation == 90 || textRotation == 180 || textRotation == 255)
                        rotation = 90;
                    else
                        rotation = textRotation % 90;

                    //degree to radian
                    Double r = Math.PI * rotation / 180.0;

                    colMaxWidth = (colMaxWidth * Math.Cos(r)) + (maxLineWidth * lineCount);
                }

                #endregion
            }
            //we can use the font's width if the text is horizontal and does not contain any rich text or new line
            else
            {
                colMaxWidth = cell.Style.Font.GetWidth(cell.GetFormattedString(), fontCache);
            }

            //return current width if something went wrong
            if (colMaxWidth <= 0)
            {
                colMaxWidth = cell.WorksheetColumn().Width;
            }

            //dispose fonts
            foreach (IDisposable font in fontCache.Values)
            {
                font.Dispose();
            }

            return colMaxWidth;
        }
    }
}
