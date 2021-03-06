ClosedXML.MergedCellsHeightCalculator
===

Unfortunately Excel's WordWrap-feature does not work on merged cells - Because of that I've created an extension for the IXLRange class which allows you to calculate the row-height that is needed to display the content in word-wrap mode.

How does it work?
I'm basically using a refactored version of XLRow's AdjustToContents()-method to get the height that would be needed to display the whole content in one row (without wordwrap),
and XLColumn's AdjustToContents()-method to get the width that would be needed. Then I'm accumulating the width of the merged cells. All these 3 variables are used to calculate the row-height:

```csharp
//the content will be in the first of the merged cells
var firstCell = range.FirstCell();

// calculate the perfect height that would be needed to display the content in one row (without word-wrap)
// -- refactored and changed version of XLRow's AdjustToContents()-method
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
int roundedMultiplier = (int) (heightMultiplier + 0.9);

//multiply the needed height with the multiplier - the multiplier should be at least 1 though
var calculatedHeight = neededHeightForOneRow * (roundedMultiplier >= 1 ? roundedMultiplier : 1);
```

If you want to adjust the row to the needed height automatically, you can use the IXLRow HeightAutoFit(bool allowHeightDecrease)-extension.
This function will search for all merged cells that are located in this row and calculate the height for these ranges.
The greatest height will be used as the new row height if it's either greater than the current height or if the user allowed the row's height to decrease.

However there's one problem. If a single cell (with activated wordwrap) has a greater calculated height than all of the merged cells, it would be better to let Excel
calculate the row height. The problem is that once a custom row height is set, Excel won't override this height. We can't change this behaviour with an extension. It's possible to set the bool "CustomHeight" of an OpenXML row in the source code of ClosedXML's save()-function though.

###Usage
---

```csharp

var range = sheet.Range("A1:C1");
range.Merge();
range.Style.Alignment.WrapText = true;
range.Value = "ehsjsadasdasdgrsoasdasdbtiejhz34908zug3489tqu32452345rt42r";

var row = sheet.Row(1);
//you can either use the calculated height of one range..
row.Height = range.CalculateMergedCellWordWrapHeight();
//.. or let the row calculate its greatest height and let it use that automatically
row.HeightAutoFit(true);

```



