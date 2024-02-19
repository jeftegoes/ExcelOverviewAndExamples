# Excel Overview and Examples <!-- omit in toc -->

## Contents <!-- omit in toc -->

- [1. References](#1-references)
  - [1.1. Relative Cell References:](#11-relative-cell-references)
  - [1.2. Absolute Cell References](#12-absolute-cell-references)
  - [1.3. Mixed Cell References](#13-mixed-cell-references)
  - [1.4. Summary](#14-summary)
- [2. Order of Operations](#2-order-of-operations)
- [3. Frequently used shortcuts](#3-frequently-used-shortcuts)

# 1. References

- In Excel, cell references can be either **relative** or **absolute**, and understanding the difference between them is crucial for creating formulas that behave as expected when copied or filled across multiple cells.

## 1.1. Relative Cell References:

- When you create a formula with relative references and then copy or fill that formula to other cells, the references adjust relative to the position of the formula.
- For example, if you have a formula `=A1+B1` in cell C1, and you copy it to cell C2, the formula will automatically adjust to `=A2+B2`.
  - This is because the formula references cells relative to its own position.

## 1.2. Absolute Cell References

- Absolute references, on the other hand, **do not change when you copy the formula to other cells**.
- They always refer to a specific cell, no matter where the formula is located. In Excel, you indicate an absolute reference by placing a dollar sign `$` before the column letter and/or row number.
- For instance, if you have a formula `=A$1+B$1` in cell C1, and you copy it to cell C2, the formula will remain as `=A$1+B$1`.
  - The dollar sign locks the reference to the specified cell.

## 1.3. Mixed Cell References

- Excel also allows for mixed references, where either the row or column is absolute while the other is relative.
- For example, `$A1` would lock the column reference to column A while allowing the row reference to adjust, and `A$1` would lock the row reference to row 1 while allowing the column reference to adjust.

## 1.4. Summary

- Relative cell references change based on the relative position of the formula.
- Absolute cell references remain fixed.
- Mixed cell references lock either the row or column while allowing the other to adjust.
- Choosing the appropriate type of reference depends on the behavior you want from your formulas when copied or filled to other cells.

[Example](Examples/RelativeVersusAbsoluteCellReferencesInFormulas.xlsx)

# 2. Order of Operations

- Follows the acronym **PEMDAS**:
  - **Parentheses:** Operations within parentheses are evaluated first, allowing you to control the sequence of calculations.
  - **Exponents:** Exponential operations (raising a number to a power) are evaluated next.
  - **Multiplication and Division:** Multiplication (\*) and division (/) are evaluated next, from left to right.
  - **Addition and Subtraction:** Addition (+) and subtraction (-) are evaluated last, also from left to right.

[Example](Examples/Order%20of%20Operation.xlsx)

# 3. Frequently used shortcuts

- This table itemizes the most frequently used shortcuts in Excel for Windows.

| To do this                 | Press            |
| -------------------------- | ---------------- |
| Close a spreadsheet        | Ctrl+W           |
| Open a spreadsheet         | Ctrl+O           |
| Go to the Home tab         | Alt+H            |
| Save a spreadsheet         | Ctrl+S           |
| Copy                       | Ctrl+C           |
| Paste                      | Ctrl+V           |
| Undo                       | Ctrl+Z           |
| Remove cell contents       | Delete key       |
| Choose a fill color        | Alt+H, H         |
| Cut                        | Ctrl+X           |
| Go to Insert tab           | Alt+N            |
| Bold                       | Ctrl+B           |
| Center align cell contents | Alt+H, A, then C |
| Go to Page Layout tab      | Alt+P            |

| To do this                | Press                     |
| ------------------------- | ------------------------- |
| Go to Data tab            | Alt+A                     |
| Go to View tab            | Alt+W                     |
| Open context menu         | Shift+F10, or Context Key |
| Add borders               | Alt+H, B                  |
| Delete column             | Alt+H,D, then C           |
| Go to Formula tab         | Alt+M                     |
| Hide the selected rows    | Ctrl+9                    |
| Hide the selected columns | Ctrl+0                    |

| To do this                                                                                                                                                                              | Press            |
| --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ---------------- |
| Move to the previous cell in a worksheet or the previous option in a dialog box.                                                                                                        | Shift+Tab        |
| Move one cell up in a worksheet.                                                                                                                                                        | Up Arrow key     |
| Move one cell down in a worksheet.                                                                                                                                                      | Down Arrow key   |
| Move one cell left in a worksheet.                                                                                                                                                      | Left Arrow key   |
| Move one cell right in a worksheet.                                                                                                                                                     | Right Arrow key  |
| Move to the edge of the current data region in a worksheet.                                                                                                                             | Ctrl+arrow key   |
| Enter End mode, move to the next nonblank cell in the same column or row as the active cell, and turn off End mode. If the cells are blank, move to the last cell in the row or column. | End, arrow key   |
| Move to the last cell on a worksheet, to the lowest used row of the rightmost used column.                                                                                              | Ctrl+End         |
| Extend the selection of cells to the last used cell on the worksheet (lower-right corner).                                                                                              | Ctrl+Shift+End   |
| Move to the cell in the upper-left corner of the window when Scroll Lock is turned on.                                                                                                  | Home+Scroll Lock |
| Move to the beginning of a worksheet.                                                                                                                                                   | Ctrl+Home        |

| To do this                                                                                            | Press          |
| ----------------------------------------------------------------------------------------------------- | -------------- |
| Move one screen down in a worksheet.                                                                  | Page Down      |
| Move to the next sheet in a workbook.                                                                 | Ctrl+Page Down |
| Move one screen to the right in a worksheet.                                                          | Alt+Page Down  |
| Move one screen up in a worksheet.                                                                    | Page Up        |
| Move one screen to the left in a worksheet.                                                           | Alt+Page Up    |
| Move to the previous sheet in a workbook.                                                             | Ctrl+Page Up   |
| Move one cell to the right in a worksheet. Or, in a protected worksheet, move between unlocked cells. | Tab            |
