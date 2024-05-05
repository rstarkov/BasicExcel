namespace BasicExcel;

public static class XlUtil
{
    public static (int row, int col) CellRef(string cell)
    {
        int row = 0;
        int col = 0;
        foreach (var c in cell)
        {
            if (char.IsDigit(c))
                row = row * 10 + (c - '0');
            else
                col = col * 26 + (c - 'A' + 1);
        }
        return (row, col);
    }

    public static string CellRef(int row, int col)
    {
        string cell = "";
        while (col > 0)
        {
            col--;
            cell = (char)('A' + col % 26) + cell;
            col /= 26;
        }
        return cell + row;
    }

    #region Test
    /*
    void test(int col, int row, string exp)
    {
        if (XlUtil.CellRef(row, col) != exp) throw new Exception();
        var (r, c) = XlUtil.CellRef(exp);
        if (r != row || c != col) throw new Exception();
    }
    test(1, 1, "A1");
    test(1, 2, "A2");
    test(1, 1999, "A1999");
    test(26, 1, "Z1");
    test(26, 2, "Z2");
    test(26, 1999, "Z1999");
    test(27, 1, "AA1");
    test(27, 2, "AA2");
    test(27, 1999, "AA1999");
    test(100, 1, "CV1");
    test(100, 2, "CV2");
    test(100, 1999, "CV1999");
    test(702, 1, "ZZ1");
    test(702, 2, "ZZ2");
    test(702, 1999, "ZZ1999");
    test(703, 1, "AAA1");
    test(703, 2, "AAA2");
    test(703, 1999, "AAA1999");
    test(16384, 1, "XFD1");
    test(16384, 2, "XFD2");
    test(16384, 1999, "XFD1999");
    test(16384, 1048576, "XFD1048576");
     */
    #endregion
}
