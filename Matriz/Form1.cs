using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using BOMCore;

namespace Matriz
{
    public partial class BOMAnalysisForm : Form
    {

        public BOMAnalysisForm()
        {
            InitializeComponent();
        }

        private void BOMAnalysisForm_Load(object sender, EventArgs e)
        {

        }

        public bool odi_ValidateBOMFormat(object[,] arr, int t = ODI.odiMatrixBOM)
        {
            bool v = true;
            switch (t)
            {
                case ODI.odiMatrixBOM:
                    if (arr.GetUpperBound(0) >= 5 && arr.GetUpperBound(1) >= ODI.ColRemark[t])// array index 夠大，裡面判斷式才不會out of index
                    {
                        int aaa = ODI.ColRemark[t];
                        v &= (Convert.ToString(arr.GetValue(ODI.RowBOMStart - 1, ODI.ColNum[t])).Equals("Item")) ? true : false; // Check 欄位標題
                        v &= (Convert.ToString(arr.GetValue(ODI.RowBOMStart - 1, ODI.ColDescription[t])).Equals("Description")) ? true : false;
                        v &= (Convert.ToString(arr.GetValue(ODI.RowBOMStart - 1, ODI.ColMFGPN[t])).Equals("Supplier PN")) ? true : false;
                        v &= (Convert.ToString(arr.GetValue(ODI.RowBOMStart - 1, ODI.ColQty[t])).Equals("Qty")) ? true : false;
                        v &= (Convert.ToString(arr.GetValue(ODI.RowBOMStart - 1, ODI.ColTotalSet[t])).Equals("Total Set")) ? true : false;
                        v &= (Convert.ToString(arr.GetValue(ODI.RowBOMStart - 2, ODI.ColMatrixA[t])).Equals("Model A")) ? true : false;
                        v &= (Convert.ToString(arr.GetValue(ODI.RowBOMStart - 1, ODI.ColRemark[t])).Equals("Remark")) ? true : false; // Check 最後一欄                        
                    }
                    else
                        v &= false;
                    break;
                case ODI.odiMfgBOM:
                    if (arr.GetUpperBound(0) >= 5 && arr.GetUpperBound(1) >= ODI.ColRemark[t])// array index 夠大，裡面判斷式才不會out of index
                    {
                        v &= (Convert.ToString(arr.GetValue(ODI.RowBOMStart - 1, ODI.ColNum[t])).Equals("Item")) ? true : false; // Check 欄位標題
                        v &= (Convert.ToString(arr.GetValue(ODI.RowBOMStart - 1, ODI.ColDescription[t])).Equals("Description")) ? true : false;
                        v &= (Convert.ToString(arr.GetValue(ODI.RowBOMStart - 1, ODI.ColMFGPN[t])).Equals("Supplier PN")) ? true : false;
                        v &= (Convert.ToString(arr.GetValue(ODI.RowBOMStart - 1, ODI.ColQty[t])).Equals("Qty")) ? true : false;
                        v &= (Convert.ToString(arr.GetValue(ODI.RowBOMStart - 1, ODI.ColCCL[t])).Equals("CCL")) ? true : false;
                        v &= (Convert.ToString(arr.GetValue(ODI.RowBOMStart - 1, ODI.ColRemark[t])).Equals("Remark")) ? true : false; // Check 最後一欄                        
                    }
                    else
                        v &= false;
                    break;
                case ODI.odiCostBOM:
                    break;
                default:
                    break;
            }
            return v;
        }
        public bool odi_ValidateBOMSheets(Excel.Sheets wkshts, int type = ODI.odiMatrixBOM)
        {
            List<string> sheetlist = ODI.get_SheetList(type);

            foreach (Excel.Worksheet w in wkshts)
                sheetlist.Remove(w.Name);

            return (sheetlist.Count > 0) ? false : true;
        }

        internal string Create_BOM(List<Group> BOM, object[,] wstarr, int type = ODI.odiMatrixBOM)
        {
            Group g = new Group();
            if (odi_ValidateBOMFormat(wstarr, type))
            {
                for (int irow = ODI.RowBOMStart; irow <= wstarr.GetUpperBound(0);)
                {
                    g = Create_BOM_getgroupfromarray(ref irow, wstarr, type);
                    BOM.Add(g);
                }
            }
            else
            {
                //Globals.ThisAddIn.odiMsgOut("BOM format Mismatch![List]");
                MessageBox.Show("BOM format Mismatch![List] \n Check FAQ page in Help menu.");
                return null;
            }
            return g.getBomKey();
        }

        internal string Create_BOM(Dictionary<string, Group> BOM, object[,] wstarr, int type = ODI.odiMatrixBOM)
        {
            Group g = new Group();
            if (!odi_ValidateBOMFormat(wstarr, type))
            {
                //Globals.ThisAddIn.odiMsgOut("BOM format Mismatch![Dictionary]");
                MessageBox.Show("BOM format Mismatch![Dictionary] \n \n Check FAQ page in Help menu.");
                return null;
            }
            else
            {
                for (int irow = ODI.RowBOMStart; irow <= wstarr.GetUpperBound(0);)
                {
                    g = Create_BOM_getgroupfromarray(ref irow, wstarr, type);

                    //先確認目前BOM中是否有相同的group, 有就合併AVL
                    if (BOM.TryGetValue(g.getKey(), out Group tempg))
                    {
                        tempg.combineGroup(g);
                    }
                    else //沒有就增加新group
                    {
                        BOM.Add(g.getKey(), g);
                    }
                }

            }
            return g.getBomKey();
        }

        private Group Create_BOM_getgroupfromarray(ref int irow, object[,] wstarr, int type)
        {
            Group g = new Group();

            // 若為 Main source
            if (wstarr[irow, ODI.ColNum[type]] != null)
            {
                g.RetrieveData(irow, wstarr, type);
                irow++;
            }

            //加入 Second source
            while (irow <= wstarr.GetUpperBound(0) && wstarr[irow, ODI.ColNum[type]] == null)
            {
                //Todo:[Err] 增加錯誤處理，讀取AVL前確認main source是否已經存在: BOM格式錯誤，2nd source xxx 不屬於 group

                if (wstarr[irow, ODI.ColHHPN[type]] != null)
                {
                    Item item = new Item(irow, wstarr, type);
                    g.AVL.Add(item.getKey(), item);
                }
                irow++;
            }
            g.setAVL(); //AVL填入Qty
            return g;
        }






        Excel.Application initailExcel()
        {
            bool flag = false;
            Excel.Application _Excel = null;

            /* 檢查PC有無Excel在執行 */
            foreach (var item in Process.GetProcesses())
            {
                if (item.ProcessName == "EXCEL")
                {
                    flag = true;
                    break;
                }
            }

            if (!flag)
            {
                _Excel = new Excel.Application();
            }
            else
            {
                object obj = Marshal.GetActiveObject("Excel.Application");//引用已在執行的Excel
                _Excel = obj as Excel.Application;
            }

            _Excel.Visible = false; //設定隱藏背景excel視窗
            _Excel.UserControl = false;
            _Excel.DisplayAlerts = false;
            return _Excel;
        }

        private void btn_Go_Click(object sender, EventArgs e)
        {
            btn_Go.Enabled = false;

            OpenFileDialog ofd = new OpenFileDialog();
            ofd.InitialDirectory = @"C:\";
            ofd.Filter = "Excel files (*.xls;*.xlsx)|*.xls;*.xlsx";
            ofd.ValidateNames = true;
            ofd.Multiselect = false;
            if (ofd.ShowDialog(this) == DialogResult.Cancel)
            {
                btn_Go.Enabled = true;
                return;
            }

            Excel.Application oXL = initailExcel();

            Excel.Workbook wkb; //要開啟的BOM 檔案
            Excel.Workbook wkbout; //要輸出的BOM 檔案
            List<String> shList = ODI.get_SheetList(ODI.odiMatrixBOM); // 取得要處理的sheet name
            shList.Reverse();
            //Dictionary<string, Group> GBOM = new Dictionary<string, Group>(); // 總BOM表 Great BOM

            int fileidx = 1; //檔案計數器 

            this.Status1.Text = "Opening BOM " + Convert.ToString(fileidx) + " ...";
    try
    {
        //Globals.ThisAddIn.odiDbgOut("Opening Workbook: " + f);
        wkb = oXL.Workbooks.Open(ofd.FileName, ReadOnly: true);
        oXL.Visible = false;
            }
    catch (Exception ex)
    {
        //Globals.ThisAddIn.odiMsgOut("[Error] Fail Opening Workbook: " + f);
        //Globals.ThisAddIn.odiMsgOut(ex.Message);
        MessageBox.Show("Error: Could not read file from disk. Original error: " + ex.Message);
        btn_Go.Enabled = true;
        return;
    }

            try
            {
                this.Status1.Text = "Reading BOM " + Convert.ToString(fileidx) + " ...";
                //Globals.ThisAddIn.odiDbgOut("Validating BOM format... ");
                if (!this.odi_ValidateBOMSheets(wkb.Sheets, ODI.odiMfgBOM))
                {
                    //Globals.ThisAddIn.odiMsgOut("[Warning] Base BOM Format mismatch [sheets missing] : " + f);
                    MessageBox.Show("BOM Format mismatch [sheets missing] : " + ofd.FileName + "\n Check FAQ page in Help menu.");
                    btn_Go.Enabled = true;
                    return;
                }

                wkbout = oXL.Workbooks.Add(); //產生輸出workbook
                oXL.Visible = false;

                wkbout.Worksheets[1].Name = "WAIT_ForDELEte";

                foreach (string sh in shList)
                {
                    List<Group> GBOM = new List<Group>();

                    /***讀取BOM*/
                    object[,] arr = null;
                    try
                    {
                        wkb.Worksheets[sh].Unprotect(string.Empty);
                        //Unprotect suceeded, 沒有密碼保護，採用陣列讀取
                        arr = wkb.Worksheets[sh].Range["A1"].CurrentRegion.Value2;
                    }
                    catch
                    {
                        //SMT/PTH/BOTTOM因為有密碼保護，無法用陣列讀取，所以只好一格一格來
                        int iRowCount = wkb.Worksheets[sh].UsedRange.Rows.Count;
                        int iColCount = wkb.Worksheets[sh].UsedRange.Columns.Count;
                        Excel.Range range;
                        arr = (object[,])Array.CreateInstance(typeof(object), new int[] { iRowCount, iColCount }, new int[] { 1, 1 });

                        for (int iRow = 1; iRow <= iRowCount; iRow++)
                        {
                            this.Status1.Text = "Reading " + sh + " " + iRow + " / " + iRowCount;
                            for (int iCol = 1; iCol <= iColCount; iCol++)
                            {
                                range = (Excel.Range)wkb.Worksheets[sh].Cells[iRow, iCol];
                                //arr[iRow, iCol] = (range.Value2 == null) ? null : range.Text.ToString();
                                arr[iRow, iCol] = range.Value2;
                            }
                        }
                    }


                    string bomkey = this.Create_BOM(GBOM, arr, ODI.odiMfgBOM);
                    if (String.IsNullOrEmpty(bomkey))
                    {
                        MessageBox.Show("[Error:] BOM Create fail (null bomkey) \n Check FAQ page in Help menu.");
                        btn_Go.Enabled = true;
                        return;
                    }
                    //Globals.ThisAddIn.odiDbgOut("BOM Created: " + bomkey);

                    this.Status1.Text = "Reading BOM " + Convert.ToString(fileidx) + " " + bomkey;


                    /***準備資料 ***/
                    List<string[]> outlist = new List<string[]>(); //每個row
                    List<int> GroupRowCountList = new List<int>(); //紀錄每個group第一行的位置

                    /*** 篩選內容，寫入outlist***/
                    var query = from data in GBOM
                                where data.CCL == "Y"
                                select data;


                    this.Status1.Text = "Picking up CCLs for you :)";
                    GroupRowCountList.Add(0);
                    int outColCount = ODI.ColRemark[ODI.odiMatrixBOM];
                    foreach (var q in query)
                    {
                        //每個q為一個group
                        //Globals.ThisAddIn.odiDbgOut(q.Value.MFGPN + " !data.Value.Qty.Contains(0): " + !q.Value.Qty.Contains(0));
                        string[] row = new string[outColCount]; //單行row
                        row[ODI.ColNum[ODI.odiMatrixBOM] - 1] = Convert.ToString(q.Num);
                        row[ODI.ColHHPN[ODI.odiMatrixBOM] - 1] = Convert.ToString(q.HHPN);
                        row[ODI.ColDescription[ODI.odiMatrixBOM] - 1] = Convert.ToString(q.Description);
                        row[ODI.ColMFG[ODI.odiMatrixBOM] - 1] = Convert.ToString(q.MFG);
                        row[ODI.ColMFGPN[ODI.odiMatrixBOM] - 1] = Convert.ToString(q.MFGPN);
                        row[ODI.ColQty[ODI.odiMatrixBOM] - 1] = Convert.ToString(q.Qty[0]);
                        row[ODI.ColLocation[ODI.odiMatrixBOM] - 1] = Convert.ToString(q.Location);
                        row[ODI.ColRemark[ODI.odiMatrixBOM] - 1] = Convert.ToString(q.Remark);

                        outlist.Add(row);

                        if (q.AVL.Count() > 0) //若有AVL，寫入下一個row
                            foreach (var a in q.AVL)
                            {
                                string[] rowa = new string[outColCount]; //單行row
                                rowa[ODI.ColNum[ODI.odiMatrixBOM] - 1] = null;
                                rowa[ODI.ColHHPN[ODI.odiMatrixBOM] - 1] = Convert.ToString(a.Value.HHPN);
                                rowa[ODI.ColDescription[ODI.odiMatrixBOM] - 1] = Convert.ToString(a.Value.Description);
                                rowa[ODI.ColMFG[ODI.odiMatrixBOM] - 1] = Convert.ToString(a.Value.MFG);
                                rowa[ODI.ColMFGPN[ODI.odiMatrixBOM] - 1] = Convert.ToString(a.Value.MFGPN);
                                rowa[ODI.ColQty[ODI.odiMatrixBOM] - 1] = null;
                                rowa[ODI.ColLocation[ODI.odiMatrixBOM] - 1] = null;
                                rowa[ODI.ColRemark[ODI.odiMatrixBOM] - 1] = Convert.ToString(a.Value.Remark);
                                outlist.Add(rowa);
                            }
                        GroupRowCountList.Add(outlist.Count);
                    }

                    //建立outarr(實際寫入excel array), size: 表頭(5rows)+資料 
                    object[,] outarr = (object[,])Array.CreateInstance(typeof(object), new int[] { outlist.Count() + 5, outColCount }, new int[] { 1, 1 });

                    //寫入結果到outarr
                    for (int ar = 0; ar < outlist.Count(); ar++)
                        for (int ac = 0; ac < outColCount; ac++)
                            outarr[ar + 6, ac + 1] = outlist[ar][ac];


                    //寫入表頭到outlist
                    outarr[1, 1] = "FUJIN PRECISION INDUSTRY(SHENZHEN) CO.,LTD";
                    outarr[2, 1] = "BILL OF MATERIAL";
                    outarr[3, 2] = arr[3, 2];
                    outarr[3, 6] = arr[3, 6];
                    outarr[3, 8] = arr[3, 8];
                    outarr[4, 2] = arr[4, 2];
                    outarr[4, 6] = arr[4, 6];
                    outarr[4, 8] = arr[4, 8];

                    outarr[5, 1] = "Item";
                    outarr[5, 2] = "HF PN";
                    outarr[5, 3] = "STD PN";
                    outarr[5, 4] = "Description";
                    outarr[5, 5] = "Supplier";
                    outarr[5, 6] = "Supplier PN";
                    outarr[5, 7] = "Qty";
                    outarr[5, 8] = "Location";
                    outarr[4, 9] = "Total Qty";
                    outarr[5, 9] = "Total Q'ty";
                    outarr[4, 10] = "Total Set";
                    outarr[5, 10] = "Total Set";
                    outarr[4, 11] = "Model A";
                    outarr[4, 12] = "Model B";
                    outarr[4, 13] = "Model C";
                    outarr[4, 14] = "Model D";
                    outarr[4, 15] = "Model E";
                    outarr[4, 16] = "Model F";
                    outarr[5, 17] = "Remark";


                    this.Status1.Text = "Creating Matrix BOM - " + sh;

                    /*** 寫入Excel ***/
                    Excel.Worksheet Wst;
                    Excel.Range rng;
                    Excel.Range rng2;
                    int lastrow = outarr.GetUpperBound(0);
                    int lastcol = outarr.GetUpperBound(1);

                    Wst = wkbout.Worksheets.Add();
                    Wst.Name = sh;

                    rng = Wst.Range["A1"].Resize[lastrow, lastcol];
                    rng.Value2 = outarr;

                    /*** 設定格式 ***/
                    oXL.ActiveWindow.DisplayGridlines = false;

                    //基本字體
                    rng = Wst.Range["A1"].EntireColumn.EntireRow;
                    rng.Font.Name = "Calibri";
                    rng.Font.Size = 10;
                    rng.Font.Color = System.Drawing.Color.FromArgb(0, 0, 0);

                    //欄寬, 對齊
                    ((Excel.Range)Wst.Columns["A:A", System.Type.Missing]).ColumnWidth = 3.75;
                    ((Excel.Range)Wst.Columns["B:B", System.Type.Missing]).ColumnWidth = 14;
                    ((Excel.Range)Wst.Columns["C:C", System.Type.Missing]).ColumnWidth = 0;
                    ((Excel.Range)Wst.Columns["D:D", System.Type.Missing]).ColumnWidth = 35;
                    ((Excel.Range)Wst.Columns["E:E", System.Type.Missing]).ColumnWidth = 10;
                    ((Excel.Range)Wst.Columns["F:F", System.Type.Missing]).ColumnWidth = 18;
                    ((Excel.Range)Wst.Columns["G:G", System.Type.Missing]).ColumnWidth = 3.75;
                    ((Excel.Range)Wst.Columns["H:H", System.Type.Missing]).ColumnWidth = 20;
                    ((Excel.Range)Wst.Columns["I:I", System.Type.Missing]).ColumnWidth = 4;
                    ((Excel.Range)Wst.Columns["J:J", System.Type.Missing]).ColumnWidth = 4;
                    ((Excel.Range)Wst.Columns["K:K", System.Type.Missing]).ColumnWidth = 3.43;
                    ((Excel.Range)Wst.Columns["L:L", System.Type.Missing]).ColumnWidth = 3.43;
                    ((Excel.Range)Wst.Columns["M:M", System.Type.Missing]).ColumnWidth = 3.43;
                    ((Excel.Range)Wst.Columns["N:N", System.Type.Missing]).ColumnWidth = 3.43;
                    ((Excel.Range)Wst.Columns["O:O", System.Type.Missing]).ColumnWidth = 3.43;
                    ((Excel.Range)Wst.Columns["P:P", System.Type.Missing]).ColumnWidth = 3.43;
                    ((Excel.Range)Wst.Columns["Q:Q", System.Type.Missing]).ColumnWidth = 50;

                    rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //基本水平置中

                    ((Excel.Range)Wst.Columns["B:F", System.Type.Missing]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ((Excel.Range)Wst.Columns["H:H", System.Type.Missing]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ((Excel.Range)Wst.Columns["Q:Q", System.Type.Missing]).HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                    ((Excel.Range)Wst.Columns["Q:Q", System.Type.Missing]).WrapText = false;

                    //標題
                    rng = Wst.get_Range("A1", "Q1");
                    rng.Merge(false);
                    rng.Font.Color = System.Drawing.Color.FromArgb(245, 245, 245);
                    rng.Interior.Color = System.Drawing.Color.FromArgb(50, 50, 50);
                    rng.Font.Name = "Arial";
                    rng.Font.FontStyle = "Bold";
                    rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rng.Font.Size = 12;

                    rng = Wst.get_Range("A2", "Q2");
                    rng.Merge(false);
                    rng.Font.Color = System.Drawing.Color.FromArgb(245, 245, 245);
                    rng.Interior.Color = System.Drawing.Color.FromArgb(50, 50, 50);
                    rng.Font.Name = "Arial";
                    rng.Font.FontStyle = "Bold";
                    rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rng.Font.Size = 12;


                    /*
                    rng = Wst.Range["G1"].Resize[1, lastcol - 6];
                    rng.Font.FontStyle = "Bold";
                    rng = Wst.Range["G2"].Resize[2, lastcol - 6];
                    rng.Font.Color = System.Drawing.Color.FromArgb(115, 117, 118);

                    rng = Wst.Range["F1"].Resize[3, 1];
                    rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    rng.Font.Color = System.Drawing.Color.FromArgb(150, 150, 150);
                    */

                    //Product code highlight
                    Wst.Range["B3"].Characters[13].Font.FontStyle = "Bold";
                    Wst.Range["F3"].Characters[13].Font.FontStyle = "Bold";
                    Wst.Range["H3"].Characters[13].Font.FontStyle = "Bold";

                    //表頭
                    rng = Wst.Range["A5"].Resize[1, ODI.ColTotalSet[ODI.odiMatrixBOM]];
                    rng.Interior.Color = System.Drawing.Color.FromArgb(52, 58, 64);
                    rng.Font.Color = System.Drawing.Color.FromArgb(245, 248, 251);
                    rng.Font.FontStyle = "Bold";
                    //rng.Font.Name = "Arial";
                    rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    //rng.Font.Size = 10;

                    //Model 欄位特殊處理
                    rng = Wst.Range["K4"].Resize[1, ODI.NumOfMatrix];
                    rng.Interior.Color = System.Drawing.Color.FromArgb(52, 58, 64);
                    rng.Font.Color = System.Drawing.Color.FromArgb(245, 248, 251);
                    rng.Font.FontStyle = "Bold";
                    rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight;
                    rng.InsertIndent(1);
                    rng.Characters[1,6].Font.Color = System.Drawing.Color.FromArgb(52, 58, 64);

                    //Remark欄位
                    rng = Wst.Range["Q5"];
                    rng.Interior.Color = System.Drawing.Color.FromArgb(52, 58, 64);
                    rng.Font.Color = System.Drawing.Color.FromArgb(245, 248, 251);
                    rng.Font.FontStyle = "Bold";
                    rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                    //Total Qty
                    rng2 = Wst.get_Range("A3", "A4");
                    rng2.Merge(false);
                    rng2.Copy();
                    rng = Wst.get_Range("I4", "I5");
                    rng.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone);
                    rng.Interior.Color = System.Drawing.Color.FromArgb(52, 58, 64);
                    rng.Font.Color = System.Drawing.Color.FromArgb(245, 248, 251);
                    rng.Font.FontStyle = "Bold";
                    rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rng.WrapText = true;
                    

                    //Total Set
                    rng = Wst.get_Range("J4", "J5");
                    rng2.Copy();
                    rng.PasteSpecial(Excel.XlPasteType.xlPasteFormats, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone);
                    rng.Interior.Color = System.Drawing.Color.FromArgb(52, 58, 64);
                    rng.Font.Color = System.Drawing.Color.FromArgb(245, 248, 251);
                    rng.Font.FontStyle = "Bold";
                    rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    rng.WrapText = true;
                    

                    //基本框線
                    rng = Wst.Range["A5"].Resize[lastrow - 4, ODI.ColRemark[ODI.odiMatrixBOM]];
                    rng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    rng.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    rng.Borders.Color = System.Drawing.Color.FromArgb(101, 101, 101);

                    rng = Wst.Range["K4"].Resize[1, ODI.NumOfMatrix];
                    rng.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    rng.Borders.Weight = Excel.XlBorderWeight.xlThin;
                    rng.Borders.Color = System.Drawing.Color.FromArgb(101, 101, 101);

                    //每個group加框上色
                    //加上公式
                    bool oddgroup = false;
                    for (int i = 0; i < GroupRowCountList.Count - 1; i++)
                    {
                        this.Status1.Text = "Creating Matrix BOM - " + sh + " Group " + Convert.ToString(i);
                        //Main source 粗體
                        rng = (Excel.Range)Wst.Cells[GroupRowCountList[i] + 6, 1].Resize[1, lastcol];
                        rng.Font.Bold = true;

                        if (oddgroup)
                        {
                            //整個group上色
                            rng = (Excel.Range)Wst.Cells[GroupRowCountList[i] + 6, 1].Resize[GroupRowCountList[i + 1] - GroupRowCountList[i], lastcol];
                            rng.Interior.Color = System.Drawing.Color.FromArgb(235, 235, 235);
                        }
                        oddgroup = !oddgroup;

                        //合併ID欄位
                        //v1.0: 為了符合原MatrixBOM格式，改為不合併，僅隱藏格線
                        rng = (Excel.Range)Wst.Cells[GroupRowCountList[i] + 6, 1].Resize[GroupRowCountList[i + 1] - GroupRowCountList[i], 1];
                        //rng.Merge(false);
                        rng.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel.XlLineStyle.xlLineStyleNone;
                        rng.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlLineStyleNone;


                        //公式
                        rng = (Excel.Range)Wst.Cells[GroupRowCountList[i] + 6, 9].Resize[GroupRowCountList[i + 1] - GroupRowCountList[i], 1];
                        rng.FormulaR1C1 = "=R" + Convert.ToString(GroupRowCountList[i] + 6) + "C[-2] * R[0]C[1]";

                        rng = (Excel.Range)Wst.Cells[GroupRowCountList[i] + 6, 10].Resize[GroupRowCountList[i + 1] - GroupRowCountList[i], 1];
                        rng.FormulaR1C1 = @"=IF(OR(RC[1]=""V"",RC[1]=""v""),R5C[1],0)+IF(OR(RC[2]=""V"",RC[2]=""v""),R5C[2],0)+IF(OR(RC[3]=""V"",RC[3]=""v""),R5C[3],0)+IF(OR(RC[4]=""V"",RC[4]=""v""),R5C[4],0)+IF(OR(RC[5]=""V"",RC[5]=""v""),R5C[5],0)+IF(OR(RC[6]=""V"",RC[6]=""v""),R5C[6],0)";

                    }

                    /*** FileList ***/
                    /*
                    rng = (Excel.Range)Wst.Cells[lastrow + 2, 1];
                    rng.Value2 = "Result generated from the following files:";
                    rng.Font.Bold = true;
                    rng.HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                    fileidx = 1;
                    foreach (string f in listBox_BOMList.Items)
                    {
                        rng = rng.Offset[1, 1]; //Offset to Next row, next column, for file name
                        rng.Value2 = f;
                        rng = rng.Offset[0, -1];// offset 1 column left, for index
                        rng.Value2 = Convert.ToString(fileidx++) + ".";
                    }
                    rng = (Excel.Range)Wst.Cells[lastrow + 3, 1].Resize[listBox_BOMList.Items.Count, 2];
                    rng.Font.Color = System.Drawing.Color.FromArgb(150, 150, 150);
                    */

                    /*** Footer ***/


                }

                wkbout.Worksheets["WAIT_ForDELEte"].Delete();

                wkb.Close(false);

            }
            catch (Exception ex)
            {
                MessageBox.Show("Something wrong with EXCEL operation, please close all EXCEL process in memory and try again. \n If still has issue, please check Matriz Page (Help-->About) for solution or support. " + ex.Message);

                /*
                wkb.Close(false);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(wkb);
                wkb = null;
                */
                oXL.Workbooks.Close();
                oXL.Quit();
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXL);
                oXL = null;
                DialogResult = DialogResult.OK;
                btn_Go.Enabled = true;
                return;
            }

            wkb = null;

            this.Status1.Text = "Matrix BOM Done. ^o^";
            oXL.DisplayAlerts = true;
            oXL.Visible = true;
            oXL.UserControl = true;

            DialogResult = DialogResult.OK;
            btn_Go.Enabled = true;

            return;

        }

        private void aboutMatrizToolStripMenuItem_Click(object sender, EventArgs e)
        {
            AboutForm dlgAbout = new AboutForm();
            dlgAbout.ShowDialog();
        }

        private void fAQToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://angeloeyez.github.io/Matriz-MatrixBOMTool/#faq");
        }
    }
}

