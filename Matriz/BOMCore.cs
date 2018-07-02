using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;


namespace BOMCore
{
    /*定義常數*/
    static class ODI
    {
        /* BOM 格式行定義 */
        public const int RowBOMStart = 6;

        /* BOM Type */
        public const int odiMatrixBOM = 0;
        public const int odiMfgBOM = 1;
        public const int odiCostBOM = 2;

        /* Matrix BOM 專用 */
        public const string ColTotalQtyCell = "I6";
        public const int NumOfMatrix = 6;

        /* BOM 格式欄位定義 (99 = not used) {odiMatrixBOM, odiMfgBOM, odiCostBOM} (CostBOM未確認)*/ 
        public static int[] ColNum = { 1, 1, 1 };
        public static int[] ColHHPN = { 2, 2, 1 };
        public static int[] ColSTDPN = { 3, 3, 3 };
        public static int[] ColDescription = { 4, 5, 1 };
        public static int[] ColMFG = { 5, 6, 1 };
        public static int[] ColMFGPN = { 6, 7, 1 };
        public static int[] ColQty = { 7, 8, 1 };
        public static int[] ColLocation = { 8, 9, 1 };
        public static int[] ColTotalQty = { 9, 99, 99 };
        public static int[] ColTotalSet = { 10, 99, 99 };
        public static int[] ColCCL = { 99, 10, 1 };
        public static int[] ColMatrixA = { 11, 99, 1 };
        public static int[] ColRemark = { NumOfMatrix + 11, 12, 1 }; //同時是最後一欄


        public static List<string> get_SheetList(int t)
        {
            switch (t)
            {
                case odiMatrixBOM:
                    return new List<string> { "SMD", "PTH", "BOTTOM" };
                case odiMfgBOM:
                    return new List<string> { "ALL", "SMD", "PTH", "BOTTOM", "MP" }; //"ALL" 要排在最前面，有用到
                case odiCostBOM:
                    return new List<string> { "SMD", "PTH", "BOTTOM" };
                default:
                    return new List<string>();
            }
        }

    }

    class Item
    {
        public int Num;
        public string HHPN;
        public string Description;
        public string MFG;
        public string MFGPN;
        public List<int> Qty = new List<int>();
        public string Location;
        public string CCL;
        public string Remark;
        public string[] Matrix = new string[ODI.NumOfMatrix];

        public List<string> ProjectName = new List<string>();
        public List<string> SCHVersion = new List<string>();
        public List<string> BOMVersion = new List<string>();
        public List<string> BOMKey = new List<string>();

        protected bool _match = true;
        public virtual bool match //Item match ()
        {
            set { _match = value; }
            get
            {
                if (Qty.Count() > 1)
                    _match = !Qty.Contains(0);
                else
                    _match = true;
                return _match;
            }
        }

        public Item(int irow, object[,] arr, int type) { RetrieveData(irow, arr, type); }
        public Item() { }

        public void RetrieveData(int irow, object[,] arr, int t)
        {

            switch (t)
            {
                case ODI.odiMatrixBOM:
                    Num = Convert.ToInt32(arr[irow, ODI.ColNum[t]]);
                    HHPN = Convert.ToString(arr[irow, ODI.ColHHPN[t]]);
                    Description = Convert.ToString(arr[irow, ODI.ColDescription[t]]).Replace("\n", "");
                    MFG = Convert.ToString(arr[irow, ODI.ColMFG[t]]);
                    MFGPN = Convert.ToString(arr[irow, ODI.ColMFGPN[t]]);
                    Qty.Add(Convert.ToInt32(arr[irow, ODI.ColQty[t]]));
                    Location = Convert.ToString(arr[irow, ODI.ColLocation[t]]);
                    Remark = Convert.ToString(arr[irow, ODI.ColRemark[t]]);

                    for (int m = 0; m < ODI.NumOfMatrix; m++)
                        Matrix[m] = Convert.ToString(arr[irow, ODI.ColMatrixA[t] + m]);

                    ProjectName.Add(getVersion(arr[3, 2]));
                    SCHVersion.Add(getVersion(arr[3, 4]));
                    BOMVersion.Add(getVersion(arr[3, 8]));
                    BOMKey.Add(getBomKey());

                    break;
                case ODI.odiMfgBOM:
                    Num = Convert.ToInt32(arr[irow, ODI.ColNum[t]]);
                    HHPN = Convert.ToString(arr[irow, ODI.ColHHPN[t]]);
                    Description = Convert.ToString(arr[irow, ODI.ColDescription[t]]).Replace("\n", "");
                    MFG = Convert.ToString(arr[irow, ODI.ColMFG[t]]);
                    MFGPN = Convert.ToString(arr[irow, ODI.ColMFGPN[t]]);
                    Qty.Add(Convert.ToInt32(arr[irow, ODI.ColQty[t]]));
                    Location = Convert.ToString(arr[irow, ODI.ColLocation[t]]);
                    CCL = Convert.ToString(arr[irow, ODI.ColCCL[t]]);
                    Remark = Convert.ToString(arr[irow, ODI.ColRemark[t]]);

                    ProjectName.Add(getVersion(arr[3, 2]));
                    SCHVersion.Add(getVersion(arr[3, 4]));
                    BOMVersion.Add(getVersion(arr[3, 8]));
                    BOMKey.Add(getBomKey());

                    break;
                case ODI.odiCostBOM:
                    break;
                default:
                    break;
            }

        }

        private string getVersion(object s)
        {
            string[] sp = { ": " };
            string[] result = Convert.ToString(s).Split(sp, StringSplitOptions.RemoveEmptyEntries);

            if (result.Length > 1)
                return result[1];
            else
                return "N/A";
        }

        public void WriteData(int irow, object[,] arr, int t)
        {
            arr[irow, ODI.ColNum[t]] = (Num > 0) ? (object)Num : null;
            arr[irow, ODI.ColHHPN[t]] = HHPN;
            arr[irow, ODI.ColDescription[t]] = Description;
            arr[irow, ODI.ColMFG[t]] = MFG;
            arr[irow, ODI.ColMFGPN[t]] = MFGPN;
            arr[irow, ODI.ColQty[t]] = (Qty[0] > 0) ? (object)Qty[0] : null;
            arr[irow, ODI.ColLocation[t]] = Location;
            arr[irow, ODI.ColRemark[t]] = Remark;

            for (int m = 0; m < ODI.NumOfMatrix; m++)
                arr[irow, ODI.ColMatrixA[t] + m] = Matrix[m];
        }
        public string getKey()
        {
            if (HHPN.Length > 0 && MFGPN.Length > 0)
                return HHPN + MFGPN;
            else
            { //[Err] 萬一HHPN和MFGPN有一個沒有值:BOM item有問題
                //Globals.ThisAddIn.odiMsgOut("HHPN / MFGPN error, check code and debug!");
                MessageBox.Show("HHPN / MFGPN error, check code and debug!");
                return "";
            }
        }

        public string getBomKey(int keyIndex = 0)
        //取得bomkey
        //默認取得第0組
        // 如果有輸入index, 取得對應的bomkey (for 多份BOM)
        {
            if (ProjectName.Count > 0 && SCHVersion.Count > 0 && BOMVersion.Count > 0)
            {
                if (ProjectName[keyIndex].Length > 0 && SCHVersion[keyIndex].Length > 0 && BOMVersion[keyIndex].Length > 0)
                    return ProjectName[keyIndex] + ":" + SCHVersion[keyIndex] + ":" + BOMVersion[keyIndex]; //bomkey;
                else
                { //[Err] 萬一ProjectName, SCHVersion 和BOMVersion有一個沒有值:BOM item有問題
                    //Globals.ThisAddIn.odiMsgOut("[Error:getBomKey()] HHPN / MFGPN error, check code and debug!");
                    MessageBox.Show("[Error:getBomKey()] HHPN / MFGPN error, check code and debug!");
                    return null;
                }
            }
            else
                return null;
        }

        public void syncMatrix(Item i)
        {
            Matrix = i.Matrix;
        }
    }

    class Group : Item
    {
        public Dictionary<string, Item> AVL = new Dictionary<string, Item>(); // AVLs

        public Group() { }
        public override bool match //Group match ()
        {
            set { _match = value; }
            get
            {
                foreach (var i in AVL)
                    _match &= i.Value.match;
                return _match;
            }
        }
        public void setAVL()
        {
            foreach (var a in AVL)
            {
                a.Value.Qty[0] = this.Qty[0];
                a.Value.ProjectName[0] = this.ProjectName[0];
                a.Value.SCHVersion[0] = this.SCHVersion[0];
                a.Value.BOMVersion[0] = this.BOMVersion[0];
            }
        }


        public void combineGroup(Group g)
        /*將單一group合併到本group*/
        {
            //確認g是單一Group
            if (g.Qty.Count > 1)
            {
                //Globals.ThisAddIn.odiMsgOut("[Error] Can't comine complex group, check code!");
                MessageBox.Show("[Error] Can't comine complex group, check code!");
                return;
            }
            else
            {
                //合併main source資料
                Qty.Add(g.Qty[0]);
                ProjectName.Add(g.ProjectName[0]);
                SCHVersion.Add(g.SCHVersion[0]);
                BOMVersion.Add(g.BOMVersion[0]);
                BOMKey.Add(g.getBomKey());

                //歷遍目前AVL，找到相同的合併，並且從g移除
                bool foundAVLing;
                foreach (var i in AVL)
                {
                    foundAVLing = g.AVL.TryGetValue(i.Key, out Item gi);

                    i.Value.ProjectName.Add(g.ProjectName[0]); //不論有沒有找到，都要增加project欄位
                    i.Value.SCHVersion.Add(g.SCHVersion[0]);
                    i.Value.BOMVersion.Add(g.BOMVersion[0]);
                    i.Value.BOMKey.Add(g.getBomKey());
                    if (foundAVLing)
                    {
                        i.Value.Qty.Add(gi.Qty[0]);
                        g.AVL.Remove(i.Key);
                    }
                    else
                        i.Value.Qty.Add(0);
                }

                //如果g還有剩下(新的)AVL item，則加入目前的AVL dictionary, 並且補足其前面的空白欄位
                if (g.AVL.Count > 0)
                {
                    foreach (var j in g.AVL)
                    {
                        j.Value.ProjectName = ProjectName;
                        j.Value.SCHVersion = SCHVersion;
                        j.Value.BOMVersion = BOMVersion;
                        j.Value.BOMKey = BOMKey;
                        for (int k = 0; k < Qty.Count() - 1; k++)
                            j.Value.Qty.Insert(0, 0);
                        AVL.Add(j.Key, j.Value);
                    }
                }
            }
        }



    }

}
