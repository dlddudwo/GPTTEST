using MetroFramework.Forms;
using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static AMI_Manager.Forms.RuleSimulator;
using System.Windows.Controls;
using MS.WindowsAPICodePack.Internal;
using Newtonsoft.Json.Serialization;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.Collections;
using System.Runtime.InteropServices;

using Application = Microsoft.Office.Interop.Excel.Application;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Input;
using System.Reflection;
using Microsoft.Office.Core;
using System.Windows.Media.Media3D;
using System.Windows.Media.Imaging;
using Microsoft.WindowsAPICodePack.Dialogs;
using static AMI_Manager.Forms.Main.LadyBugForm;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using Rectangle = System.Drawing.Rectangle;
using System.Xml.Schema;
using System.Windows.Media.Effects;
using CsvHelper;
using CsvHelper.Configuration;
using System.Globalization;
using System.Windows.Shapes;
using Path = System.IO.Path;
using System.Windows.Documents;
using System.Drawing.Imaging;
using Image = System.Drawing.Image;
using Font = System.Drawing.Font;
using Control = System.Windows.Forms.Control;
using ListViewItem = System.Windows.Forms.ListViewItem;

namespace AMI_Manager.Forms.Main
{

    public partial class LadyBugForm : MetroForm
    {

        //CONFIG사용할 함수 선언 
        // Windows API 함수 선언
        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        private static extern int GetPrivateProfileString(
            string lpAppName,      // 섹션 이름 (예: "Database")
            string lpKeyName,      // 키 이름 (예: "Server")
            string lpDefault,      // 기본값 (키가 없을 때 반환)
            StringBuilder lpReturnedString, // 결과를 저장할 버퍼
            int nSize,            // 버퍼 크기
            string lpFileName);    // INI 파일 경로

        // INI 파일에서 값 읽기
        private static readonly Dictionary<string, string> IniValueCache = new Dictionary<string, string>();
        public static string ReadValue(string section, string key, string defaultValue, string filePath)
        {
            string cacheKey = filePath + "|" + section + "|" + key + "|" + defaultValue;
            string cachedValue;
            if (IniValueCache.TryGetValue(cacheKey, out cachedValue))
            {
                return cachedValue;
            }

            StringBuilder value = new StringBuilder(255); // 충분한 크기의 버퍼
            GetPrivateProfileString(section, key, defaultValue, value, value.Capacity, filePath);
            string result = value.ToString();
            IniValueCache[cacheKey] = result;
            return result;
        }

        private static void ClearIniValueCache()
        {
            IniValueCache.Clear();
        }


        public struct VP_INFO
        {
            public int VP_NUM;
            public string Path_Vp_Result;
            public string Path_Vp_Recipe;
            public string Path_Vp_Log;
            public string Path_Vp_Inspection;
        }
        public struct Navi
        {
            //기본정보
            public int Panel_index;
            public int Defect_index;
            public string judge;
            public string classify;
            public int type;

            //세부정보
            public string panel_id;
            public int vp_num;

        }





        private readonly ConfigState _configState = new ConfigState();
        private readonly SearchState _searchState = new SearchState();
        private readonly DefectState _defectState = new DefectState();

        public List<VP_INFO> vp_info_list { get { return _configState.VpInfoList; } }
        public string Model_name { get { return _configState.ModelName; } set { _configState.ModelName = value; } }


        //변수정리 Config 관련 -> 나중에 구조체로좀 묶자
        public int View_mode { get { return _configState.ViewMode; } set { _configState.ViewMode = value; } }
        public List<string> VIEW_DEFECT_FEATURE_NAME { get { return _configState.ViewDefectFeatureNames; } }
        public string POSITION_FEATURE_NAME_X { get { return _configState.PositionFeatureNameX; } set { _configState.PositionFeatureNameX = value; } }
        public string POSITION_FEATURE_NAME_Y { get { return _configState.PositionFeatureNameY; } set { _configState.PositionFeatureNameY = value; } }
        public int Swap_X { get { return _configState.SwapX; } set { _configState.SwapX = value; } }
        public int Swap_Y { get { return _configState.SwapY; } set { _configState.SwapY = value; } }

        public int View_feature_num { get { return _configState.ViewFeatureNum; } set { _configState.ViewFeatureNum = value; } }
        private int defectMapRowCount { get { return _configState.DefectMapRowCount; } set { _configState.DefectMapRowCount = value; } }
        private int defectMapColCount { get { return _configState.DefectMapColCount; } set { _configState.DefectMapColCount = value; } }
        private int defectMapRowColCount { get { return _configState.DefectMapRowColCount; } set { _configState.DefectMapRowColCount = value; } }
        private int defectMapRowSwap { get { return _configState.DefectMapRowSwap; } set { _configState.DefectMapRowSwap = value; } }
        private int defectMapColSwap { get { return _configState.DefectMapColSwap; } set { _configState.DefectMapColSwap = value; } }
        private string[] ignoreIndexArray { get { return _configState.IgnoreIndexArray; } set { _configState.IgnoreIndexArray = value; } }

        //변수정리  Product Search 
        public Insp_info insp_info;
        public string recipe_path { get { return _searchState.RecipePath; } set { _searchState.RecipePath = value; } }
        public string Disk_base { get { return _configState.DiskBase; } set { _configState.DiskBase = value; } }
        public string inspection_path_base { get { return _configState.InspectionPathBase; } set { _configState.InspectionPathBase = value; } }
        public string recipe_path_base { get { return _configState.RecipePathBase; } set { _configState.RecipePathBase = value; } }
        public string result_path_base { get { return _configState.ResultPathBase; } set { _configState.ResultPathBase = value; } }
        public string log_path_base { get { return _configState.LogPathBase; } set { _configState.LogPathBase = value; } }

        public string Simulator_Config_path { get { return _configState.SimulatorConfigPath; } set { _configState.SimulatorConfigPath = value; } }

        private DateTime StartDate { get { return _searchState.StartDate; } set { _searchState.StartDate = value; } }
        private DateTime EndDate { get { return _searchState.EndDate; } set { _searchState.EndDate = value; } }
        private List<string> matchingFiles { get { return _searchState.MatchingFiles; } }
        private List<InspectionData> inspectionDataList { get { return _searchState.InspectionDataList; } }
        private bool Panel_ID_Search { get { return _searchState.PanelIdSearch; } set { _searchState.PanelIdSearch = value; } }

        //변수정리 Defect Search
        public List<Dictionary<string, object>> Feature_row { get { return _defectState.FeatureRow; } }
        public List<Dictionary<string, object>> Feature_row_post { get { return _defectState.FeatureRowPost; } }
        public List<string> Crop_bin_path_Pre { get { return _defectState.CropBinPathPre; } }
        public List<string> Crop_bin_path_Post { get { return _defectState.CropBinPathPost; } }
        public List<(string FolderName, string FolderPath)> lstMatchingFolders { get { return _searchState.MatchingFolders; } }
        private List<Classify_info> classify_Infos { get { return _defectState.ClassifyInfos; } }
        private List<Classify_Pre_info> classify_Pre_Infos { get { return _defectState.ClassifyPreInfos; } }
        private List<Classify_Post_info> classify_Post_Infos { get { return _defectState.ClassifyPostInfos; } }
        public List<string> used_Feature_Pre { get { return _defectState.UsedFeaturePre; } }
        public List<string> used_Feature_Post { get { return _defectState.UsedFeaturePost; } }
        public List<System.Drawing.Point> Defect_Position { get { return _defectState.DefectPosition; } }
        public List<bool> Defect_Position_Judge { get { return _defectState.DefectPositionJudge; } }
        private List<int> VP_Defect_num { get { return _defectState.VpDefectNum; } }
        public Navi navi = new Navi();





        public double picturebox_ratio_x, picturebox_ratio_y, select_nymber_defect = 9999;
        public int Judge_mode = 0;

        private bool check_picturebox = false;
        private Bitmap originalImage;
        private int divisionCount_row = 2; // 이미지를 나눌 섹션 수
        private int divisionCount_col = 8; // 이미지를 나눌 섹션 수
        private Rectangle highlightRect; // 강조 표시할 영역

        private int Defect_map_select_index = -1;


        //변수정리 Classifier 
        public Bitmap Bitmap_Crop = new Bitmap(512, 128);
        CSimulationRun cSimulationrun = new CSimulationRun();


        //Listview Sort 관련 
        private int lastSortedColumn = -1;
        private SortOrder lastSortOrder = SortOrder.None;


        private int DGV_MAIN_INDEX = 0;
        private int CheckClassifier = 0;
        private bool isLoading;
        private bool isDefectSelectionUpdating;
        private int lastDefectListSelectedIndex = -1;
        private float pbMainZoom = 1f;
        private PointF pbMainPanOffset = new PointF(0f, 0f);
        private bool pbMainDragging;
        private System.Drawing.Point pbMainLastMousePoint;
        private const float PbMainMinZoom = 0.2f;
        private const float PbMainMaxZoom = 40f;
        private const float PbMainZoomStep = 1.2f;
        private const float PbMainPixelLabelZoomThreshold = 12f;








        //Product Search 
        public string Excel_wirte_path = null;

        public LadyBugForm()
        {
            InitializeComponent();
            this.Theme = MetroFramework.MetroThemeStyle.Dark;
        }
        public LadyBugForm(ManagerForm _managerForm)
        {
            InitializeComponent();
            this.Theme = MetroFramework.MetroThemeStyle.Dark;
        }

        private void LadyBug_Load(object sender, EventArgs e)
        {
            if (isLoading)
            {
                return;
            }

            isLoading = true;
            try
            {
                ClearIniValueCache();
                InitializeConfig();
                InitializeUiDefaults();
                InitializeDefectListUi();
            }
            finally
            {
                isLoading = false;
            }
        }

        private void InitializeConfig()
        {
            int vp_num = Convert.ToInt32(ReadValue("BASE", "VP_NUM", "0", Simulator_Config_path));
            if (vp_num == 0)
                MessageBox.Show("Check Simulation_Config.ini file ");

            for (int i = 0; i < vp_num; i++)
            {
                CBB_VP_NUM.Items.Add("VP" + (i + 1));
            }

            for (int i = 0; i < vp_num; i++)
            {
                string section = "VP" + (i + 1);
                VP_INFO vp_info = new VP_INFO();
                vp_info.Path_Vp_Result = ReadValue(section, "PATH_VP_RESULT", "", Simulator_Config_path);
                vp_info.Path_Vp_Recipe = ReadValue(section, "PATH_VP_RECIPE", "0", Simulator_Config_path);
                vp_info.Path_Vp_Log = ReadValue(section, "PATH_VP_LOG", "0", Simulator_Config_path);
                vp_info.Path_Vp_Inspection = ReadValue(section, "PATH_VP_INSPECTION", "0", Simulator_Config_path);
                vp_info_list.Add(vp_info);

                Crop_bin_path_Pre.Add("");
                Crop_bin_path_Post.Add("");
                VP_Defect_num.Add(0);
            }

            CBB_DEFECT_NAME.Items.Add("ALL");
            CBB_DEFECT_NAME_EX.Items.Add("ALL");

            for (int i = 0; i < 999; i++)
            {
                string classifyName = ReadValue("CIASSIFY_NAME", "NAME" + (i + 1), "", Simulator_Config_path);
                if (classifyName != "")
                {
                    CBB_DEFECT_NAME.Items.Add(classifyName);
                    CBB_DEFECT_NAME_EX.Items.Add(classifyName);
                }
            }

            CBB_JUDGE.SelectedIndex = 0;
            CBB_DEFECT_NAME.SelectedIndex = 0;
            CBB_DEFECT_NAME_EX.SelectedIndex = 0;

            View_mode = Convert.ToInt32(ReadValue("BASE", "VEIW_MODE", "1", Simulator_Config_path));
            VIEW_DEFECT_FEATURE_NAME.Clear();
            for (int i = 0; i < 999; i++)
            {
                string featureName = ReadValue("DEFECT_FEATURE", "NAME" + (i + 1), "", Simulator_Config_path);
                if (featureName != "")
                {
                    VIEW_DEFECT_FEATURE_NAME.Add(featureName);
                }
            }

            View_feature_num = VIEW_DEFECT_FEATURE_NAME.Count;
            POSITION_FEATURE_NAME_X = ReadValue("POSITION_FEATURE_NAME", "X", "", Simulator_Config_path);
            POSITION_FEATURE_NAME_Y = ReadValue("POSITION_FEATURE_NAME", "Y", "", Simulator_Config_path);
            Swap_X = Convert.ToInt32(ReadValue("POSITION_FEATURE_NAME", "SWAPX", "", Simulator_Config_path));
            Swap_Y = Convert.ToInt32(ReadValue("POSITION_FEATURE_NAME", "SWAPY", "", Simulator_Config_path));

            int parsedRowCount;
            int parsedColCount;
            int parsedRowColCount;
            int parsedRowSwap;
            int parsedColSwap;
            if (!int.TryParse(ReadValue("DEFECT_MAP", "ROW_COUNT", "10", Simulator_Config_path), out parsedRowCount) || parsedRowCount <= 0)
            {
                parsedRowCount = 10;
            }
            if (!int.TryParse(ReadValue("DEFECT_MAP", "COL_COUNT", "10", Simulator_Config_path), out parsedColCount) || parsedColCount <= 0)
            {
                parsedColCount = 10;
            }
            if (!int.TryParse(ReadValue("DEFECT_MAP", "ROWCOL_SWAP", "0", Simulator_Config_path), out parsedRowColCount) || (parsedRowColCount != 0 && parsedRowColCount != 1))
            {
                parsedRowColCount = 0;
            }
            if (!int.TryParse(ReadValue("DEFECT_MAP", "ROW_SWAP", "1", Simulator_Config_path), out parsedRowSwap) || (parsedRowSwap != 0 && parsedRowSwap != 1))
            {
                parsedRowSwap = 0;
            }
            if (!int.TryParse(ReadValue("DEFECT_MAP", "COL_SWAP", "1", Simulator_Config_path), out parsedColSwap) || (parsedColSwap != 0 && parsedColSwap != 1))
            {
                parsedColSwap = 0;
            }

            string rawValue = ReadValue("DEFECT_MAP", "IGNORE_INDEX", "", Simulator_Config_path);

            ignoreIndexArray = rawValue
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim().ToUpper())

                .ToArray();

            defectMapRowCount = parsedRowCount;
            defectMapColCount = parsedColCount;
            defectMapRowColCount = parsedRowColCount;
            defectMapRowSwap = parsedRowSwap;
            defectMapColSwap = parsedColSwap;

            if (vp_info_list.Count > 0)
            {
                string[] recipe = GetSearchFolder_only(vp_info_list[0].Path_Vp_Recipe);
                for (int i = 0; i < recipe.Length; i++)
                    CBB_RECIPE.Items.Add(Path.GetFileName(recipe[i]));
            }

            CBB_RECIPE.DropDownStyle = ComboBoxStyle.DropDownList;
        }

        private void InitializeUiDefaults()
        {
            DGV_MAIN.ColumnCount = 3;
            DGV_MAIN.Columns[0].Name = "INDEX";
            DGV_MAIN.Columns[1].Name = "KEY";
            DGV_MAIN.Columns[2].Name = "VALUE";
            DGV_MAIN.Columns[0].Width = 50;
            DGV_MAIN.Columns[1].Width = 150;
            DGV_MAIN.Columns[2].Width = 150;

            RB_RECIPE_BACK.Checked = true;
            CB_RECIPE_WRITE.Checked = false;
            TB_RECIPE.Text = "5";

            PB_DEFECT_ARRAY.SizeMode = PictureBoxSizeMode.StretchImage;
            PB_DEFECTMAP.SizeMode = PictureBoxSizeMode.StretchImage;
            PB_MAIN.SizeMode = PictureBoxSizeMode.Normal;
            PB_MAIN.TabStop = true;

            PB_DEFECT_ARRAY.MouseMove += PictureBox_MouseMove;
            PB_DEFECT_ARRAY.MouseClick += PictureBox_MouseClick;
            PB_DEFECT_ARRAY.Paint += PictureBox_Paint;




            dtpDateFrom.Value = DateTime.Now;
            StartDate = dtpDateFrom.Value;
            StartDate = new DateTime(StartDate.Year, StartDate.Month, StartDate.Day, 0, 0, 0);
            dtpDateFrom.Value = StartDate;

            dtpDateTo.Value = DateTime.Now;
            EndDate = dtpDateTo.Value;
            EndDate = new DateTime(EndDate.Year, EndDate.Month, EndDate.Day, 23, 59, 59);
            dtpDateTo.Value = EndDate;

            CBB_DEFECT_JUDGE.SelectedIndex = 0;
            RTB_ORIGIN_POST.Hide();
            RTB_REPLACE_POST.Hide();
            RTB_ORIGIN_PRE.Show();
            RTB_REPLACE_PRE.Show();

            RegisterComboBoxDropDownAutoWidth(this);
        }

        private void InitializeDefectListUi()
        {
            LV_DEFECT_LIST.View = View.Details;
            LV_DEFECT_LIST.Columns.Add("Index");
            for (int i = 0; i < View_feature_num; i++)
            {
                LV_DEFECT_LIST.Columns.Add(VIEW_DEFECT_FEATURE_NAME[i]);
                CBB_NAVI_CLASSIFY.Items.Add(VIEW_DEFECT_FEATURE_NAME[i]);
            }
            LV_DEFECT_LIST.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize);
            CBB_NAVI_JUDGE.Items.Add("ALL");
            CBB_NAVI_JUDGE.Items.Add("OK");
            CBB_NAVI_JUDGE.Items.Add("NG");
        }

        private void SetPictureBoxImage(PictureBox pictureBox, Bitmap nextImage)
        {
            if (pictureBox == null)
            {
                return;
            }

            Image previous = pictureBox.Image;
            pictureBox.Image = nextImage;
            if (previous != null && !ReferenceEquals(previous, nextImage))
            {
                previous.Dispose();
            }

            if (ReferenceEquals(pictureBox, PB_MAIN))
            {
                ResetPbMainView();
                PB_MAIN.Invalidate();
            }
        }

        private void ResetPbMainView()
        {
            pbMainZoom = 1f;
            pbMainPanOffset = new PointF(0f, 0f);
            pbMainDragging = false;
        }

        private void DrawDefectMapGrid(Graphics g, int width, int height)
        {
            int rowCount = defectMapRowCount;
            int colCount = defectMapColCount;

            if (g == null || width <= 0 || height <= 0 || rowCount <= 0 || colCount <= 0)
            {
                return;
            }

            float cellWidth = width / (float)colCount;
            float cellHeight = height / (float)rowCount;
            Color gridColor = Color.FromArgb(96, 210, 210, 210); // 흐린 연회색
            Color textColor = Color.FromArgb(120, 185, 185, 185);

            using (Pen gridPen = new Pen(gridColor, 1f))
            using (Brush textBrush = new SolidBrush(textColor))
            using (Font textFont = new Font("Segoe UI", 8f, FontStyle.Regular))
            using (StringFormat textFormat = new StringFormat())
            {
                textFormat.Alignment = StringAlignment.Center;
                textFormat.LineAlignment = StringAlignment.Center;

                for (int i = 0; i <= colCount; i++)
                {
                    float x = i * cellWidth;
                    g.DrawLine(gridPen, x, 0, x, height);
                }

                for (int i = 0; i <= rowCount; i++)
                {
                    float y = i * cellHeight;
                    g.DrawLine(gridPen, 0, y, width, y);
                }

                if (defectMapRowColCount == 0)
                {

                    HashSet<string> ignoreIndexSet = new HashSet<string>(
     ignoreIndexArray.Select(x => x.ToUpper())
 );

                    List<string> availableRowLabels = new List<string>();

                    for (int i = 0; i < 26; i++)
                    {
                        string label = ((char)('A' + i)).ToString();

                        if (!ignoreIndexSet.Contains(label))
                        {
                            availableRowLabels.Add(label);
                        }
                    }

                    if (defectMapRowColCount == 0)
                    {
                        for (int row = 0; row < rowCount; row++)
                        {
                            int logicalRow = defectMapRowSwap == 1
                                ? row
                                : (rowCount - 1 - row);

                            if (logicalRow >= availableRowLabels.Count)
                                continue;

                            string rowLabel = availableRowLabels[logicalRow];

                            for (int col = 0; col < colCount; col++)
                            {
                                int colLabelIndex = defectMapColSwap == 1
                                    ? (col + 1)
                                    : (colCount - col);

                                string cellLabel = rowLabel + colLabelIndex.ToString();

                                RectangleF cellRect = new RectangleF(
                                    col * cellWidth,
                                    row * cellHeight,
                                    cellWidth,
                                    cellHeight);

                                g.DrawString(cellLabel, textFont, textBrush, cellRect, textFormat);
                            }
                        }
                    }
                }
                else
                {
                    HashSet<string> ignoreIndexSet = new HashSet<string>(
    ignoreIndexArray.Select(x => x.ToUpper())
);


                    List<string> availableColLabels = new List<string>();

                    for (int i = 0; i < 26; i++)
                    {
                        string label = ((char)('A' + i)).ToString();

                        if (!ignoreIndexSet.Contains(label))
                        {
                            availableColLabels.Add(label);
                        }
                    }

                    for (int row = 0; row < rowCount; row++)
                    {
                        int rowLabelIndex = defectMapRowSwap == 1
                            ? (row + 1)
                            : (rowCount - row);

                        for (int col = 0; col < colCount; col++)
                        {
                            int logicalCol = defectMapColSwap == 1
                                ? col
                                : (colCount - 1 - col);

                            if (logicalCol >= availableColLabels.Count)
                                continue;

                            string colLabel = availableColLabels[logicalCol];

                            string cellLabel = colLabel + rowLabelIndex.ToString();

                            RectangleF cellRect = new RectangleF(
                                col * cellWidth,
                                row * cellHeight,
                                cellWidth,
                                cellHeight);

                            g.DrawString(cellLabel, textFont, textBrush, cellRect, textFormat);
                        }
                    }
                }

            }
        }

        private void MTP_DEFECTSEARCH_Click(object sender, EventArgs e)
        {

        }

        private void groupBox2_Enter(object sender, EventArgs e)
        {

        }

        private void LV_PANEL_LIST_SelectedIndexChanged(object sender, EventArgs e)
        {

            //EXCEL EXPORT 하기위한 조건문 
            if (LV_PANEL_LIST.SelectedItems.Count > 0 && Excel_wirte_path != null)
                BTN_WRITE_EXCEL.BackColor = Color.ForestGreen;
            else
                BTN_WRITE_EXCEL.BackColor = Color.LightGray;

            System.Windows.Forms.ListView.SelectedIndexCollection indices = LV_PANEL_LIST.SelectedIndices;
            LB_SELECT_COUNT.Text = indices.Count.ToString();

            if (indices.Count > 0)
            {
                TrySyncSelectedPanelContext();
            }
        }

        private void LV_PANEL_LIST_DoubleClick(object sender, EventArgs e)
        {
            int vp_num_int;
            if (!TrySyncSelectedPanelContext(out vp_num_int))
            {
                MessageBox.Show("선택된 패널 정보를 찾을 수 없습니다.");
                return;
            }

            //해당 디펙을 상세하게 들어가기 위한 탭변경 초기화 
            TC_LADYBUG.SelectedIndex = 1;
            string panelid = insp_info.Pid;

            Console.WriteLine("select panelid : " + panelid);

            BTN_DEFECT_PANEL.Text = panelid;
            Validate();
            listview2_init(vp_num_int);
        }

        private bool TrySyncSelectedPanelContext()
        {
            int vpNumInt;
            return TrySyncSelectedPanelContext(out vpNumInt);
        }

        private bool TrySyncSelectedPanelContext(out int vpNumInt)
        {
            vpNumInt = 0;

            if (LV_PANEL_LIST.SelectedItems.Count == 0)
            {
                return false;
            }

            System.Windows.Forms.ListViewItem selectedItem = LV_PANEL_LIST.SelectedItems[0];
            if (selectedItem.SubItems.Count <= 6)
            {
                return false;
            }

            string selectPid = selectedItem.SubItems[1].Text;
            string vpNumber = selectedItem.SubItems[6].Text;
            if (string.IsNullOrEmpty(selectPid) || string.IsNullOrEmpty(vpNumber))
            {
                return false;
            }

            string vpSuffix = vpNumber.Substring(vpNumber.Length - 1);
            for (int i = 0; i < inspectionDataList.Count; i++)
            {
                string vpnum = inspectionDataList[i].Vpnum;
                if (string.IsNullOrEmpty(vpnum))
                {
                    continue;
                }

                string dataVpSuffix = vpnum.Substring(vpnum.Length - 1);
                if (inspectionDataList[i].SerialNumber == selectPid && dataVpSuffix == vpSuffix)
                {
                    insp_info.listview_index = i;
                    navi.Panel_index = i;
                    insp_info.insp_Data = inspectionDataList[i].InspectionDate;
                    insp_info.Pid = inspectionDataList[i].SerialNumber;
                    insp_info.Vision_Num = "VP0" + vpSuffix;

                    return int.TryParse(vpSuffix, out vpNumInt);
                }
            }

            return false;
        }

        public void listview2_init(int vp_num_int)
        {
            Feature_row.Clear();
            Feature_row_post.Clear();
            Defect_Position.Clear();
            Defect_Position_Judge.Clear();
            LV_DEFECT_LIST.Items.Clear();

            for (int vp_num_list_count = 1; vp_num_list_count < vp_info_list.Count + 1; vp_num_list_count++)
            {
                vp_num_int = vp_num_list_count;
                try//List Item 정보 뿌리기 
                {
                    //index 기준은 public으로 갖고있고 해당 정보를 토대로 listview 뿌리기 
                    //string basepath = Disk_base + result_path_base;
                    //int vp_num_int = Convert.ToInt32(insp_info.Vision_Num.Substring(3, 1));
                    string basepath = vp_info_list[vp_num_int - 1].Path_Vp_Result;
                    insp_info.Vision_Num = insp_info.Vision_Num.Substring(0, 3) + vp_num_int.ToString();
                    string img_path = basepath + "\\" + insp_info.insp_Data + "\\" + Model_name + "\\" + insp_info.Pid + "_" + insp_info.Vision_Num;
                    Console.WriteLine("basepath : " + basepath);
                    Console.WriteLine("img_path : " + img_path);

                    //string img_path = basepath + insp_info.insp_Data + "\\" + "x2292" + "\\" + insp_info.Pid + "_" + insp_info.Vision_Num;
                    string[] files = Directory.GetFiles(img_path, "*.csv");


                    //SP임시
                    if (false)
                    {
                        if (View_mode == 1)
                            Crop_bin_path_Pre[vp_num_list_count - 1] = Directory.GetFiles(img_path, "*.bin")[0];
                        else
                            Crop_bin_path_Post[vp_num_list_count - 1] = Directory.GetFiles(img_path, "*.bin")[0];
                    }
                    else
                    {
                        Crop_bin_path_Pre[vp_num_list_count - 1] = Directory.GetFiles(img_path, "*Pre*.bin")[0];
                        Crop_bin_path_Post[vp_num_list_count - 1] = Directory.GetFiles(img_path, "*Post*.bin")[0];
                    }



                    //RECIPE 읽어서 일단 크기좀알아내자 
                    using (StreamReader file = File.OpenText(vp_info_list[0].Path_Vp_Recipe + "\\" + CBB_RECIPE.Text + "\\" + "Recipe.json"))
                    {
                        CBB_SINGLE_RECIPE.Items.Add("ALL");
                        using (JsonTextReader reader = new JsonTextReader(file))
                        {
                            JObject jsondata = (JObject)JToken.ReadFrom(reader);

                            insp_info.panel_width = Convert.ToInt32(jsondata["INSP_INFO"]["PANEL_SIZE_WIDHT"].ToString());
                            insp_info.panel_Height = Convert.ToInt32(jsondata["INSP_INFO"]["PANEL_SIZE_HEIGHT"].ToString());
                        }
                    }

                    //컨트롤의 가로 = 패널의 세로   세로는 가로 일단 배율부터 찾자
                    picturebox_ratio_x = (double)insp_info.panel_width / PB_DEFECTMAP.Width;
                    picturebox_ratio_y = (double)insp_info.panel_Height / PB_DEFECTMAP.Height;

                    if (files.Length > 0)
                    {
                        string pre_path = "";
                        string post_path = "";
                        //여기다가 pre정보 넣어줘야함 
                        for (int i = 0; i < files.Length; i++)
                        {
                            if (files[i].Contains("Pre"))
                                pre_path = files[i];

                            if (files[i].Contains("Post"))
                                post_path = files[i];
                        }
                        listview2_ReadCSV(pre_path, post_path, vp_num_list_count);
                    }
                }
                catch (Exception e)
                {
                    Console.Write(e.Message);
                }
            }
            LV_DEFECT_LIST.BeginUpdate();
            try
            {
                if (View_mode == 1)//Pre
                {
                    for (int i = 0; i < Feature_row.Count; i++)
                    {
                        var listViewItem = new System.Windows.Forms.ListViewItem((i + 1).ToString());
                        for (int j = 0; j < VIEW_DEFECT_FEATURE_NAME.Count; j++)
                            listViewItem.SubItems.Add(Feature_row[i][VIEW_DEFECT_FEATURE_NAME[j]]?.ToString());
                        LV_DEFECT_LIST.Items.Add(listViewItem);
                    }
                }
                else   //Post
                {
                    for (int i = 0; i < Feature_row_post.Count; i++)
                    {
                        var listViewItem = new System.Windows.Forms.ListViewItem((i + 1).ToString());
                        for (int j = 0; j < VIEW_DEFECT_FEATURE_NAME.Count; j++)
                            listViewItem.SubItems.Add(Feature_row_post[i][VIEW_DEFECT_FEATURE_NAME[j]]?.ToString());
                        LV_DEFECT_LIST.Items.Add(listViewItem);
                    }
                }
            }
            finally
            {
                LV_DEFECT_LIST.EndUpdate();
            }
            //그뒤에  PB에 좌표 투영
            // PictureBox의 크기와 동일한 Bitmap 생성
            Bitmap bitmap = new Bitmap((int)(insp_info.panel_width / picturebox_ratio_x), (int)(insp_info.panel_Height / picturebox_ratio_y));
            using (Graphics gridGraphics = Graphics.FromImage(bitmap))
            {
                DrawDefectMapGrid(gridGraphics, bitmap.Width, bitmap.Height);
            }

            // Graphics 객체를 사용하여 Bitmap에 그리기
            //using (Graphics g = Graphics.FromImage(bitmap))
            //{
            //    for(int i=0;i< Feature_row.Count;i++) 
            //    {
            //        g.FillEllipse(Brushes.White, (int)(Convert.ToInt32(Feature_row[i]["PIXEL_Y"]?.ToString())/ picturebox_ratio_x), (int)(Convert.ToInt32(Feature_row[i]["PIXEL_X"]?.ToString())/ picturebox_ratio_y), 5, 5); // 반지름 5짜리 원
            //    }
            //}

            PB_DEFECTMAP.Invalidate();
            // PictureBox에 Bitmap 설정
            SetPictureBoxImage(PB_DEFECTMAP, bitmap);
            //Classity_Read(recipe_path);
            //Classity_Read(Disk_base + recipe_path_base + CBB_RECIPE.Text.ToString() + @"\Classifier.json");
            Classity_Read(vp_info_list[0].Path_Vp_Recipe + @"\" + CBB_RECIPE.Text + @"\Classifier.json");

        }

        private void listview2_ReadCSV(string pre_path, string post_path, int vpnum)
        {
            //Feature_row.Clear();
            //Feature_row_post.Clear();
            //Defect_Position.Clear();




            //정보읽기-Feature
            if (true)
            {
                int Swap_X_POS, Swap_Y_POS;
                int header_check_bit = 0;
                string[] header = null;

                //pre
                foreach (string line in File.ReadLines(pre_path))
                {
                    //Console.Write("2-파일 위치 : " + path);
                    if (header_check_bit == 0)
                    {
                        header_check_bit++;
                        header = line.Split(',');
                        continue;
                    }
                    Dictionary<string, object> Feature_defect = new Dictionary<string, object>();
                    var values = line.Split(',');
                    for (int i = 0; i < values.Length; i++)
                    {
                        //Feature_row[header[i]] = values[i].ToString();
                        Feature_defect.Add(header[i], values[i]);
                    }
                    Feature_row.Add(Feature_defect);
                    VP_Defect_num[vpnum - 1] = Feature_row.Count;
                    //System.Drawing.Point point = new System.Drawing.Point((int)(Convert.ToInt32(Feature_defect["PIXEL_Y"]?.ToString()) / picturebox_ratio_x), (int)(Convert.ToInt32(Feature_defect["PIXEL_X"]?.ToString()) / picturebox_ratio_y));
                    if (View_mode == 1)
                    {
                        if (Swap_X == 1)
                            Swap_X_POS = (int)((insp_info.panel_width - Convert.ToInt32(Feature_defect[POSITION_FEATURE_NAME_X]?.ToString())) / picturebox_ratio_x);
                        else
                            Swap_X_POS = (int)(Convert.ToInt32(Feature_defect[POSITION_FEATURE_NAME_X]?.ToString()) / picturebox_ratio_x);
                        if (Swap_Y == 1)
                            Swap_Y_POS = (int)((insp_info.panel_Height - Convert.ToInt32(Feature_defect[POSITION_FEATURE_NAME_Y]?.ToString())) / picturebox_ratio_y);
                        else
                            Swap_Y_POS = (int)(Convert.ToInt32(Feature_defect[POSITION_FEATURE_NAME_Y]?.ToString()) / picturebox_ratio_y);
                        System.Drawing.Point point = new System.Drawing.Point(Swap_X_POS, Swap_Y_POS);
                        Defect_Position.Add(point);

                        //여기는 judge부 
                        string Defect_position_judge = Feature_defect["DEFECT_JUDGE"].ToString();
                        if (Defect_position_judge == "OK")
                            Defect_Position_Judge.Add(true);
                        else
                            Defect_Position_Judge.Add(false);
                    }

                }
            }

            //post
            if (true)
            {
                //정보읽기-Feature
                int header_check_bit = 0;
                string[] header = null;

                foreach (string line in File.ReadLines(post_path))
                {
                    //Console.Write("2-파일 위치 : " + path);
                    if (header_check_bit == 0)
                    {
                        header_check_bit++;
                        header = line.Split(',');
                        continue;
                    }
                    Dictionary<string, object> Feature_defect = new Dictionary<string, object>();
                    var values = line.Split(',');
                    for (int i = 0; i < values.Length; i++)
                    {
                        //Feature_row[header[i]] = values[i].ToString();
                        Feature_defect.Add(header[i], values[i]);
                    }
                    Feature_row_post.Add(Feature_defect);
                    //System.Drawing.Point point = new System.Drawing.Point((int)(Convert.ToInt32(Feature_defect["PIXEL_Y"]?.ToString()) / picturebox_ratio_x), (int)(Convert.ToInt32(Feature_defect["PIXEL_X"]?.ToString()) / picturebox_ratio_y));
                    if (View_mode == 2)
                    {
                        System.Drawing.Point point = new System.Drawing.Point((int)(Convert.ToInt32(Feature_defect[POSITION_FEATURE_NAME_X]?.ToString()) / picturebox_ratio_x), (int)(Convert.ToInt32(Feature_defect[POSITION_FEATURE_NAME_Y]?.ToString()) / picturebox_ratio_y));
                        Defect_Position.Add(point);

                        //여기는 judge부 
                        string Defect_position_judge = Feature_defect["DEFECT_JUDGE"].ToString();
                        if (Defect_position_judge == "OK")
                            Defect_Position_Judge.Add(true);
                        else
                            Defect_Position_Judge.Add(false);


                    }
                }
            }

            if (false)
            {
                for (int i = 0; i < Feature_row.Count; i++)
                {
                    var listViewItem = new System.Windows.Forms.ListViewItem((i + 1).ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["DefectJudge"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["PRE_CLASSIFY_GROUP"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["PIXEL_X"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["PIXEL_Y"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["Area"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["GrayAVG_Pre"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["GrayMin_Pre"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["GrayMax_Pre"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["OriginalImage_Column"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["OriginalImage_Row"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["PTNType"]?.ToString());
                    LV_DEFECT_LIST.Items.Add(listViewItem);
                }
            }
            //if(true)
            //{
            //    if (View_mode == 1)//Pre
            //    {
            //        for (int i = 0; i < Feature_row.Count; i++)
            //        {
            //            var listViewItem = new System.Windows.Forms.ListViewItem((i + 1).ToString());
            //            for (int j = 0; j < VIEW_DEFECT_FEATURE_NAME.Count; j++)
            //                listViewItem.SubItems.Add(Feature_row[i][VIEW_DEFECT_FEATURE_NAME[j]]?.ToString());
            //            LV_DEFECT_LIST.Items.Add(listViewItem);
            //        }
            //    }
            //    else   //Post
            //    {
            //        for (int i = 0; i < Feature_row_post.Count; i++)
            //        {
            //            var listViewItem = new System.Windows.Forms.ListViewItem((i + 1).ToString());
            //            for (int j = 0; j < VIEW_DEFECT_FEATURE_NAME.Count; j++)
            //                listViewItem.SubItems.Add(Feature_row_post[i][VIEW_DEFECT_FEATURE_NAME[j]]?.ToString());
            //            LV_DEFECT_LIST.Items.Add(listViewItem);
            //        }
            //    }
            //}
        }

        private void listview2_ReadCSV_Only(string path)
        {
            Feature_row.Clear();

            //정보읽기-Feature
            int header_check_bit = 0;
            string[] header = null; ;
            foreach (string line in File.ReadLines(path))
            {
                if (header_check_bit == 0)
                {
                    header_check_bit++;
                    header = line.Split(',');
                    continue;
                }
                Dictionary<string, object> Feature_defect = new Dictionary<string, object>();
                var values = line.Split(',');
                for (int i = 0; i < values.Length; i++)
                {
                    //Feature_row[header[i]] = values[i].ToString();
                    Feature_defect.Add(header[i], values[i]);
                }

                Feature_row.Add(Feature_defect);
            }
        }

        private void Classity_Read(string path)
        {
            //클래시파이 정보 뿌리기
            CBB_SINGLE_RECIPE.Items.Clear();
            CBB_SINGLE_RECIPE_COPY.Items.Clear();
            // Json 파일 읽기
            //string Model_name = CBB_RECIPE.Text;
            //BTN_RECIPE.Text = Model_name;

            //string recipe_path = Disk_base +  recipe_path_base + Model_name + @"\Classifier.json";

            classify_Pre_Infos.Clear();
            classify_Post_Infos.Clear();

            using (StreamReader file = File.OpenText(path))
            {
                CBB_SINGLE_RECIPE.Items.Add("ALL");
                CBB_SINGLE_RECIPE_COPY.Items.Add("ALL");
                using (JsonTextReader reader = new JsonTextReader(file))
                {
                    //JArray json = (JArray)JToken.ReadFrom(reader);


                    //PRE부터 저장
                    JObject json_PRE = (JObject)(JToken.ReadFrom(reader));
                    JArray json_pre = (JArray)json_PRE["PRE_CLASSIFIER"];
                    JArray json_post = (JArray)json_PRE["POST_CLASSIFIER"];


                    // Config
                    for (int i = 0; i < json_pre.Count; i++)
                    {
                        Classify_Pre_info class_info = new Classify_Pre_info();
                        class_info.SCRIPT_NAME = json_pre[i]["SCRIPT_NAME"].ToString();
                        class_info.JUDGE = json_pre[i]["JUDGE"].ToString();
                        class_info.PRIORITY = json_pre[i]["PRIORITY"].ToString();
                        class_info.SCRIPT = json_pre[i]["SCRIPT"].ToString();
                        classify_Pre_Infos.Add(class_info);
                        if (View_mode == 1)
                        {
                            CBB_SINGLE_RECIPE.Items.Add(class_info.SCRIPT_NAME);
                            CBB_SINGLE_RECIPE_COPY.Items.Add(class_info.SCRIPT_NAME);
                        }
                    }

                    for (int i = 0; i < json_post.Count; i++)
                    {
                        Classify_Post_info class_info = new Classify_Post_info();
                        class_info.SCRIPT_NAME = json_post[i]["SCRIPT_NAME"].ToString();
                        class_info.JUDGE = json_post[i]["JUDGE"].ToString();
                        class_info.BYPASS = json_post[i]["BYPASS"].ToString();
                        class_info.DISTANCE = json_post[i]["DISTANCE"].ToString();
                        class_info.MIN_POINTS = json_post[i]["MIN_POINTS"].ToString();
                        class_info.MERGE_DEFECT = json_post[i]["MERGE_DEFECT"].ToString();
                        class_info.DEFECT_TYPE = json_post[i]["DEFECT_TYPE"].ToString();
                        class_info.SCRIPT_NAME_LIST = json_post[i]["SCRIPT_NAME_LIST"].ToString();
                        class_info.SCRIPT = json_post[i]["SCRIPT"].ToString();
                        classify_Post_Infos.Add(class_info);
                        if (View_mode == 2)
                        {
                            CBB_SINGLE_RECIPE.Items.Add(class_info.SCRIPT_NAME);
                            CBB_SINGLE_RECIPE_COPY.Items.Add(class_info.SCRIPT_NAME);
                        }
                    }

                }
            }
            LB_RECIPE_PATH.Text = vp_info_list[0].Path_Vp_Recipe;
        }

        private void Classity_Read_DefectName(string path)
        {
            //클래시파이 정보 뿌리기
            using (StreamReader file = File.OpenText(path))
            {
                using (JsonTextReader reader = new JsonTextReader(file))
                {
                    //JArray json = (JArray)(JToken.ReadFrom(reader));
                    JObject json_PRE = (JObject)(JToken.ReadFrom(reader));
                    JArray json = (JArray)json_PRE["PRE_CLASSIFIER"];

                    // Config
                    for (int i = 0; i < json.Count; i++)
                    {
                        CBB_SINGLE_RECIPE.Items.Add(json[i]["SCRIPT_NAME"].ToString());
                    }
                }
            }
        }





        private void dtpDateFrom_ValueChanged(object sender, EventArgs e)
        {
            StartDate = dtpDateFrom.Value;
            StartDate = new DateTime(StartDate.Year, StartDate.Month, StartDate.Day, 0, 0, 0);
            dtpDateFrom.Value = StartDate;
        }

        private void dtpDateTo_ValueChanged(object sender, EventArgs e)
        {
            EndDate = dtpDateTo.Value;
            EndDate = new DateTime(EndDate.Year, EndDate.Month, EndDate.Day, 23, 59, 59);
            dtpDateTo.Value = EndDate;
        }

        private void BTN_SEARCH_Click(object sender, EventArgs e)
        {
            //모델네임 확인
            if (CB_RECIPE_WRITE.Checked) // 만약 직접입력창이 활성화되었을 경우
                Model_name = TB_RECIPE.Text.ToString();
            else  //라디오버튼으로 나눠서 사용할때 
            {
                if (RB_RECIPE_FRONT.Checked) // front 인경우
                    Model_name = CBB_RECIPE.Text.Substring(0, Convert.ToInt32(TB_RECIPE.Text));   //앞에서 정해진 크기만큼 짜르기
                else
                    Model_name = CBB_RECIPE.Text.Substring(CBB_RECIPE.Text.Length - Convert.ToInt32(TB_RECIPE.Text));   //앞에서 정해진 크기만큼 짜르기
            }
            LB_RECIPE.Text = Model_name;

            insp_info.Recipe_Model = vp_info_list[0].Path_Vp_Recipe + "\\" + Model_name;


            //날짜별 검색
            try
            {
                DateTime StartDate_CHECK = ParseHHMMSS(StartDate, TB_STARTTIME.Text);
                DateTime EndDate_CHECK = ParseHHMMSS(EndDate, TB_ENDTIME.Text);

                bool useJudgeFilter = CBB_JUDGE.SelectedIndex != 0;
                bool useDefectFilter = CBB_DEFECT_NAME.SelectedIndex != 0;
                string judgeFilter = CBB_JUDGE.Text;
                string defectFilter = CBB_DEFECT_NAME.Text;

                matchingFiles.Clear();

                for (int i = 0; i < vp_info_list.Count; i++)
                {
                    IEnumerable<string> files = Directory.EnumerateFiles(vp_info_list[i].Path_Vp_Log, "*.csv", SearchOption.TopDirectoryOnly);
                    foreach (string file in files)
                    {
                        string fileName = System.IO.Path.GetFileNameWithoutExtension(file);
                        if (string.IsNullOrEmpty(fileName) || fileName.Length < 8)
                        {
                            continue;
                        }

                        string datePart = fileName.Substring(0, 8);
                        DateTime fileDate;
                        if (DateTime.TryParseExact(datePart, "yyyyMMdd", null, DateTimeStyles.None, out fileDate))
                        {
                            if (fileDate >= StartDate && fileDate <= EndDate)
                            {
                                matchingFiles.Add(file);
                            }
                        }
                    }
                }

                inspectionDataList.Clear();

                for (int i = 0; i < matchingFiles.Count; i++)
                {
                    string logFilePath = matchingFiles[i];
                    string fileName = System.IO.Path.GetFileNameWithoutExtension(logFilePath);
                    string[] fileNameTokens = fileName.Split('_');
                    string vpnum = fileNameTokens.Length > 2 ? fileNameTokens[2] : string.Empty;

                    bool isHeader = true;
                    foreach (string line in File.ReadLines(logFilePath))
                    {
                        if (isHeader)
                        {
                            isHeader = false;
                            continue;
                        }

                        if (string.IsNullOrWhiteSpace(line))
                        {
                            continue;
                        }

                        var values = line.Split(',');
                        if (values.Length < 5)
                        {
                            continue;
                        }

                        InspectionData inspectionData = new InspectionData();
                        inspectionData.SerialNumber = values[0];
                        inspectionData.InspectionDate = values[1];
                        inspectionData.InspectionTime = values[2];
                        inspectionData.FinalJudgment = values[3];
                        inspectionData.ClassifyGroup = values[4];
                        inspectionData.Vpnum = vpnum;

                        DateTime SelectDate_CHECK = ParseHHMMSS(inspectionData.InspectionDate, inspectionData.InspectionTime);
                        if (SelectDate_CHECK < StartDate_CHECK || EndDate_CHECK < SelectDate_CHECK)
                        {
                            continue;
                        }

                        if (useJudgeFilter && inspectionData.FinalJudgment != judgeFilter)
                        {
                            continue;
                        }

                        if (useDefectFilter && inspectionData.ClassifyGroup != defectFilter)
                        {
                            continue;
                        }

                        inspectionDataList.Add(inspectionData);
                    }
                }

                listview_print(inspectionDataList);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }




        private void BTN_IDSEARCH_Click(object sender, EventArgs e)
        {

            if (!string.IsNullOrEmpty(TB_IDSEARCH.Text))
            {
                // panel_id 리스트 생성
                List<string> panel_id = new List<string>();
                if (TB_IDSEARCH.Text.Contains("|") || TB_IDSEARCH.Text.Contains("\n") || TB_IDSEARCH.Text.Contains("\r"))
                {
                    panel_id = TB_IDSEARCH.Text.Split(new[] { '|', '\n', '\r' }, StringSplitOptions.RemoveEmptyEntries)
                                              .Select(s => s.Trim())
                                              .ToList();
                }
                else
                {
                    panel_id.Add(TB_IDSEARCH.Text.Trim());
                }
                LV_PANEL_LIST.Items.Clear();
                // panel_id 리스트 내의 각 ID에 대해 검색 수행
                foreach (string id in panel_id)
                {
                    List<InspectionData> filteredList = inspectionDataList.Where(data => data.SerialNumber != null && data.SerialNumber.Contains(id)).ToList();
                    listview_print_noclear(filteredList);
                }

                BTN_IDSEARCH.Text = "▶";
            }
            else
            {
                listview_print(inspectionDataList);
                BTN_IDSEARCH.Text = "▶";
            }


        }

        private void listview_print(List<InspectionData> data)
        {
            LV_PANEL_LIST.BeginUpdate();
            try
            {
                LV_PANEL_LIST.Items.Clear();
                if (data.Count == 0)
                {
                    LB_TOTAL_COUNT.Text = "0";
                    return;
                }

                ListViewItem[] items = new ListViewItem[data.Count];
                for (int i = 0; i < data.Count; i++)
                {
                    items[i] = CreatePanelListViewItem(data[i], i);
                }

                LV_PANEL_LIST.Items.AddRange(items);
                LB_TOTAL_COUNT.Text = data.Count.ToString();
            }
            finally
            {
                LV_PANEL_LIST.EndUpdate();
            }
        }

        private void listview_print_noclear(List<InspectionData> data)
        {
            LV_PANEL_LIST.BeginUpdate();
            try
            {
                int baseIndex = LV_PANEL_LIST.Items.Count;
                if (data.Count == 0)
                {
                    LB_TOTAL_COUNT.Text = LV_PANEL_LIST.Items.Count.ToString();
                    return;
                }

                ListViewItem[] items = new ListViewItem[data.Count];
                for (int i = 0; i < data.Count; i++)
                {
                    items[i] = CreatePanelListViewItem(data[i], baseIndex + i);
                }

                LV_PANEL_LIST.Items.AddRange(items);
                LB_TOTAL_COUNT.Text = LV_PANEL_LIST.Items.Count.ToString();
            }
            finally
            {
                LV_PANEL_LIST.EndUpdate();
            }
        }

        private ListViewItem CreatePanelListViewItem(InspectionData data, int index)
        {
            ListViewItem item = new ListViewItem(index.ToString());
            item.SubItems.Add(data.SerialNumber);
            item.SubItems.Add(data.InspectionDate);
            item.SubItems.Add(data.InspectionTime);
            item.SubItems.Add(data.FinalJudgment);
            item.SubItems.Add(data.ClassifyGroup);
            item.SubItems.Add(data.Vpnum);
            item.SubItems.Add("");
            item.SubItems.Add("");
            return item;
        }


        //--UTIL --------------------------------
        public string[] GetSearchFolder_all(String _strPath)
        {
            string[] files = { "", };
            try
            {
                // TopDirectoryOnly
                files = Directory.GetDirectories(_strPath, "*.*", SearchOption.AllDirectories);
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message);
            }
            return files;
        }                   //폴더내 폴더검색 _하위폴더 포함
        public string[] GetSearchFolder_only(String _strPath)
        {
            string[] files = { "", };
            try
            {
                // TopDirectoryOnly
                files = Directory.GetDirectories(_strPath, "*.*", SearchOption.TopDirectoryOnly);
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message);
            }
            return files;
        }                  //폴더내 폴더검색 _해당폴더내
        public string[] GetSearchFile_only(string _strPath)
        {
            string[] files = { "", };
            try
            {
                // TopDirectoryOnly
                files = Directory.GetFiles(_strPath, "*.*", SearchOption.TopDirectoryOnly);
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message);
            }
            return files;
        }
        public string[] GetSearchFile_only_recipe(string _strPath)
        {
            string[] files = { "", };
            try
            {
                // TopDirectoryOnly
                files = Directory.GetFiles(_strPath, @"Recipe*.json", SearchOption.TopDirectoryOnly);
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message);
            }
            return files;
        }
        public string[] GetSearchFile_only_classify(string _strPath)
        {
            string[] files = { "", };
            try
            {
                // TopDirectoryOnly
                files = Directory.GetFiles(_strPath, @"Classifier*.json", SearchOption.TopDirectoryOnly);
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message);
            }
            return files;
        }

        static DateTime ParseHHMMSS(string date, string hhmmss)
        {
            int year = int.Parse(date.Substring(0, 4));
            int month = int.Parse(date.Substring(4, 2));
            int day = int.Parse(date.Substring(6, 2));

            int hour = int.Parse(hhmmss.Substring(0, 2));
            int minute = int.Parse(hhmmss.Substring(3, 2));
            int second = int.Parse(hhmmss.Substring(6, 2));

            return new DateTime(year, month, day,
                                hour, minute, second);
        }
        static DateTime ParseHHMMSS(DateTime date, string hhmmss)
        {
            int hour = int.Parse(hhmmss.Substring(0, 2));
            int minute = int.Parse(hhmmss.Substring(2, 2));
            int second = int.Parse(hhmmss.Substring(4, 2));

            return new DateTime(date.Year, date.Month, date.Day,
                                hour, minute, second);
        }

        private void CBB_RECIPE_SelectedIndexChanged(object sender, EventArgs e)
        {
            string sub_recipe;
            if (CBB_RECIPE.Text.Length < 5)
                sub_recipe = CBB_RECIPE.Text;
            else
                sub_recipe = CBB_RECIPE.Text.Substring(CBB_RECIPE.Text.Length - 5);
            Classity_Read_DefectName(vp_info_list[0].Path_Vp_Recipe + @"\" + CBB_RECIPE.Text + @"\Classifier.json");


        }

        private void CBB_SINGLE_RECIPE_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void BTN_DEFECT_JUDGE_Click(object sender, EventArgs e)
        {
            // PRE
            //PRE 용 DLG 셋팅 
            //LISTVIEW 초기화 다시하기
            LV_DEFECT_LIST_SelectedIndexChangedAsync(this, new EventArgs());
            string[] change_word = { };
            string classify_name = CBB_SINGLE_RECIPE.Text.ToString();
            Classify_Pre_info select_classify_pre;
            select_classify_pre.SCRIPT = "";

            Classify_Post_info select_classify_post;
            select_classify_post.SCRIPT = "";


            //현재 쓰고있는 FEATURE만 보여주기위한 List변수
            used_Feature_Pre.Clear();
            used_Feature_Post.Clear();


            //PRE->POST TRUE면 검증 
            if (true)
            {


                //Post 사용여부 
                bool use_Post = false;

                for (int i = 0; i < classify_Pre_Infos.Count; i++)
                {
                    if (classify_Pre_Infos[i].SCRIPT_NAME == classify_name)
                    {
                        select_classify_pre = classify_Pre_Infos[i];

                        for (int j = 0; j < classify_Post_Infos.Count; j++)
                        {
                            if (classify_Post_Infos[j].SCRIPT_NAME == classify_name)
                                if (classify_Post_Infos[j].BYPASS == "True")
                                {
                                    select_classify_post = classify_Post_Infos[j];
                                    use_Post = true;
                                }
                        }
                    }
                }
                List<string> NG_str_PRE = new List<string>();
                if (LV_DEFECT_LIST.SelectedIndices.Count > 0)
                {
                    string[] Origin_word = select_classify_pre.SCRIPT.Split(' ');
                    int selectedIndex;
                    // 선택된 인덱스를 사용하여 작업 수행
                    selectedIndex = LV_DEFECT_LIST.SelectedIndices[0];

                    System.Windows.Forms.ListViewItem selectedItem = LV_DEFECT_LIST.Items[selectedIndex];
                    // 특정 열의 값 가져오기 (예: 첫 번째 열 = 인덱스 0)
                    int columnValue = Convert.ToInt32(selectedItem.SubItems[0].Text) - 1;
                    selectedIndex = columnValue;
                    try // Defect Featrue 출력 
                    {
                        int index = 0;
                        string[] words = select_classify_pre.SCRIPT.Split(' ');
                        foreach (var kvp in Feature_row[selectedIndex])
                        {
                            //select_classify.SCRIPT=select_classify.SCRIPT.Replace(kvp.Key.ToString(), kvp.Value.ToString());
                            if (words.Contains(kvp.Key))
                            {
                                bool push = true;
                                for (int i = 0; i < words.Length; i++)
                                {
                                    if (words[i] == kvp.Key)
                                    {
                                        if (push)
                                        {
                                            used_Feature_Pre.Add(kvp.Key.ToString());
                                            push = false;
                                        }
                                        words[i] = kvp.Value.ToString();
                                    }
                                }
                            }
                        }
                        for (int i = 0; i < words.Length; i++)
                        {
                            if (words[i] == "and")
                            {
                                words[i] = "&&";
                            }
                            if (words[i] == "or")
                            {
                                words[i] = "||";
                            }
                        }
                        change_word = words;
                        //Array.Reverse(words);
                        string reversedString = string.Join(" ", words);
                        CMergeClassify classifier = new CMergeClassify();
                        int result = classifier.ComputeExpression(reversedString);
                        Console.WriteLine($"PRE - Result: {result}"); // 출력: Result: 1
                        FindIncorrectPart(reversedString, NG_str_PRE);
                        LB_DEFECT_JUDGE.UseStyleColors = false; // 사용자 정의 색상을 사용하

                        if (result == 1)
                        {
                            LB_DEFECT_JUDGE.UseStyleColors = true; // 사용자 정의 색상을 사용하
                                                                   //LB_DEFECT_JUDGE.ForeColor = System.Drawing.Color.Green; // 예: 파란색
                            LB_DEFECT_JUDGE.Text = "PRE - Classify 조건식 : 참";
                        }
                        if (result == 0)
                        {
                            LB_DEFECT_JUDGE.UseStyleColors = true; // 사용자 정의 색상을 사용하
                                                                   //LB_DEFECT_JUDGE.ForeColor = System.Drawing.Color.Red; // 예: 파란색
                            LB_DEFECT_JUDGE.Text = "PRE - Classify 조건식 : 불";
                        }
                        if (result == 2)
                        {
                            LB_DEFECT_JUDGE.ForeColor = System.Drawing.Color.Blue; // 예: 파란색
                                                                                   //LB_DEFECT_JUDGE.UseStyleColors = true; // 사용자 정의 색상을 사용하
                            LB_DEFECT_JUDGE.Text = "PRE - Classify 조건식 : Error";
                        }

                        RTB_ORIGIN_PRE.Text = string.Join(" ", Origin_word);
                        RTB_REPLACE_PRE.Text = string.Join(" ", change_word);

                        List<string> preErrorSegments = GetLogicalErrorSegments(reversedString);
                        HighlightSegmentsInRichTextBox(RTB_REPLACE_PRE, RTB_REPLACE_PRE.Text, preErrorSegments);
                    }
                    catch
                    {
                        Console.WriteLine("Script Judge Error");
                    }
                    ApplyFilter(used_Feature_Pre);

                    // 문자열 단위가 아닌 논리 구조(식) 단위로 오류 부분 강조
                }

                //NG 항목대하여 색처리 하기위해서?
                //기존 REPLACE 대상에 대하여 NG_STR 찾기




                if (use_Post) //POST 검증
                {
                    BTN_CHANGE.BackColor = Color.Yellow;

                    //Post 경로 추가 
                    List<string> NG_str_Post = new List<string>();
                    if (LV_DEFECT_LIST.SelectedIndices.Count > 0)
                    {
                        string[] Origin_word = select_classify_post.SCRIPT.Split(' ');
                        int selectedIndex;
                        // 선택된 인덱스를 사용하여 작업 수행
                        selectedIndex = LV_DEFECT_LIST.SelectedIndices[0];

                        System.Windows.Forms.ListViewItem selectedItem = LV_DEFECT_LIST.Items[selectedIndex];
                        ////
                        //var key_PIXEL_X = Feature_row[selectedIndex]["PIXEL_X"];
                        //var key_PIXEL_Y = Feature_row[selectedIndex]["PIXEL_Y"];
                        //var key_BTS3 = Feature_row[selectedIndex]["BTS3"];

                        //// 인덱스 찾기
                        //int index_ = Feature_row_post.FindIndex(dict => dict.ContainsKey("PIXEL_X") && dict["PIXEL_X"].Equals(key_PIXEL_X) && dict.ContainsKey("PIXEL_Y") && dict["PIXEL_Y"].Equals(key_PIXEL_Y) && dict.ContainsKey("BTS3") && dict["BTS3"].Equals(key_BTS3));


                        // 특정 열의 값 가져오기 (예: 첫 번째 열 = 인덱스 0)
                        int columnValue = Convert.ToInt32(selectedItem.SubItems[0].Text) - 1;
                        selectedIndex = columnValue;
                        try // Defect Featrue 출력 
                        {
                            int index = 0;
                            string[] words = select_classify_post.SCRIPT.Split(' ');
                            foreach (var kvp in Feature_row_post[columnValue])
                            {
                                //select_classify.SCRIPT=select_classify.SCRIPT.Replace(kvp.Key.ToString(), kvp.Value.ToString());
                                if (words.Contains(kvp.Key))
                                {
                                    bool push = true;
                                    for (int i = 0; i < words.Length; i++)
                                    {
                                        if (words[i] == kvp.Key)
                                        {
                                            if (push)
                                            {
                                                used_Feature_Post.Add(kvp.Key.ToString());
                                                push = false;
                                            }
                                            words[i] = kvp.Value.ToString();
                                        }
                                    }
                                }
                            }
                            for (int i = 0; i < words.Length; i++)
                            {
                                if (words[i] == "and")
                                {
                                    words[i] = "&&";
                                }
                                if (words[i] == "or")
                                {
                                    words[i] = "||";
                                }
                            }
                            change_word = words;
                            //Array.Reverse(words);
                            string reversedString = string.Join(" ", words);
                            CMergeClassify classifier = new CMergeClassify();
                            int result = classifier.ComputeExpression(reversedString);
                            Console.WriteLine($"POST - Result: {result}"); // 출력: Result: 1
                            FindIncorrectPart(reversedString, NG_str_Post);
                            LB_DEFECT_JUDGE.UseStyleColors = false; // 사용자 정의 색상을 사용하

                            if (result == 1)
                            {
                                LB_DEFECT_JUDGE.UseStyleColors = true; // 사용자 정의 색상을 사용하
                                                                       //LB_DEFECT_JUDGE.ForeColor = System.Drawing.Color.Green; // 예: 파란색
                                LB_DEFECT_JUDGE.Text = "POST - Classify 조건식 : 참";
                            }
                            if (result == 0)
                            {
                                LB_DEFECT_JUDGE.UseStyleColors = true; // 사용자 정의 색상을 사용하
                                                                       //LB_DEFECT_JUDGE.ForeColor = System.Drawing.Color.Red; // 예: 파란색
                                LB_DEFECT_JUDGE.Text = "POST - Classify 조건식 : 불";
                            }
                            if (result == 2)
                            {
                                LB_DEFECT_JUDGE.ForeColor = System.Drawing.Color.Blue; // 예: 파란색
                                                                                       //LB_DEFECT_JUDGE.UseStyleColors = true; // 사용자 정의 색상을 사용하
                                LB_DEFECT_JUDGE.Text = "POST - Classify 조건식 : Error";
                            }

                            RTB_ORIGIN_POST.Text = string.Join(" ", Origin_word);
                            RTB_REPLACE_POST.Text = string.Join(" ", change_word);

                            List<string> postErrorSegments = GetLogicalErrorSegments(reversedString);
                            HighlightSegmentsInRichTextBox(RTB_REPLACE_POST, RTB_REPLACE_POST.Text, postErrorSegments);
                        }
                        catch
                        {
                            Console.WriteLine("Script Judge Error");
                        }
                        ApplyFilter(used_Feature_Post);

                        // 문자열 단위가 아닌 논리 구조(식) 단위로 오류 부분 강조

                    }

                }
            }



        }
        public void SetClassifierPrePost(int mode)
        {
            if (mode == 1) //pre mode 
            {
                RTB_ORIGIN_PRE.Show();
                RTB_ORIGIN_POST.Hide();

                RTB_REPLACE_PRE.Show();
                RTB_REPLACE_POST.Hide();

                label13.Show();
                label16.Hide();
            }
            else
            {
                RTB_ORIGIN_PRE.Hide();
                RTB_ORIGIN_POST.Show();

                RTB_REPLACE_PRE.Hide();
                RTB_REPLACE_POST.Show();

                label13.Hide();
                label16.Show();
            }
        }

        public void FindIncorrectPart(string expression, List<string> NG_str)
        {
            // Split the expression into parts
            expression = expression.Replace("&&", "&");
            expression = expression.Replace("||", "|");
            string[] parts_info = expression.Split(new char[] { '(', ')', '&', '|' }, StringSplitOptions.RemoveEmptyEntries);
            string[] parts = RemoveEmptyStrings(parts_info);
            foreach (string part in parts)
            {
                try
                {
                    NCalc.Expression nCalcPart = new NCalc.Expression(part);
                    bool partResult = (bool)nCalcPart.Evaluate();
                    if (!partResult)
                    {
                        Console.WriteLine($"The incorrect part is: {part}");
                        NG_str.Add(part);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error evaluating part: {part}. Error: {ex.Message}");
                }
            }
        }

        private List<string> GetLogicalErrorSegments(string expression)
        {
            List<string> result = new List<string>();
            CollectLogicalErrorSegments(expression, result);
            return result.Distinct().ToList();
        }

        private void CollectLogicalErrorSegments(string expression, List<string> collector)
        {
            string current = TrimOuterParentheses(expression.Trim());
            if (string.IsNullOrEmpty(current))
            {
                return;
            }

            bool evaluateOk;
            bool evaluateValue;
            if (!TryEvaluateBooleanExpression(current, out evaluateOk, out evaluateValue))
            {
                collector.Add(current);
                return;
            }

            if (evaluateValue)
            {
                return;
            }

            List<string> andParts = SplitTopLevelExpression(current, "&&");
            if (andParts.Count > 1)
            {
                foreach (string part in andParts)
                {
                    bool partOk;
                    bool partValue;
                    if (!TryEvaluateBooleanExpression(part, out partOk, out partValue) || !partValue)
                    {
                        CollectLogicalErrorSegments(part, collector);
                    }
                }
                return;
            }

            List<string> orParts = SplitTopLevelExpression(current, "||");
            if (orParts.Count > 1)
            {
                foreach (string part in orParts)
                {
                    bool partOk;
                    bool partValue;
                    if (!TryEvaluateBooleanExpression(part, out partOk, out partValue) || !partValue)
                    {
                        CollectLogicalErrorSegments(part, collector);
                    }
                }
                return;
            }

            collector.Add(current);
        }

        private static string TrimOuterParentheses(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return text;
            }

            string current = text.Trim();
            bool changed = true;
            while (changed)
            {
                changed = false;
                if (current.Length >= 2 && current[0] == '(' && current[current.Length - 1] == ')')
                {
                    int depth = 0;
                    bool isWrapped = true;
                    for (int i = 0; i < current.Length; i++)
                    {
                        if (current[i] == '(')
                            depth++;
                        else if (current[i] == ')')
                            depth--;

                        if (depth == 0 && i < current.Length - 1)
                        {
                            isWrapped = false;
                            break;
                        }
                    }

                    if (isWrapped)
                    {
                        current = current.Substring(1, current.Length - 2).Trim();
                        changed = true;
                    }
                }
            }

            return current;
        }

        private static List<string> SplitTopLevelExpression(string expression, string delimiter)
        {
            List<string> parts = new List<string>();
            if (string.IsNullOrEmpty(expression))
            {
                return parts;
            }

            int depth = 0;
            int start = 0;
            for (int i = 0; i < expression.Length; i++)
            {
                char ch = expression[i];
                if (ch == '(')
                    depth++;
                else if (ch == ')')
                    depth--;

                if (depth == 0 && i + delimiter.Length <= expression.Length
                    && expression.Substring(i, delimiter.Length) == delimiter)
                {
                    parts.Add(expression.Substring(start, i - start).Trim());
                    i += delimiter.Length - 1;
                    start = i + 1;
                }
            }

            parts.Add(expression.Substring(start).Trim());
            return parts.Where(p => !string.IsNullOrEmpty(p)).ToList();
        }

        private static bool TryEvaluateBooleanExpression(string expression, out bool evaluateOk, out bool value)
        {
            evaluateOk = false;
            value = false;
            try
            {
                NCalc.Expression expr = new NCalc.Expression(expression);
                object raw = expr.Evaluate();
                if (raw is bool)
                {
                    value = (bool)raw;
                    evaluateOk = true;
                    return true;
                }

                if (raw is int)
                {
                    value = (int)raw != 0;
                    evaluateOk = true;
                    return true;
                }

                bool boolParsed;
                if (bool.TryParse(Convert.ToString(raw), out boolParsed))
                {
                    value = boolParsed;
                    evaluateOk = true;
                    return true;
                }
            }
            catch
            {
            }

            return false;
        }

        private static void HighlightSegmentsInRichTextBox(System.Windows.Forms.RichTextBox richTextBox, string fullText, List<string> segments)
        {
            if (richTextBox == null)
            {
                return;
            }

            richTextBox.SelectAll();
            richTextBox.SelectionColor = Color.Black;

            if (string.IsNullOrEmpty(fullText) || segments == null || segments.Count == 0)
            {
                return;
            }

            foreach (string segment in segments)
            {
                if (string.IsNullOrWhiteSpace(segment))
                {
                    continue;
                }

                int startIndex = 0;
                while (startIndex < fullText.Length)
                {
                    int found = fullText.IndexOf(segment, startIndex, StringComparison.Ordinal);
                    if (found < 0)
                    {
                        break;
                    }

                    richTextBox.SelectionStart = found;
                    richTextBox.SelectionLength = segment.Length;
                    richTextBox.SelectionColor = Color.Red;
                    startIndex = found + segment.Length;
                }
            }
        }

        private void ApplyFilter(List<string> used_Feature)
        {
            List<int> indexs = new List<int>();
            for (int i = 0; i < DGV_MAIN.Rows.Count; i++)
            {
                string feature_name = DGV_MAIN.Rows[i].Cells[1].Value?.ToString();
                foreach (string feature in used_Feature)
                {
                    if (feature == feature_name)
                    {
                        indexs.Add(i);
                    }
                }
            }
            //-------------
            // Create a list to store the rows to keep
            List<DataGridViewRow> rowsToKeep = new List<DataGridViewRow>();

            // Iterate through the rows and add the ones to keep to the list
            foreach (DataGridViewRow row in DGV_MAIN.Rows)
            {
                if (indexs.Contains(row.Index))
                {
                    rowsToKeep.Add(row);
                }
            }
            // Clear the current rows
            DGV_MAIN.Rows.Clear();
            // Add the rows to keep back to the DataGridView
            foreach (DataGridViewRow row in rowsToKeep)
            {
                DGV_MAIN.Rows.Add(row.Cells[0].Value, row.Cells[1].Value, row.Cells[2].Value);
            }
        }





        private Task LV_DEFECT_LIST_SelectedIndexChangedAsync(object sender, EventArgs e)
        {
            if (isDefectSelectionUpdating)
            {
                return Task.CompletedTask;
            }

            isDefectSelectionUpdating = true;
            DGV_MAIN.SuspendLayout(); // 레이아웃 계산 중지
            try
            {
                //항목이 바뀔때 선택된 인덱스를 먼저 찾고 해당 이미지 경로에 접근해서 이미지가 있는경우에 위에 보여준다~ 
                if (LV_DEFECT_LIST.SelectedIndices.Count <= 0)
                {
                    select_nymber_defect = 9999;
                    lastDefectListSelectedIndex = -1;
                    PB_DEFECTMAP.Invalidate();
                    return Task.CompletedTask;
                }

                int selectedIndex = LV_DEFECT_LIST.SelectedIndices[0];
                System.Windows.Forms.ListViewItem selectedItem = LV_DEFECT_LIST.Items[selectedIndex];
                int columnValue = GetDefectSourceIndexFromListItem(selectedItem);
                if (columnValue < 0)
                {
                    return Task.CompletedTask;
                }

                if (columnValue == lastDefectListSelectedIndex)
                {
                    return Task.CompletedTask;
                }
                lastDefectListSelectedIndex = columnValue;

                navi.Defect_index = columnValue;
                try // Defect Featrue 출력 
                {

                    List<object[]> data = new List<object[]>();
                    int index_test = 0;

                    if (View_mode == 1)//pre
                    {
                        foreach (var kvp in Feature_row[columnValue])
                        {
                            data.Add(new object[] { index_test++, kvp.Key, kvp.Value });
                        }

                        UpdateDefectGridRows(data);

                        int Ptn_num = 0;
                        int defectIdx = columnValue;

                        int sum = 0;
                        int VP_select_num = 0;
                        for (int vpnum = 0; vpnum < VP_Defect_num.Count; vpnum++)
                        {

                            if (defectIdx < (sum + VP_Defect_num[vpnum]))
                            {
                                VP_select_num = vpnum;
                                defectIdx -= sum;
                                break;
                            }
                            sum += VP_Defect_num[vpnum];
                        }

                        LoadFileNamesFromBinary(Crop_bin_path_Pre[VP_select_num], Ptn_num, defectIdx);
                    }
                    else
                    {
                        foreach (var kvp in Feature_row_post[columnValue])
                        {
                            data.Add(new object[] { index_test++, kvp.Key, kvp.Value });
                        }

                        UpdateDefectGridRows(data);

                        int Ptn_num = 0;
                        int defectIdx = columnValue;


                        int sum = 0;
                        int VP_select_num = 0;
                        for (int vpnum = 0; vpnum < VP_Defect_num.Count; vpnum++)
                        {

                            if (defectIdx <= (sum + VP_Defect_num[vpnum]))
                            {
                                VP_select_num = vpnum;
                                defectIdx -= sum;
                                break;
                            }
                            sum += VP_Defect_num[vpnum];
                        }
                        LoadFileNamesFromBinary(Crop_bin_path_Post[VP_select_num], Ptn_num, defectIdx);
                    }

                    select_nymber_defect = columnValue;
                    PB_DEFECTMAP.Invalidate();
                    check_picturebox = true;

                    //DGV_MAIN 선택된 행이있으면 해당 행으로 옮겨가자 
                    if (DGV_MAIN.Rows.Count > 0)
                    {
                        if (DGV_MAIN_INDEX < 0 || DGV_MAIN_INDEX >= DGV_MAIN.Rows.Count)
                        {
                            DGV_MAIN_INDEX = 0;
                        }
                        DGV_MAIN.Rows[DGV_MAIN_INDEX].Selected = true;
                        DGV_MAIN.CurrentCell = DGV_MAIN.Rows[DGV_MAIN_INDEX].Cells[0];
                    }

                }
                catch
                {
                    //feature img_path 경로가 없을때  
                }
            }
            finally
            {
                //해당 Feature 정보로 classify 한번 돌리기 
                DGV_MAIN.ResumeLayout(); // 레이아웃 계산 재개
                isDefectSelectionUpdating = false;
            }

            return Task.CompletedTask;
        }

        private void UpdateDefectGridRows(List<object[]> rows)
        {
            DGV_MAIN.Rows.Clear();
            if (rows == null || rows.Count == 0)
            {
                return;
            }

            DataGridViewRow[] gridRows = new DataGridViewRow[rows.Count];
            for (int i = 0; i < rows.Count; i++)
            {
                DataGridViewRow row = new DataGridViewRow();
                row.CreateCells(DGV_MAIN, rows[i]);
                gridRows[i] = row;
            }
            DGV_MAIN.Rows.AddRange(gridRows);
        }

        public string[] GetSearchFile_all(String _strPath)
        {
            string[] files = { "", };
            try
            {
                // TopDirectoryOnly
                files = Directory.GetFiles(_strPath, "*.*", SearchOption.AllDirectories);
            }
            catch (IOException ex)
            {
                MessageBox.Show(ex.Message);
            }
            return files;
        }                     //폴더내 파일검색 _하위폴더 포함

        private void BTN_RECIPE_Click(object sender, EventArgs e)
        {
            Process.Start("explorer.exe", vp_info_list[0].Path_Vp_Recipe);

        }

        private void PictureBox_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (check_picturebox)
            {
                // PictureBox 내 마우스 위치 가져오기
                int mouseX = e.X;
                int mouseY = e.Y;

                // 섹션 크기 계산
                int sectionWidth = originalImage.Width / divisionCount_col;
                int sectionHeight = originalImage.Height / divisionCount_row;

                // 마우스 위치에 따른 섹션 계산
                int sectionX = mouseX / (PB_DEFECT_ARRAY.Width / divisionCount_col);
                int sectionY = mouseY / (PB_DEFECT_ARRAY.Height / divisionCount_row);

                // 강조 표시할 영역 설정
                highlightRect = new Rectangle(
                    sectionX * (PB_DEFECT_ARRAY.Width / divisionCount_col),
                    sectionY * (PB_DEFECT_ARRAY.Height / divisionCount_row),
                    PB_DEFECT_ARRAY.Width / divisionCount_col,
                    PB_DEFECT_ARRAY.Height / divisionCount_row
                );

                // PictureBox 다시 그리기
                PB_DEFECT_ARRAY.Invalidate();
            }
        }


        private void PictureBox_MouseClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (check_picturebox)
            {
                // 섹션 크기 계산
                int sectionWidth = originalImage.Width / divisionCount_col;
                int sectionHeight = originalImage.Height / divisionCount_row;

                // 클릭된 섹션 계산
                int sectionX = e.X / (PB_DEFECT_ARRAY.Width / divisionCount_col);
                int sectionY = e.Y / (PB_DEFECT_ARRAY.Height / divisionCount_row);

                // 크롭 영역 설정
                Rectangle cropRect = new Rectangle(
                    sectionX * sectionWidth,
                    sectionY * sectionHeight,
                    sectionWidth,
                    sectionHeight
                );
                // 크롭된 이미지 생성
                Bitmap croppedImage = new Bitmap(cropRect.Width, cropRect.Height);
                using (Graphics g = Graphics.FromImage(croppedImage))
                {
                    g.DrawImage(originalImage, new Rectangle(0, 0, croppedImage.Width, croppedImage.Height), cropRect, GraphicsUnit.Pixel);
                }

                SetPictureBoxImage(PB_MAIN, croppedImage);
            }
        }
        private void PictureBox_Paint(object sender, PaintEventArgs e)
        {
            if (check_picturebox)
            {
                // 강조 표시할 영역에 빨간 테두리 그리기
                if (highlightRect != Rectangle.Empty)
                {
                    using (Pen pen = new Pen(Color.Red, 2))
                    {
                        e.Graphics.DrawRectangle(pen, highlightRect);
                    }
                }
            }
        }


        private void BTN_DEFECT_PANEL_Click(object sender, EventArgs e)
        {
            Clipboard.SetText(inspectionDataList[insp_info.listview_index].SerialNumber);
        }

        private void BTN_PANEL_JUDGE_Click(object sender, EventArgs e)
        {
            ////후보점 정보에 대하여 판정을 모든 판정을 해야함
            //bool Defect_Judge = true;
            ////classify의 모든 조건문 
            //for (int i = 0; i < classify_Infos.Count; i++)
            //{
            //    List<string> NG_str = new List<string>();
            //    if (LV_DEFECT_LIST.SelectedIndices.Count > 0)
            //    {
            //        string[] Origin_word = classify_Infos[i].SCRIPT.Split(' ');
            //        int selectedIndex;
            //        // 선택된 인덱스를 사용하여 작업 수행
            //        selectedIndex = LV_DEFECT_LIST.SelectedIndices[0];
            //        int index = 0;
            //        string[] words = classify_Infos[i].SCRIPT.Split(' ');
            //        foreach (var kvp in Feature_row[LV_DEFECT_LIST.SelectedIndices[0]])
            //        {
            //            //select_classify.SCRIPT=select_classify.SCRIPT.Replace(kvp.Key.ToString(), kvp.Value.ToString());
            //            if (words.Contains(kvp.Key))
            //            {
            //                bool push = true;
            //                for (int key_i = 0; key_i < words.Length; key_i++)
            //                {
            //                    if (words[key_i] == kvp.Key)
            //                    {
            //                        if (push)
            //                        {
            //                            used_Feature.Add(kvp.Key.ToString());
            //                            push = false;
            //                        }
            //                        words[key_i] = kvp.Value.ToString();
            //                    }
            //                }
            //            }
            //        }
            //        for (int words_i = 0; words_i < words.Length; words_i++)
            //        {
            //            if (words[words_i] == "and")
            //            {
            //                words[words_i] = "&&";
            //            }
            //            if (words[words_i] == "or")
            //            {
            //                words[words_i] = "||";
            //            }
            //        }
            //        string reversedString = string.Join(" ", words);
            //        CMergeClassify classifier = new CMergeClassify();
            //        int result = classifier.ComputeExpression(reversedString);
            //        if (result == 1)
            //        {
            //            Defect_Judge = false;
            //            LB_PANEL_JUDGE.Text = classify_Infos[i].SCRIPT_NAME;
            //            break;
            //        }

            //        Console.WriteLine($"{classify_Infos[i].SCRIPT_NAME}: {result}"); // 출력: Result: 1
            //    }
            //}

        }

        private void metroButton1_Click(object sender, EventArgs e)
        {

            try
            {
                if (Convert.ToInt32(TB_COL_COUNT.Text) > 0 && Convert.ToInt32(TB_ROW_COUNT.Text) > 0)
                {
                    divisionCount_col = Convert.ToInt32(TB_COL_COUNT.Text);
                    divisionCount_row = Convert.ToInt32(TB_ROW_COUNT.Text);
                }
                PB_DEFECT_ARRAY.Update();
            }
            catch { }


        }

        private void BTN_JUDGE_Click(object sender, EventArgs e) // 클릭시 LISTVIEW가  필터링 되도록 구성.
        {

            string Judge_name = "0";
            switch (Judge_mode)
            {
                case 0:  //all 보여주기 > OK 만보여주기

                    Judge_name = "OK";
                    break;
                case 1:   // OK -> NG 만 보여주기
                    Judge_name = "NG";
                    break;

                case 2:   //NG -> ALL 다보여주기 
                    break;
            }
            LV_DEFECT_LIST.BeginUpdate();
            try
            {
                LV_DEFECT_LIST.Items.Clear();

                for (int i = 0; i < Feature_row.Count; i++)
                {
                    if (Feature_row[i]["DefectJudge"]?.ToString() != Judge_name && (Judge_mode == 0 || Judge_mode == 1))
                        continue;
                    var listViewItem = new System.Windows.Forms.ListViewItem((i + 1).ToString());

                    for (int feature_num = 0; feature_num < View_feature_num; feature_num++)
                    {
                        listViewItem.SubItems.Add(Feature_row[i][VIEW_DEFECT_FEATURE_NAME[feature_num]].ToString());
                    }

                    LV_DEFECT_LIST.Items.Add(listViewItem);
                }
            }
            finally
            {
                LV_DEFECT_LIST.EndUpdate();
            }

            switch (Judge_mode)
            {
                case 0:  //all 보여주기 > OK 만보여주기
                    BTN_JUDGE.Text = "Judge (OK)";
                    Judge_mode = 1;
                    break;
                case 1:   // OK -> NG 만 보여주기
                    BTN_JUDGE.Text = "Judge (NG)";
                    Judge_mode = 2;
                    break;
                case 2:   //NG -> ALL 다보여주기 
                    BTN_JUDGE.Text = "Judge (ALL)";
                    Judge_mode = 0;
                    break;
            }

        }

        private void BTN_SIMULATION_RUN_Click(object sender, EventArgs e)
        {
            //검사 진행할 RECIPE 및 CLASSFFIY 주소 설정 이거야뭐 VP0 기준이라고 생각하자 
            string SimulRecipePath = vp_info_list[0].Path_Vp_Recipe + "\\" + @CBB_PRODUCT_RECIPE_FOLDER.Text + @"\" + CBB_PRODUCT_RECIPE_FILE.Text;
            string SimulClassfiyPath = vp_info_list[0].Path_Vp_Recipe + "\\" + @CBB_PRODUCT_CLASSIFY_FOLDER.Text + @"\" + CBB_PRODUCT_CLASSIFY_FILE.Text;
            cSimulationrun.Recile_Path = SimulRecipePath;
            cSimulationrun.Classify_Path = SimulClassfiyPath;

            if (LV_PANEL_LIST.SelectedItems.Count == 0)
            {
                MessageBox.Show("시뮬레이션 대상을 먼저 선택해주세요.");
                return;
            }

            List<string> PanelSimulationFilePath = new List<string>();
            List<int> targetListViewIndices = new List<int>();
            PanelSimulationFilePath.Clear();

            foreach (System.Windows.Forms.ListViewItem selected in LV_PANEL_LIST.SelectedItems)
            {
                if (selected.SubItems.Count > 8)
                {
                    selected.SubItems[7].Text = string.Empty;
                    selected.SubItems[8].Text = string.Empty;
                }
            }

            //먼저 조회된 PANEL 기준으로만 재검사를할지말지 
            // 조회된 애들만 검사 
            //조회된 대상들은 검사 이력에대한 Panel.Simulation.json 파일위치를 얻어내야함 
            string recipeFolderName = CBB_PRODUCT_RECIPE_FOLDER.Text;
            string recipeFolderSuffix = recipeFolderName.Length >= 5 ? recipeFolderName.Substring(recipeFolderName.Length - 5) : recipeFolderName;

            foreach (System.Windows.Forms.ListViewItem selectedItem in LV_PANEL_LIST.SelectedItems)
            {
                if (selectedItem.SubItems.Count <= 6)
                {
                    continue;
                }

                string serialNumber = selectedItem.SubItems[1].Text;
                string vpNumberText = selectedItem.SubItems[6].Text;
                if (string.IsNullOrEmpty(serialNumber) || string.IsNullOrEmpty(vpNumberText))
                {
                    continue;
                }

                string selectedVpSuffix = vpNumberText.Substring(vpNumberText.Length - 1);
                int sourceIndex = -1;
                for (int i = 0; i < inspectionDataList.Count; i++)
                {
                    string vpnum = inspectionDataList[i].Vpnum;
                    if (string.IsNullOrEmpty(vpnum))
                    {
                        continue;
                    }

                    string vpSuffix = vpnum.Substring(vpnum.Length - 1);
                    if (inspectionDataList[i].SerialNumber == serialNumber && vpSuffix == selectedVpSuffix)
                    {
                        sourceIndex = i;
                        break;
                    }
                }

                if (sourceIndex < 0)
                {
                    continue;
                }

                int vp_num;
                if (!int.TryParse(inspectionDataList[sourceIndex].Vpnum.Substring(inspectionDataList[sourceIndex].Vpnum.Length - 1), out vp_num))
                {
                    continue;
                }

                if (vp_num <= 0 || vp_num > vp_info_list.Count)
                {
                    continue;
                }

                string simulationPath = vp_info_list[vp_num - 1].Path_Vp_Inspection + "\\" + inspectionDataList[sourceIndex].InspectionDate + "\\" + recipeFolderSuffix + "\\" + inspectionDataList[sourceIndex].SerialNumber + "_VP0" + vp_num;
                if (!Directory.Exists(simulationPath))
                {
                    continue;
                }

                PanelSimulationFilePath.Add(simulationPath);
                targetListViewIndices.Add(selectedItem.Index);
            }

            if (PanelSimulationFilePath.Count == 0)
            {
                return;
            }

            //정상동작 완료 
            cSimulationrun.PanelSimulationRun(SimulClassfiyPath, PanelSimulationFilePath);

            //아래 결과쓰기 
            for (int i = 0; i < cSimulationrun.pnael_Result_Infos.Count; i++)
            {
                int resultIndex = cSimulationrun.pnael_Result_Infos[i].index;
                if (resultIndex < 0 || resultIndex >= targetListViewIndices.Count)
                {
                    continue;
                }

                int targetListIndex = targetListViewIndices[resultIndex];
                if (targetListIndex < 0 || targetListIndex >= LV_PANEL_LIST.Items.Count)
                {
                    continue;
                }

                System.Windows.Forms.ListViewItem targetItem = LV_PANEL_LIST.Items[targetListIndex];
                if (targetItem.SubItems.Count <= 8)
                {
                    continue;
                }

                if (targetItem.SubItems[5].Text == cSimulationrun.pnael_Result_Infos[i].Classify_group)
                    targetItem.SubItems[7].Text = "O";
                else
                    targetItem.SubItems[7].Text = "X";

                targetItem.SubItems[8].Text = cSimulationrun.pnael_Result_Infos[i].Classify_group;
            }
        }

        private void CBB_SINGLE_RECIPE_COPY_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                string[] change_word = { };
                string classify_name = CBB_SINGLE_RECIPE_COPY.Text.ToString();
                Classify_Pre_info select_classify_pre;
                select_classify_pre.SCRIPT = "";

                Classify_Post_info select_classify_post;
                select_classify_post.SCRIPT = "";


                //현재 쓰고있는 FEATURE만 보여주기위한 List변수
                used_Feature_Pre.Clear();
                used_Feature_Post.Clear();

                //Post 사용여부 
                bool use_Post = false;

                for (int i = 0; i < classify_Pre_Infos.Count; i++)
                {
                    if (classify_Pre_Infos[i].SCRIPT_NAME == classify_name)
                    {
                        select_classify_pre = classify_Pre_Infos[i];

                        for (int j = 0; j < classify_Post_Infos.Count; j++)
                        {
                            if (classify_Post_Infos[j].SCRIPT_NAME == classify_name)
                                if (classify_Post_Infos[j].BYPASS == "True")
                                {
                                    select_classify_post = classify_Post_Infos[i];
                                    use_Post = true;
                                }
                        }
                    }
                }
                List<string> NG_str_PRE = new List<string>();
                if (LV_DEFECT_LIST.SelectedIndices.Count > 0)
                {
                    string[] Origin_word = select_classify_pre.SCRIPT.Split(' ');
                    int selectedIndex;
                    // 선택된 인덱스를 사용하여 작업 수행
                    selectedIndex = LV_DEFECT_LIST.SelectedIndices[0];

                    System.Windows.Forms.ListViewItem selectedItem = LV_DEFECT_LIST.Items[selectedIndex];
                    // 특정 열의 값 가져오기 (예: 첫 번째 열 = 인덱스 0)
                    int columnValue = Convert.ToInt32(selectedItem.SubItems[0].Text) - 1;
                    selectedIndex = columnValue;
                    try // Defect Featrue 출력 
                    {
                        int index = 0;
                        string[] words = select_classify_pre.SCRIPT.Split(' ');
                        foreach (var kvp in Feature_row[selectedIndex])
                        {
                            //select_classify.SCRIPT=select_classify.SCRIPT.Replace(kvp.Key.ToString(), kvp.Value.ToString());
                            if (words.Contains(kvp.Key))
                            {
                                bool push = true;
                                for (int i = 0; i < words.Length; i++)
                                {
                                    if (words[i] == kvp.Key)
                                    {
                                        if (push)
                                        {
                                            used_Feature_Pre.Add(kvp.Key.ToString());
                                            push = false;
                                        }
                                        words[i] = kvp.Value.ToString();
                                    }
                                }
                            }
                        }
                        for (int i = 0; i < words.Length; i++)
                        {
                            if (words[i] == "and")
                            {
                                words[i] = "&&";
                            }
                            if (words[i] == "or")
                            {
                                words[i] = "||";
                            }
                        }
                        change_word = words;
                        //Array.Reverse(words);
                        string reversedString = string.Join(" ", words);
                        CMergeClassify classifier = new CMergeClassify();
                        int result = classifier.ComputeExpression(reversedString);
                        Console.WriteLine($"Result: {result}"); // 출력: Result: 1
                        FindIncorrectPart(reversedString, NG_str_PRE);
                        LB_DEFECT_JUDGE.UseStyleColors = false; // 사용자 정의 색상을 사용하

                        if (result == 1)
                        {
                            LB_DEFECT_JUDGE.UseStyleColors = true; // 사용자 정의 색상을 사용하
                                                                   //LB_DEFECT_JUDGE.ForeColor = System.Drawing.Color.Green; // 예: 파란색
                            LB_DEFECT_JUDGE.Text = "Classify 조건식 : 참";
                        }
                        if (result == 0)
                        {
                            LB_DEFECT_JUDGE.UseStyleColors = true; // 사용자 정의 색상을 사용하
                                                                   //LB_DEFECT_JUDGE.ForeColor = System.Drawing.Color.Red; // 예: 파란색
                            LB_DEFECT_JUDGE.Text = "Classify 조건식 : 불";
                        }
                        if (result == 2)
                        {
                            LB_DEFECT_JUDGE.ForeColor = System.Drawing.Color.Blue; // 예: 파란색
                                                                                   //LB_DEFECT_JUDGE.UseStyleColors = true; // 사용자 정의 색상을 사용하
                            LB_DEFECT_JUDGE.Text = "Classify 조건식 : Error";
                        }

                        RTB_ORIGIN_PRE.Text = string.Join(" ", Origin_word);
                        RTB_REPLACE_PRE.Text = string.Join(" ", change_word);
                    }
                    catch
                    {
                        Console.WriteLine("Script Judge Error");
                    }
                    ApplyFilter(used_Feature_Pre);

                    //빨간색 변환 
                    for (int i = 0; i < NG_str_PRE.Count; i++)
                    {
                        int start_index = 0, end_index = 0;
                        start_index = string.Join(" ", change_word).IndexOf(NG_str_PRE[i]);
                        end_index = NG_str_PRE[i].Length;

                        RTB_REPLACE_PRE.SelectionStart = start_index; // Start at the beginning
                        RTB_REPLACE_PRE.SelectionLength = end_index; // Select the first 5 characters
                        RTB_REPLACE_PRE.SelectionColor = Color.Red; // Set the color to red
                                                                    //origin 변경을 위해서는  앞의 인덱스를 알아야함.
                        string start_str = string.Join(" ", change_word).Substring(0, start_index);
                        string end_str = string.Join(" ", change_word).Substring(0, start_index + end_index);
                        start_str = start_str.Replace("&&", "&");
                        start_str = start_str.Replace("||", "|");
                        string[] start_parts_info = start_str.Split(new char[] { '(', ')', '&', '|' }, StringSplitOptions.RemoveEmptyEntries);
                        string[] start_parts = RemoveEmptyStrings(start_parts_info);
                        start_parts.Count();
                        string[] Origin_word_TEMP = Origin_word.ToArray();
                        for (int origin_num = 0; origin_num < Origin_word_TEMP.Length; origin_num++)
                        {
                            if (Origin_word_TEMP[origin_num] == "and")
                            {
                                Origin_word_TEMP[origin_num] = "&&";
                            }
                            if (Origin_word_TEMP[origin_num] == "or")
                            {
                                Origin_word_TEMP[origin_num] = "||";
                            }
                        }
                        string Origin_word_TEMP_str = string.Join(" ", Origin_word_TEMP);

                        Origin_word_TEMP_str = Origin_word_TEMP_str.Replace("&&", "&");
                        Origin_word_TEMP_str = Origin_word_TEMP_str.Replace("||", "|");
                        string[] Origin_start_parts_info = Origin_word_TEMP_str.Split(new char[] { '(', ')', '&', '|' }, StringSplitOptions.RemoveEmptyEntries);
                        string[] Origin_start_parts = RemoveEmptyStrings(Origin_start_parts_info);

                        string find_str = Origin_start_parts[start_parts.Count()];
                        int replace_start_index = select_classify_pre.SCRIPT.IndexOf(find_str);
                        int replace_end_index = find_str.Length;

                        RTB_ORIGIN_PRE.SelectionStart = replace_start_index; // Start at the beginning
                        RTB_ORIGIN_PRE.SelectionLength = replace_end_index; // Select the first 5 characters
                        RTB_ORIGIN_PRE.SelectionColor = Color.Red; // Set the color to red
                    }
                }

                //NG 항목대하여 색처리 하기위해서?
                //기존 REPLACE 대상에 대하여 NG_STR 찾기




                if (use_Post) //POST 검증
                {
                    BTN_CHANGE.ForeColor = Color.Yellow;

                    //Post 경로 추가 
                    List<string> NG_str_Post = new List<string>();
                    if (LV_DEFECT_LIST.SelectedIndices.Count > 0)
                    {
                        string[] Origin_word = select_classify_post.SCRIPT.Split(' ');
                        int selectedIndex;
                        // 선택된 인덱스를 사용하여 작업 수행
                        selectedIndex = LV_DEFECT_LIST.SelectedIndices[0];

                        System.Windows.Forms.ListViewItem selectedItem = LV_DEFECT_LIST.Items[selectedIndex];
                        // 특정 열의 값 가져오기 (예: 첫 번째 열 = 인덱스 0)
                        int columnValue = Convert.ToInt32(selectedItem.SubItems[0].Text) - 1;
                        selectedIndex = columnValue;
                        try // Defect Featrue 출력 
                        {
                            int index = 0;
                            string[] words = select_classify_post.SCRIPT.Split(' ');
                            foreach (var kvp in Feature_row_post[selectedIndex])
                            {
                                //select_classify.SCRIPT=select_classify.SCRIPT.Replace(kvp.Key.ToString(), kvp.Value.ToString());
                                if (words.Contains(kvp.Key))
                                {
                                    bool push = true;
                                    for (int i = 0; i < words.Length; i++)
                                    {
                                        if (words[i] == kvp.Key)
                                        {
                                            if (push)
                                            {
                                                used_Feature_Post.Add(kvp.Key.ToString());
                                                push = false;
                                            }
                                            words[i] = kvp.Value.ToString();
                                        }
                                    }
                                }
                            }
                            for (int i = 0; i < words.Length; i++)
                            {
                                if (words[i] == "and")
                                {
                                    words[i] = "&&";
                                }
                                if (words[i] == "or")
                                {
                                    words[i] = "||";
                                }
                            }
                            change_word = words;
                            //Array.Reverse(words);
                            string reversedString = string.Join(" ", words);
                            CMergeClassify classifier = new CMergeClassify();
                            int result = classifier.ComputeExpression(reversedString);
                            Console.WriteLine($"Result: {result}"); // 출력: Result: 1
                            FindIncorrectPart(reversedString, NG_str_Post);
                            LB_DEFECT_JUDGE.UseStyleColors = false; // 사용자 정의 색상을 사용하

                            if (result == 1)
                            {
                                LB_DEFECT_JUDGE.UseStyleColors = true; // 사용자 정의 색상을 사용하
                                                                       //LB_DEFECT_JUDGE.ForeColor = System.Drawing.Color.Green; // 예: 파란색
                                LB_DEFECT_JUDGE.Text = "Classify 조건식 : 참";
                            }
                            if (result == 0)
                            {
                                LB_DEFECT_JUDGE.UseStyleColors = true; // 사용자 정의 색상을 사용하
                                                                       //LB_DEFECT_JUDGE.ForeColor = System.Drawing.Color.Red; // 예: 파란색
                                LB_DEFECT_JUDGE.Text = "Classify 조건식 : 불";
                            }
                            if (result == 2)
                            {
                                LB_DEFECT_JUDGE.ForeColor = System.Drawing.Color.Blue; // 예: 파란색
                                                                                       //LB_DEFECT_JUDGE.UseStyleColors = true; // 사용자 정의 색상을 사용하
                                LB_DEFECT_JUDGE.Text = "Classify 조건식 : Error";
                            }

                            RTB_ORIGIN_POST.Text = string.Join(" ", Origin_word);
                            RTB_REPLACE_POST.Text = string.Join(" ", change_word);
                        }
                        catch
                        {
                            Console.WriteLine("Script Judge Error");
                        }
                        ApplyFilter(used_Feature_Post);

                        //빨간색 변환 
                        for (int i = 0; i < NG_str_Post.Count; i++)
                        {
                            int start_index = 0, end_index = 0;
                            start_index = string.Join(" ", change_word).IndexOf(NG_str_Post[i]);
                            end_index = NG_str_Post[i].Length;

                            RTB_REPLACE_POST.SelectionStart = start_index; // Start at the beginning
                            RTB_REPLACE_POST.SelectionLength = end_index; // Select the first 5 characters
                            RTB_REPLACE_POST.SelectionColor = Color.Red; // Set the color to red
                                                                         //origin 변경을 위해서는  앞의 인덱스를 알아야함.
                            string start_str = string.Join(" ", change_word).Substring(0, start_index);
                            string end_str = string.Join(" ", change_word).Substring(0, start_index + end_index);
                            start_str = start_str.Replace("&&", "&");
                            start_str = start_str.Replace("||", "|");
                            string[] start_parts_info = start_str.Split(new char[] { '(', ')', '&', '|' }, StringSplitOptions.RemoveEmptyEntries);
                            string[] start_parts = RemoveEmptyStrings(start_parts_info);
                            start_parts.Count();
                            string[] Origin_word_TEMP = Origin_word.ToArray();
                            for (int origin_num = 0; origin_num < Origin_word_TEMP.Length; origin_num++)
                            {
                                if (Origin_word_TEMP[origin_num] == "and")
                                {
                                    Origin_word_TEMP[origin_num] = "&&";
                                }
                                if (Origin_word_TEMP[origin_num] == "or")
                                {
                                    Origin_word_TEMP[origin_num] = "||";
                                }
                            }
                            string Origin_word_TEMP_str = string.Join(" ", Origin_word_TEMP);

                            Origin_word_TEMP_str = Origin_word_TEMP_str.Replace("&&", "&");
                            Origin_word_TEMP_str = Origin_word_TEMP_str.Replace("||", "|");
                            string[] Origin_start_parts_info = Origin_word_TEMP_str.Split(new char[] { '(', ')', '&', '|' }, StringSplitOptions.RemoveEmptyEntries);
                            string[] Origin_start_parts = RemoveEmptyStrings(Origin_start_parts_info);

                            string find_str = Origin_start_parts[start_parts.Count()];
                            int replace_start_index = select_classify_post.SCRIPT.IndexOf(find_str);
                            int replace_end_index = find_str.Length;

                            RTB_ORIGIN_POST.SelectionStart = replace_start_index; // Start at the beginning
                            RTB_ORIGIN_POST.SelectionLength = replace_end_index; // Select the first 5 characters
                            RTB_ORIGIN_POST.SelectionColor = Color.Red; // Set the color to red
                        }
                    }
                }
            }
            catch
            { MessageBox.Show("올바르지 않은 접근입니다."); }
        }

        private void RTB_ORIGIN_TextChanged(object sender, EventArgs e)
        {
            BTN_RULE_INIT.Text = "INIT";
            Validate();
        }

        private void BTN_RULE_INIT_Click(object sender, EventArgs e)
        {
            if (BTN_RULE_INIT.Text == "INIT")
                BTN_RULE_INIT.Text = "-";

        }

        private void BTN_RULE_RUN_Click(object sender, EventArgs e)
        {

            //string[] change_word = { };
            //List<string> NG_str = new List<string>();


            //string classify_name = CBB_SINGLE_RECIPE_COPY.Text.ToString();
            //string[] Origin_word = RTB_ORIGIN_PRE.Text.Split(' ');

            //int result = -1;
            //Classify_info select_classify;
            //select_classify.SCRIPT = RTB_ORIGIN_PRE.Text;
            //try // Defect Featrue 출력 
            //{


            //    int index = 0;
            //    string[] words = select_classify.SCRIPT.Split(' ');
            //    foreach (var kvp in Feature_row[LV_DEFECT_LIST.SelectedIndices[0]])
            //    {
            //        //select_classify.SCRIPT=select_classify.SCRIPT.Replace(kvp.Key.ToString(), kvp.Value.ToString());
            //        if (words.Contains(kvp.Key))
            //        {
            //            bool push = true;
            //            for (int i = 0; i < words.Length; i++)
            //            {
            //                if (words[i] == kvp.Key)
            //                {
            //                    if (push)
            //                    {
            //                        used_Feature.Add(kvp.Key.ToString());
            //                        push = false;
            //                    }
            //                    words[i] = kvp.Value.ToString();
            //                }
            //            }
            //        }
            //    }

            //    for (int i = 0; i < words.Length; i++)
            //    {
            //        if (words[i] == "and")
            //        {
            //            words[i] = "&&";
            //        }
            //        if (words[i] == "or")
            //        {
            //            words[i] = "||";
            //        }
            //    }
            //    change_word = words;
            //    //Array.Reverse(words);
            //    string reversedString = string.Join(" ", words);
            //    CMergeClassify classifier = new CMergeClassify();
            //    result = classifier.ComputeExpression(reversedString);
            //    Console.WriteLine($"Result: {result}"); // 출력: Result: 1
            //    FindIncorrectPart(reversedString, NG_str);

            //    if (result == 1)
            //    {
            //        LB_DEFECT_JUDGE_CHECK.Text = "Classify 조건식 : 참";
            //    }
            //    if (result == 0)
            //    {
            //        LB_DEFECT_JUDGE_CHECK.Text = "Classify 조건식 : 불";
            //    }
            //    if (result == 2)
            //    {
            //        LB_DEFECT_JUDGE_CHECK.Text = "Classify 조건식 : Error";
            //    }

            //    RTB_ORIGIN_PRE.Text = string.Join(" ", Origin_word);
            //    RTB_REPLACE_PRE.Text = string.Join(" ", change_word);
            //}
            //catch
            //{
            //    Console.WriteLine("Script Judge Error");
            //}

            //ApplyFilter();


            //if (result == 0 || result == 2)
            //{

            //    //빨간색 변환 
            //    for (int i = 0; i < NG_str.Count; i++)
            //    {
            //        int start_index = 0, end_index = 0;
            //        start_index = string.Join(" ", change_word).IndexOf(NG_str[i]);
            //        end_index = NG_str[i].Length;

            //        RTB_REPLACE_PRE.SelectionStart = start_index; // Start at the beginning
            //        RTB_REPLACE_PRE.SelectionLength = end_index; // Select the first 5 characters
            //        RTB_REPLACE_PRE.SelectionColor = Color.Red; // Set the color to red

            //        //origin 변경을 위해서는  앞의 인덱스를 알아야함.
            //        string start_str = string.Join(" ", change_word).Substring(0, start_index);
            //        string end_str = string.Join(" ", change_word).Substring(0, start_index + end_index);

            //        start_str = start_str.Replace("&&", "&");
            //        start_str = start_str.Replace("||", "|");
            //        string[] start_parts_info = start_str.Split(new char[] { '(', ')', '&', '|' }, StringSplitOptions.RemoveEmptyEntries);
            //        string[] start_parts = RemoveEmptyStrings(start_parts_info);
            //        start_parts.Count();
            //        string[] Origin_word_TEMP = Origin_word.ToArray();

            //        for (int origin_num = 0; origin_num < Origin_word_TEMP.Length; origin_num++)
            //        {
            //            if (Origin_word_TEMP[origin_num] == "and")
            //                Origin_word_TEMP[origin_num] = "&&";
            //            if (Origin_word_TEMP[origin_num] == "or")
            //                Origin_word_TEMP[origin_num] = "||";
            //        }

            //        string Origin_word_TEMP_str = string.Join(" ", Origin_word_TEMP);

            //        Origin_word_TEMP_str = Origin_word_TEMP_str.Replace("&&", "&");
            //        Origin_word_TEMP_str = Origin_word_TEMP_str.Replace("||", "|");
            //        string[] Origin_start_parts_info = Origin_word_TEMP_str.Split(new char[] { '(', ')', '&', '|' }, StringSplitOptions.RemoveEmptyEntries);
            //        string[] Origin_start_parts = RemoveEmptyStrings(Origin_start_parts_info);

            //        string find_str = Origin_start_parts[start_parts.Count()];
            //        int replace_start_index = select_classify.SCRIPT.IndexOf(find_str);
            //        int replace_end_index = find_str.Length;

            //        RTB_ORIGIN_PRE.SelectionStart = replace_start_index; // Start at the beginning
            //        RTB_ORIGIN_PRE.SelectionLength = replace_end_index; // Select the first 5 characters
            //        RTB_ORIGIN_PRE.SelectionColor = Color.Red; // Set the color to red

            //    }
            //}
        }
        private void TC_LADYBUG_SelectedIndexChanged(object sender, EventArgs e)
        {

            CBB_PRODUCT_RECIPE_FOLDER.Items.Clear();
            CBB_PRODUCT_CLASSIFY_FOLDER.Items.Clear();
            //RECIPE 폴더COMBOBOX 추가
            string[] recipe = GetSearchFolder_only(vp_info_list[0].Path_Vp_Recipe);
            for (int i = 0; i < recipe.Length; i++)
            {
                CBB_PRODUCT_RECIPE_FOLDER.Items.Add(Path.GetFileName(recipe[i]));
                CBB_PRODUCT_CLASSIFY_FOLDER.Items.Add(Path.GetFileName(recipe[i]));
            }
        }

        private void CBB_PRODUCT_RECIPE_FOLDER_SelectedIndexChanged(object sender, EventArgs e)
        {
            CBB_PRODUCT_RECIPE_FILE.Items.Clear();
            //RECIPE 폴더COMBOBOX 추가
            string[] recipe_file = GetSearchFile_only_recipe(vp_info_list[0].Path_Vp_Recipe + "\\" + CBB_PRODUCT_RECIPE_FOLDER.SelectedItem.ToString());
            for (int i = 0; i < recipe_file.Length; i++)
            {
                CBB_PRODUCT_RECIPE_FILE.Items.Add(Path.GetFileName(recipe_file[i]));
            }
        }

        private void CBB_PRODUCT_CLASSIFY_FOLDER_SelectedIndexChanged(object sender, EventArgs e)
        {
            CBB_PRODUCT_CLASSIFY_FILE.Items.Clear();
            //RECIPE 폴더COMBOBOX 추가
            string[] classify_file = GetSearchFile_only_classify(vp_info_list[0].Path_Vp_Recipe + "\\" + CBB_PRODUCT_CLASSIFY_FOLDER.Text);
            for (int i = 0; i < classify_file.Length; i++)
            {
                CBB_PRODUCT_CLASSIFY_FILE.Items.Add(Path.GetFileName(classify_file[i]));
            }
        }





        private void LV_PANEL_LIST_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            SortOrder sortOrder;

            if (e.Column == lastSortedColumn)
                // 이전과 동일한 컬럼 클릭 시 정렬 방향 변경
                sortOrder = lastSortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            else
                // 새로운 컬럼 클릭 시 기본적으로 오름차순
                sortOrder = SortOrder.Ascending;

            LV_PANEL_LIST.ListViewItemSorter = new ListViewSorter(e.Column, sortOrder);

            // 정렬 후 상태 저장
            lastSortedColumn = e.Column;
            lastSortOrder = sortOrder;
        }

        private void CBB_RECIPE_TextUpdate(object sender, EventArgs e)
        {
        }

        private void CBB_RECIPE_SelectionChangeCommitted(object sender, EventArgs e)
        {
        }



        private void BTN_WRITE_EXCEL_Click(object sender, EventArgs e)
        {
            if (BTN_WRITE_EXCEL.BackColor == Color.ForestGreen)
            {
                int Row_num = 1;
                int column = 1;
                //엑셀 생성 
                Excel.Application excelApp = new Excel.Application();
                Excel.Workbook workbook = excelApp.Workbooks.Add();
                Excel.Worksheet worksheet = workbook.Worksheets.get_Item(1) as Excel.Worksheet;
                worksheet.Rows.RowHeight = 100;
                System.Windows.Forms.ListViewItem selectedItem = LV_PANEL_LIST.SelectedItems[0]; // 첫 번째 선택된 항목 접근
                string select_pid = selectedItem.SubItems[1].Text.ToString();  //시리얼 넘버 
                string vpnumber = selectedItem.SubItems[6].Text.ToString();  //시리얼 넘버 접근 
                //1. 헤더 부터 (셀 단위 COM 호출 최소화)
                List<object> headerValues = new List<object>();
                for (int i = 0; i < LV_PANEL_LIST.Columns.Count; i++)
                {
                    headerValues.Add(LV_PANEL_LIST.Columns[i].Text);
                }
                headerValues.Add("Defect Image");

                if (View_mode == 1)
                {
                    foreach (var kvp in Feature_row[0])
                    {
                        headerValues.Add(kvp.Key);
                    }
                }
                else
                {
                    foreach (var kvp in Feature_row_post[0])
                    {
                        headerValues.Add(kvp.Key);
                    }
                }

                object[,] headerArray = new object[1, headerValues.Count];
                for (int i = 0; i < headerValues.Count; i++)
                {
                    headerArray[0, i] = headerValues[i];
                }

                Excel.Range headerRange = worksheet.Range[
                    worksheet.Cells[Row_num, 1],
                    worksheet.Cells[Row_num, headerValues.Count]];
                headerRange.Value2 = headerArray;
                worksheet.Columns[LV_PANEL_LIST.Columns.Count + 1].ColumnWidth = 100;

                Row_num++;

                Dictionary<string, int> inspectionIndexBySerial = new Dictionary<string, int>(StringComparer.Ordinal);
                for (int i = 0; i < inspectionDataList.Count; i++)
                {
                    string serial = inspectionDataList[i].SerialNumber;
                    if (!inspectionIndexBySerial.ContainsKey(serial))
                    {
                        inspectionIndexBySerial.Add(serial, i);
                    }
                }

                //패널별로 반복문
                for (int panel_num = 0; panel_num < LV_PANEL_LIST.SelectedItems.Count; panel_num++)
                {
                    //현재 상황 확인하기 위한 console 창
                    Console.WriteLine("WRITE PANEL_INFO " + panel_num + " / " + LV_PANEL_LIST.SelectedItems.Count);
                    try
                    {
                        System.Windows.Forms.ListViewItem selectedItem_info = LV_PANEL_LIST.SelectedItems[panel_num]; // 첫 번째 선택된 항목 접근
                        select_pid = selectedItem_info.SubItems[1].Text.ToString();  //시리얼 넘버 접근 
                        vpnumber = selectedItem_info.SubItems[6].Text.ToString();  //시리얼 넘버 접근 
                        int inspectionIndex;
                        if (!inspectionIndexBySerial.TryGetValue(select_pid, out inspectionIndex))
                        {
                            continue;
                        }

                        insp_info.listview_index = inspectionIndex;
                        insp_info.insp_Data = inspectionDataList[insp_info.listview_index].InspectionDate;
                        insp_info.Pid = inspectionDataList[insp_info.listview_index].SerialNumber;
                        insp_info.Vision_Num = "VP0" + vpnumber.Substring(vpnumber.Length - 1);

                        int vp_num_int = Convert.ToInt32(insp_info.Vision_Num.Substring(3, 1));
                        int vpIndex = vp_num_int - 1;
                        string basepath = vp_info_list[vp_num_int - 1].Path_Vp_Result;
                        string img_path = basepath + "\\" + insp_info.insp_Data + "\\" + Model_name + "\\" + insp_info.Pid + "_" + insp_info.Vision_Num;

                        //string img_path = basepath + insp_info.insp_Data + "\\" + "x2292" + "\\" + insp_info.Pid + "_" + insp_info.Vision_Num;
                        string[] files = Directory.GetFiles(img_path, "*Pre*.csv");

                        Crop_bin_path_Pre[vpIndex] = Directory.GetFiles(img_path, "*Pre*.bin")[0];
                        Crop_bin_path_Post[vpIndex] = Directory.GetFiles(img_path, "*Post*.bin")[0];
                        if (files.Length > 0)
                            listview2_ReadCSV_Only(files[0]);

                        int image_column = 0;
                        // 실 정보 쓰는 곳 
                        if (View_mode == 1)
                        {
                            for (int feature_count = 0; feature_count < Feature_row.Count; feature_count++)
                            {

                                //FILTER 1) JUDGE 
                                string Filter_Defect_Judge = CBB_DEFECT_JUDGE.Text.ToString();
                                if (Filter_Defect_Judge == "OK")
                                    if (Feature_row[feature_count]["DefectJudge"]?.ToString() != "OK")
                                        continue;
                                    else { }
                                else if (Filter_Defect_Judge == "NG")
                                    if (Feature_row[feature_count]["DefectJudge"]?.ToString() != "NG")
                                        continue;
                                    else { }
                                else { }

                                //FILTER 2) Classify_Group
                                string Filter_Defect_Classify_Group = CBB_DEFECT_NAME_EX.Text.ToString();
                                if (Filter_Defect_Classify_Group == "ALL")
                                { }
                                else if (Filter_Defect_Classify_Group != Feature_row[feature_count]["Classify_Group"]?.ToString())
                                    continue;
                                else
                                { }






                                column = 1;
                                object[,] dataArray_panel = new object[1, selectedItem_info.SubItems.Count]; // +1은 헤더용
                                int col_index_panel = 0;
                                for (int i = 0; i < selectedItem_info.SubItems.Count; i++)
                                {
                                    string columnText = selectedItem_info.SubItems[i].Text;
                                    dataArray_panel[0, col_index_panel] = columnText;
                                    col_index_panel++;
                                }
                                Excel.Range range_panel = worksheet.Range[worksheet.Cells[Row_num, column], worksheet.Cells[Row_num, column + selectedItem_info.SubItems.Count - 1]];
                                range_panel.Value2 = dataArray_panel;

                                column += selectedItem_info.SubItems.Count;

                                //이미지 넣어주기
                                image_column = column;
                                if (CB_SAVE_IMAGE.Checked)
                                {
                                    try
                                    {
                                        using (Bitmap bitmap = LoadFileNamesFromBinary(
                                            Crop_bin_path_Pre[vpIndex],
                                            0, feature_count))
                                        {
                                            string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".png");
                                            bitmap.Save(tempFilePath, ImageFormat.Png);

                                            try
                                            {
                                                Excel.Range cell = worksheet.Cells[Row_num, column];
                                                float left = (float)cell.Left;
                                                float top = (float)cell.Top;
                                                float width = bitmap.Width;
                                                float height = 100f;

                                                Excel.Shape shape = worksheet.Shapes.AddPicture(
                                                    tempFilePath,
                                                    Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoTrue,
                                                    left, top, width, height);

                                                shape.Placement = Excel.XlPlacement.xlMoveAndSize;
                                            }
                                            finally
                                            {
                                                if (File.Exists(tempFilePath))
                                                {
                                                    File.Delete(tempFilePath);
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("Image Paste Error : " + ex.ToString());
                                    }
                                    finally
                                    {
                                        //// 임시 파일 정리
                                        //if (File.Exists(tempFilePath))
                                        //    File.Delete(tempFilePath);
                                    }
                                }


                                if (false)
                                {
                                    try
                                    {
                                        // Bitmap 이미지를 로드
                                        Bitmap bitmap = LoadFileNamesFromBinary(Crop_bin_path_Pre[vpIndex], 0, feature_count);//vp2추가필요

                                        // Bitmap 이미지를 메모리 스트림으로 변환
                                        using (MemoryStream memoryStream = new MemoryStream())
                                        {
                                            bitmap.Save(memoryStream, System.Drawing.Imaging.ImageFormat.Png);
                                            byte[] imageBytes = memoryStream.ToArray();

                                            // 임시 파일 생성
                                            string tempFilePath = Path.GetTempFileName();
                                            File.WriteAllBytes(tempFilePath, imageBytes);

                                            // 엑셀 워크시트에 이미지 삽입
                                            Excel.Range cell = worksheet.Cells[Row_num, column];
                                            float left = (float)((double)cell.Left);
                                            float top = (float)((double)cell.Top);
                                            float width = bitmap.Width;
                                            float height = 100;

                                            Excel.Shape shape = worksheet.Shapes.AddPicture(
                                                tempFilePath,  // 임시 파일 경로
                                                Microsoft.Office.Core.MsoTriState.msoFalse,
                                                Microsoft.Office.Core.MsoTriState.msoCTrue,
                                                left, top, width, height
                                            );

                                            shape.Placement = Excel.XlPlacement.xlMoveAndSize;


                                            // 임시 파일 삭제
                                            File.Delete(tempFilePath);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("Image Paste Error :" + ex.ToString());
                                    }

                                }

                                //-----

                                column++;

                                if (View_mode == 1)
                                {
                                    object[,] dataArray_defect = new object[1, Feature_row[feature_count].Count]; // +1은 헤더용
                                    int col_index_defect = 0;
                                    foreach (var kvp in Feature_row[feature_count])
                                    {
                                        dataArray_defect[0, col_index_defect] = kvp.Value;
                                        col_index_defect++;
                                    }
                                    Excel.Range range_defect = worksheet.Range[worksheet.Cells[Row_num, column], worksheet.Cells[Row_num, column + Feature_row[feature_count].Count - 1]];
                                    range_defect.Value2 = dataArray_defect;
                                }
                                else
                                {
                                    object[,] dataArray_defect = new object[1, Feature_row_post[feature_count].Count]; // +1은 헤더용
                                    int col_index_defect = 0;
                                    foreach (var kvp in Feature_row_post[feature_count])
                                    {
                                        dataArray_defect[0, col_index_defect] = kvp.Value;
                                        col_index_defect++;
                                    }
                                    Excel.Range range_defect = worksheet.Range[worksheet.Cells[Row_num, column], worksheet.Cells[Row_num, column + Feature_row_post[feature_count].Count - 1]];
                                    range_defect.Value2 = dataArray_defect;
                                }

                                Row_num++;
                            }
                        }
                        else
                        {
                            for (int feature_count = 0; feature_count < Feature_row_post.Count; feature_count++)
                            {
                                String Filter_Defect_Judge = CBB_DEFECT_JUDGE.Text.ToString();
                                if (Filter_Defect_Judge == "OK")
                                    if (Feature_row_post[feature_count]["DefectJudge"]?.ToString() != "OK")
                                        continue;
                                    else { }
                                else if (Filter_Defect_Judge == "NG")
                                    if (Feature_row_post[feature_count]["DefectJudge"]?.ToString() != "NG")
                                        continue;
                                    else { }
                                else { }
                                column = 1;
                                object[,] dataArray_panel = new object[1, selectedItem_info.SubItems.Count]; // +1은 헤더용
                                int col_index_panel = 0;
                                for (int i = 0; i < selectedItem_info.SubItems.Count; i++)
                                {
                                    string columnText = selectedItem_info.SubItems[i].Text;
                                    dataArray_panel[0, col_index_panel] = columnText;
                                    col_index_panel++;
                                }
                                Excel.Range range_panel = worksheet.Range[worksheet.Cells[Row_num, column], worksheet.Cells[Row_num, column + selectedItem_info.SubItems.Count - 1]];
                                range_panel.Value2 = dataArray_panel;

                                column += selectedItem_info.SubItems.Count;

                                //이미지 넣어주기
                                image_column = column;



                                if (CB_SAVE_IMAGE.Checked)
                                {
                                    try
                                    {
                                        using (Bitmap bitmap = LoadFileNamesFromBinary(Crop_bin_path_Pre[vpIndex], 0, feature_count))
                                        {
                                            string tempFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".png");
                                            bitmap.Save(tempFilePath, ImageFormat.Png);

                                            try
                                            {
                                                Excel.Range cell = worksheet.Cells[Row_num, column];
                                                float left = (float)((double)cell.Left);
                                                float top = (float)((double)cell.Top);
                                                float width = bitmap.Width;
                                                float height = 100;

                                                Excel.Shape shape = worksheet.Shapes.AddPicture(
                                                    tempFilePath,
                                                    Microsoft.Office.Core.MsoTriState.msoFalse,
                                                    Microsoft.Office.Core.MsoTriState.msoCTrue,
                                                    left, top, width, height
                                                );
                                                shape.Placement = Excel.XlPlacement.xlMoveAndSize;
                                            }
                                            finally
                                            {
                                                if (File.Exists(tempFilePath))
                                                {
                                                    File.Delete(tempFilePath);
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine("Image Paste Error :" + ex.ToString());
                                    }

                                }

                                //-----

                                column++;

                                if (View_mode == 1)
                                {
                                    object[,] dataArray_defect = new object[1, Feature_row[feature_count].Count]; // +1은 헤더용
                                    int col_index_defect = 0;
                                    foreach (var kvp in Feature_row[feature_count])
                                    {
                                        dataArray_defect[0, col_index_defect] = kvp.Value;
                                        col_index_defect++;
                                    }
                                    Excel.Range range_defect = worksheet.Range[worksheet.Cells[Row_num, column], worksheet.Cells[Row_num, column + Feature_row[feature_count].Count - 1]];
                                    range_defect.Value2 = dataArray_defect;
                                }
                                else
                                {
                                    object[,] dataArray_defect = new object[1, Feature_row_post[feature_count].Count]; // +1은 헤더용
                                    int col_index_defect = 0;
                                    foreach (var kvp in Feature_row_post[feature_count])
                                    {
                                        dataArray_defect[0, col_index_defect] = kvp.Value;
                                        col_index_defect++;
                                    }
                                    Excel.Range range_defect = worksheet.Range[worksheet.Cells[Row_num, column], worksheet.Cells[Row_num, column + Feature_row_post[feature_count].Count - 1]];
                                    range_defect.Value2 = dataArray_defect;
                                }

                                Row_num++;
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Image Folder Emtpy Error :" + ex.ToString());
                    }


                }
                workbook.SaveAs(Excel_wirte_path + "\\RESULT.xlsx");
                MessageBox.Show("Excel_Write_OK");
                workbook.Close();

            }
        }

        private void CBB_DEFECT_JUDGE_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void MTP_PRODUCTSEARCH_Click(object sender, EventArgs e)
        {

        }

        private void BTN_EXCEL_FOLDER_Click(object sender, EventArgs e)
        {
            CommonOpenFileDialog dialog = new CommonOpenFileDialog();
            dialog.IsFolderPicker = true; // true : 폴더 선택 / false : 파일 선택

            if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                Excel_wirte_path = dialog.FileName;
        }

        private void CBB_JUDGE_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void PB_DEFECT_ARRAY_Click(object sender, EventArgs e)
        {

        }

        private void PB_DEFECT_ARRAY_DoubleClick(object sender, EventArgs e)
        {

        }

        private void DGV_MAIN_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DGV_MAIN_INDEX = DGV_MAIN.CurrentCell.RowIndex;
            }
            catch { }

        }

        private void PB_MAIN_Click(object sender, EventArgs e)
        {
            ImageControl image_form = new ImageControl((Bitmap)PB_MAIN.Image);
            image_form.Show();
        }

        private void DGV_MAIN_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                DGV_MAIN_INDEX = DGV_MAIN.CurrentCell.RowIndex;
            }
            catch { }
        }

        private void LV_DEFECT_LIST_ColumnClick(object sender, ColumnClickEventArgs e)
        {

            SortOrder sortOrder;

            if (e.Column == lastSortedColumn)
            {
                // 이전과 동일한 컬럼 클릭 시 정렬 방향 변경
                sortOrder = lastSortOrder == SortOrder.Ascending ? SortOrder.Descending : SortOrder.Ascending;
            }
            else
            {
                // 새로운 컬럼 클릭 시 기본적으로 오름차순
                sortOrder = SortOrder.Ascending;
            }

            LV_DEFECT_LIST.ListViewItemSorter = new ListViewSorter(e.Column, sortOrder);

            // 정렬 후 상태 저장
            lastSortedColumn = e.Column;
            lastSortOrder = sortOrder;

        }

        private void CBB_JUDGE_DropDown(object sender, EventArgs e)
        {
            ApplyComboBoxDropDownWidth(sender as System.Windows.Forms.ComboBox);
        }

        private void RegisterComboBoxDropDownAutoWidth(Control parent)
        {
            if (parent == null)
            {
                return;
            }

            foreach (Control control in parent.Controls)
            {
                System.Windows.Forms.ComboBox combo = control as System.Windows.Forms.ComboBox;
                if (combo != null)
                {
                    combo.DropDown -= ComboBox_DropDown_AutoWidth;
                    combo.DropDown += ComboBox_DropDown_AutoWidth;
                    ApplyComboBoxDropDownWidth(combo);
                }

                if (control.HasChildren)
                {
                    RegisterComboBoxDropDownAutoWidth(control);
                }
            }
        }

        private void ComboBox_DropDown_AutoWidth(object sender, EventArgs e)
        {
            ApplyComboBoxDropDownWidth(sender as System.Windows.Forms.ComboBox);
        }

        private void ApplyComboBoxDropDownWidth(System.Windows.Forms.ComboBox combo)
        {
            if (combo == null)
            {
                return;
            }

            int preferredWidth = GetLargestTextExtent(combo);
            if (preferredWidth <= 0)
            {
                return;
            }

            int minWidth = combo.Width;
            int finalWidth = preferredWidth > minWidth ? preferredWidth : minWidth;
            int maxWidth = Screen.FromControl(combo).WorkingArea.Width - 20;
            if (finalWidth > maxWidth)
            {
                finalWidth = maxWidth;
            }

            combo.DropDownWidth = finalWidth;
        }

        private int GetLargestTextExtent(System.Windows.Forms.ComboBox cbo)
        {
            int maxLen = -1;

            if (cbo.Items.Count >= 1)
            {
                using (Graphics g = cbo.CreateGraphics())
                {
                    int vertScrollBarWidth = 0;
                    if (cbo.Items.Count > cbo.MaxDropDownItems)
                    {
                        vertScrollBarWidth = SystemInformation.VerticalScrollBarWidth;
                    }
                    for (int nLoopCut = 0; nLoopCut < cbo.Items.Count; nLoopCut++)
                    {
                        int newWidth = (int)g.MeasureString(cbo.Items[nLoopCut].ToString(), cbo.Font).Width + vertScrollBarWidth;
                        if (newWidth > maxLen)
                            maxLen = newWidth;
                    }
                }
            }
            return maxLen;
        }

        private void CBB_RECIPE_DropDown(object sender, EventArgs e)
        {
            ApplyComboBoxDropDownWidth(sender as System.Windows.Forms.ComboBox);
        }

        private void PB_DEFECTMAP_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            // 여기에다가 마우스올리면 해당디펙 선택가능하게하는 기능 만들기, 

            System.Drawing.Point mousePosition = new System.Drawing.Point(e.Location.X, e.Location.Y);

            //// 5픽셀 이내에 있는지 확인
            //foreach (var defect in Defect_Position)
            //{
            //}

            for (int i = 0; i < Defect_Position.Count(); i++)
            {
                if (IsWithinRange(Defect_Position[i], mousePosition, 5))
                {
                    //현재 거리가 5pixel 
                    Defect_map_select_index = i;

                }
            }

        }

        public static bool IsWithinRange(System.Drawing.Point point1, System.Drawing.Point point2, int range)
        {
            double distance = Math.Sqrt(Math.Pow(point1.X - point2.X, 2) + Math.Pow(point1.Y - point2.Y, 2));
            return distance <= range;
        }

        private int GetDefectSourceIndexFromListItem(System.Windows.Forms.ListViewItem item)
        {
            if (item == null || item.SubItems.Count == 0)
            {
                return -1;
            }

            int rawIndex;
            if (!int.TryParse(item.SubItems[0].Text, out rawIndex))
            {
                return -1;
            }

            return rawIndex - 1;
        }

        private int FindListViewIndexByDefectSourceIndex(int defectSourceIndex)
        {
            if (defectSourceIndex < 0)
            {
                return -1;
            }

            for (int i = 0; i < LV_DEFECT_LIST.Items.Count; i++)
            {
                if (GetDefectSourceIndexFromListItem(LV_DEFECT_LIST.Items[i]) == defectSourceIndex)
                {
                    return i;
                }
            }

            return -1;
        }

        private void PB_DEFECTMAP_MouseClick(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (Defect_map_select_index < 0)
            {
                return;
            }

            int listViewIndex = FindListViewIndexByDefectSourceIndex(Defect_map_select_index);
            if (listViewIndex < 0 || listViewIndex >= LV_DEFECT_LIST.Items.Count)
            {
                return;
            }

            LV_DEFECT_LIST.BeginUpdate();
            try
            {
                foreach (System.Windows.Forms.ListViewItem item in LV_DEFECT_LIST.SelectedItems)
                {
                    item.Selected = false; // 선택 상태 해제
                }

                LV_DEFECT_LIST.Items[listViewIndex].Selected = true; // 행 선택
                LV_DEFECT_LIST.Items[listViewIndex].Focused = true;  // 포커스 설정
                LV_DEFECT_LIST.EnsureVisible(listViewIndex);         // 선택된 행이 보이도록 스크롤
            }
            finally
            {
                LV_DEFECT_LIST.EndUpdate();
            }
        }

        private void PB_DEFECTMAP_Click(object sender, EventArgs e)
        {

        }

        private void CBB_VP_NUM_SelectedIndexChanged(object sender, EventArgs e)
        {
            int select_num = CBB_VP_NUM.SelectedIndex;
            insp_info.Vision_Num = "VP0" + (select_num + 1).ToString();
            listview2_init(select_num + 1);
        }

        private void LV_DEFECT_LIST_SelectedIndexChanged(object sender, EventArgs e)
        {
            LV_DEFECT_LIST_SelectedIndexChangedAsync(this, new EventArgs());
            //DGV_MAIN.SuspendLayout(); // 레이아웃 계산 중지
            ////항목이 바뀔때 선택된 인덱스를 먼저 찾고 해당 이미지 경로에 접근해서 이미지가 있는경우에 위에 보여준다~ 
            //int selectedIndex;
            //if (LV_DEFECT_LIST.SelectedIndices.Count > 0)
            //{


            //    // 선택된 인덱스를 사용하여 작업 수행
            //    selectedIndex = LV_DEFECT_LIST.SelectedIndices[0];
            //    try // Defect Featrue 출력 
            //    {

            //        System.Windows.Forms.ListViewItem selectedItem = LV_DEFECT_LIST.Items[selectedIndex];

            //        // 특정 열의 값 가져오기 (예: 첫 번째 열 = 인덱스 0)
            //        int columnValue = Convert.ToInt32(selectedItem.SubItems[0].Text) - 1;
            //        DGV_MAIN.Rows.Clear();
            //        int index = 0;
            //        foreach (var kvp in Feature_row[columnValue])
            //        {
            //            DGV_MAIN.Rows.Add(index, kvp.Key, kvp.Value);
            //            index++;
            //        }

            //        //이미지 보여주기
            //        int Ptn_num = Convert.ToInt32(Feature_row[columnValue]["ptnIdx"]?.ToString());
            //        int defectIdx = Convert.ToInt32(Feature_row[columnValue]["defectIdx"]?.ToString());
            //        LoadFileNamesFromBinary(Crop_bin_path, Ptn_num, defectIdx);
            //        select_nymber_defect = columnValue;
            //        PB_DEFECTMAP.Invalidate();
            //        check_picturebox = true;


            //        //DGV_MAIN 선택된 행이있으면 해당 행으로 옮겨가자 
            //        DGV_MAIN.Rows[DGV_MAIN_INDEX].Selected = true;
            //        DGV_MAIN.CurrentCell = DGV_MAIN.Rows[DGV_MAIN_INDEX].Cells[0];
            //    }
            //    catch
            //    {
            //        //feature img_path 경로가 없을때  
            //    }
            //}
            //DGV_MAIN.ResumeLayout(); // 레이아웃 계산 재개

            //해당 Feature 정보로 classify 한번 돌리기 
        }

        private void LV_PANEL_LIST_DrawColumnHeader(object sender, DrawListViewColumnHeaderEventArgs e)
        {
            e.Graphics.FillRectangle(Brushes.DarkGray, e.Bounds);
            e.DrawText();
        }

        private void LV_PANEL_LIST_DrawItem(object sender, DrawListViewItemEventArgs e)
        {
            e.DrawDefault = true;
        }

        private void CB_SIMULATION_PANEL_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void CBB_DEFECT_NAME_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void GB_CONDITIONSEARCH_Enter(object sender, EventArgs e)
        {

        }

        private void metroButton3_Click(object sender, EventArgs e)
        {

            // POST  ▶ PRE
            if (CheckClassifier == 0)
            {
                RTB_ORIGIN_POST.Hide();
                RTB_REPLACE_POST.Hide();
                RTB_ORIGIN_PRE.Show();
                RTB_REPLACE_PRE.Show();
                BTN_CHANGE.Text = "POST ▶";
                CheckClassifier = 1;
            }

            // PRE   ▶ POST
            else
            {
                RTB_ORIGIN_POST.Show();
                RTB_REPLACE_POST.Show();
                RTB_ORIGIN_PRE.Hide();
                RTB_REPLACE_PRE.Hide();
                BTN_CHANGE.Text = "PRE ▶";
                CheckClassifier = 0;
            }
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void CBB_PRODUCT_CLASSIFY_FILE_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            TextBoxBase textBox = sender as TextBoxBase;
            if (textBox != null)
            {
                ApplyPanelSerialFilter(textBox.Text);
            }

        }

        private void metroTextBox2_TextChanged(object sender, EventArgs e)
        {
            TextBoxBase textBox = sender as TextBoxBase;
            if (textBox != null)
            {
                ApplyPanelSerialFilter(textBox.Text);
            }

        }

        private void ApplyPanelSerialFilter(string keyword)
        {
            if (string.IsNullOrWhiteSpace(keyword))
            {
                listview_print(inspectionDataList);
                return;
            }

            List<InspectionData> filteredList = inspectionDataList
                .Where(data => !string.IsNullOrEmpty(data.SerialNumber)
                    && data.SerialNumber.IndexOf(keyword, StringComparison.OrdinalIgnoreCase) >= 0)
                .ToList();

            // LV_PANEL_LIST_DoubleClick / SelectedIndexChanged는 기존 inspectionDataList 기준으로 panel 정보를 찾기 때문에
            // 표시 목록만 필터링해도 더블클릭 동작은 기존과 동일하게 연동된다.
            listview_print(filteredList);
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (CB_CHECK_ID.Checked)
            {
                TB_IDSEARCH.Height = 300;
                TB_IDSEARCH.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
                groupBox1.Height = 560;
                groupBox1.BringToFront();
            }
            else
            {
                TB_IDSEARCH.Height = 25;
                TB_IDSEARCH.ScrollBars = System.Windows.Forms.ScrollBars.None;
                groupBox1.Height = 60;

            }
        }

        private void LV_DEFECT_LIST_MouseDown(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            int selectedIndex = -1;
            if (LV_DEFECT_LIST.SelectedIndices.Count > 0)
            {
                // 선택된 인덱스를 사용하여 작업 수행
                selectedIndex = LV_DEFECT_LIST.SelectedIndices[0];

            }
            //오른쪽 클릭했을경우 이미지 저장 화면 출력하기
            if (e.Button == MouseButtons.Right && selectedIndex != -1)
            {
                CommonOpenFileDialog dialog = new CommonOpenFileDialog();
                dialog.IsFolderPicker = false; // true : 폴더 선택 / false : 파일 선택

                if (dialog.ShowDialog() == CommonFileDialogResult.Ok)
                {
                    System.Drawing.Image image = Clipboard.GetImage();
                    image.Save(dialog.FileName, ImageFormat.Png);

                }
            }
        }


        //네비게이션 기능 --------------------------------
        private void CBB_NAVI_JUDGE_SelectedIndexChanged(object sender, EventArgs e)
        {
            navi.judge = CBB_NAVI_JUDGE.Text;
        }

        private void CBB_NAVI_CLASSIFY_SelectedIndexChanged(object sender, EventArgs e)
        {
            navi.judge = CBB_NAVI_CLASSIFY.Text;
        }
        private void RB_NAVI_PANEL_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_NAVI_PANEL.Checked)
                navi.type = 1;
            else
                navi.type = 2;
        }
        private void RB_NAVI_DEFECT_CheckedChanged(object sender, EventArgs e)
        {
            if (RB_NAVI_PANEL.Checked)
                navi.type = 1;
            else
                navi.type = 2;
        }

        private void BTN_NAVI_PREV_Click(object sender, EventArgs e)
        {
            //후보점 단위인지, 패널단위인지 확인
            if (RB_NAVI_PANEL.Checked) //패널 단위인경우 
            {

            }
            else  // 후보점 단위인경우 
            {

                // 선택된 항목이 없으면 마지막 항목을 선택
                if (LV_DEFECT_LIST.SelectedItems.Count == 0)
                {
                    //이전패널로 넘기는거구현해야함 
                }

                // 현재 선택된 항목의 인덱스
                int curIdx = LV_DEFECT_LIST.SelectedItems[0].Index;

                // 첫 번째 항목이면 더 이상 이동할 수 없음 → 메시지 표시
                if (curIdx == 0)
                {
                    //이전패널로 넘기는거구현해야함 
                }

                // 이전 항목 선택
                LV_DEFECT_LIST.Items[curIdx - 1].Selected = true;
                // 현재 항목 해제
                LV_DEFECT_LIST.Items[curIdx].Selected = false;

                LV_DEFECT_LIST.Focus();   // 포커스 이동

            }
        }

        private void BTN_NAVI_NEXT_Click(object sender, EventArgs e)
        {


            //후보점 단위인지, 패널단위인지 확인
            if (RB_NAVI_PANEL.Checked) //패널 단위인경우 
            {
                //단순하게 다음패널로 넘기기


                //다음패널 어떻게 보여줄거? 

                //string panelid = inspectionDataList[navi.Panel_index+1].SerialNumber;
                //BTN_DEFECT_PANEL.Text = panelid;
                //Validate();
                //int vp_num_int = Convert.ToInt32(insp_info.Vision_Num.Substring(3, 1));
                //listview2_init(1);


            }
            else  // 후보점 단위인경우 
            {
                // 현재 선택된 항목이 있는지 확인
                if (LV_DEFECT_LIST.SelectedItems.Count == 0)
                {
                    // 선택된 항목이 없으면 첫 번째 항목을 선택
                    if (LV_DEFECT_LIST.Items.Count > 0)
                    {
                        LV_DEFECT_LIST.Items[0].Selected = true;
                        navi.Defect_index = 0;
                    }
                    return;
                }
                else
                {
                    // 현재 선택된 항목의 인덱스 가져오기
                    int currentIndex = LV_DEFECT_LIST.SelectedItems[0].Index;
                    // 마지막 항목인지 확인
                    if (currentIndex == LV_DEFECT_LIST.Items.Count - 1)
                    {
                        // 마지막 항목이면 다음 패널로 넘겨줘야함  위에꺼 구현해서 연동하자 

                    }

                    // 다음 항목 선택
                    LV_DEFECT_LIST.Items[currentIndex + 1].Selected = true;
                    //이전항목 해재
                    LV_DEFECT_LIST.Items[currentIndex].Selected = false;
                    // 선택된 항목을 강조하기 위해 포커스 이동 (선택 모드에 따라 필요 없을 수도 있음)
                    LV_DEFECT_LIST.Focus();
                    navi.Defect_index = currentIndex + 1;
                }

            }



        }



        //----------------------------------------------------------------



        private void TB_RECIPE_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!CB_RECIPE_WRITE.Checked)
            {
                // 숫자(0-9)가 아니면 입력 차단
                if (!char.IsControl(e.KeyChar) && !char.IsDigit(e.KeyChar))
                {
                    e.Handled = true; // 입력 차단
                }
            }

        }

        private void PB_DEFECTMAP_Paint(object sender, PaintEventArgs e)
        {
            try
            {
                Bitmap bitmap = new Bitmap((int)(insp_info.panel_width / picturebox_ratio_x), (int)(insp_info.panel_Height / picturebox_ratio_y));

                // Graphics 객체를 사용하여 Bitmap에 그리기
                using (Graphics g = Graphics.FromImage(bitmap))
                {
                    DrawDefectMapGrid(g, bitmap.Width, bitmap.Height);

                    int radian = 5;
                    int pos_x = 0;
                    int pos_y = 0;
                    int count = 0;
                    if (View_mode == 1)
                        count = Feature_row.Count;
                    else
                        count = Feature_row_post.Count;

                    for (int i = 0; i < count; i++)
                    {
                        if (Judge_mode == 0)//ok,ng둘다
                        {

                            if (radian > Defect_Position[i].X)
                                pos_x = radian;
                            else if (Defect_Position[i].X > bitmap.Width - radian)
                                pos_x = bitmap.Width - radian;
                            else
                                pos_x = Defect_Position[i].X;

                            if (radian > Defect_Position[i].Y)
                                pos_y = radian;
                            else if (Defect_Position[i].Y > bitmap.Height - radian)
                                pos_y = bitmap.Height - radian;
                            else
                                pos_y = Defect_Position[i].Y;

                            if (i == select_nymber_defect)
                                g.FillEllipse(Brushes.Red, pos_x - radian, pos_y - radian, radian * 2, radian * 2); // 반지름 5짜리 원
                            else if (i == Defect_map_select_index)
                                g.FillEllipse(Brushes.Yellow, pos_x - radian, pos_y - radian, radian * 2, radian * 2); // 반지름 5짜리 원
                            else
                                g.FillEllipse(Brushes.White, pos_x - (radian / 2), pos_y - (radian / 2), radian, radian); // 반지름 5짜리 원}
                        }
                        else if (Judge_mode == 2)
                        {
                            if (!Defect_Position_Judge[i])
                            {

                                if (radian > Defect_Position[i].X)
                                    pos_x = radian;
                                else if (Defect_Position[i].X > bitmap.Width - radian)
                                    pos_x = bitmap.Width - radian;
                                else
                                    pos_x = Defect_Position[i].X;

                                if (radian > Defect_Position[i].Y)
                                    pos_y = radian;
                                else if (Defect_Position[i].Y > bitmap.Height - radian)
                                    pos_y = bitmap.Height - radian;
                                else
                                    pos_y = Defect_Position[i].Y;

                                if (i == select_nymber_defect)
                                    g.FillEllipse(Brushes.Red, pos_x - radian, pos_y - radian, radian * 2, radian * 2); // 반지름 5짜리 원
                                else if (i == Defect_map_select_index)
                                    g.FillEllipse(Brushes.Yellow, pos_x - radian, pos_y - radian, radian * 2, radian * 2); // 반지름 5짜리 원
                                else
                                    g.FillEllipse(Brushes.OrangeRed, pos_x - (radian / 2), pos_y - (radian / 2), radian, radian); // 반지름 5짜리 원}
                            }
                        }
                        else
                        {
                            if (Defect_Position_Judge[i])
                            {

                                if (radian > Defect_Position[i].X)
                                    pos_x = radian;
                                else if (Defect_Position[i].X > bitmap.Width - radian)
                                    pos_x = bitmap.Width - radian;
                                else
                                    pos_x = Defect_Position[i].X;

                                if (radian > Defect_Position[i].Y)
                                    pos_y = radian;
                                else if (Defect_Position[i].Y > bitmap.Height - radian)
                                    pos_y = bitmap.Height - radian;
                                else
                                    pos_y = Defect_Position[i].Y;

                                if (i == select_nymber_defect)
                                    g.FillEllipse(Brushes.Red, pos_x - radian, pos_y - radian, radian * 2, radian * 2); // 반지름 5짜리 원
                                else if (i == Defect_map_select_index)
                                    g.FillEllipse(Brushes.Yellow, pos_x - radian, pos_y - radian, radian * 2, radian * 2); // 반지름 5짜리 원
                                else
                                    g.FillEllipse(Brushes.White, pos_x - (radian / 2), pos_y - (radian / 2), radian, radian); // 반지름 5짜리 원}
                            }
                        }
                    }

                }

                SetPictureBoxImage(PB_DEFECTMAP, bitmap);
            }
            catch { }
        }


        private void SyncDefectSelectionAfterFilter()
        {
            int currentSourceIndex = (int)select_nymber_defect;
            int listIndex = FindListViewIndexByDefectSourceIndex(currentSourceIndex);

            if (listIndex < 0)
            {
                select_nymber_defect = 9999;
                lastDefectListSelectedIndex = -1;
                Defect_map_select_index = -1;
                PB_DEFECTMAP.Invalidate();
                return;
            }

            if (listIndex >= 0 && listIndex < LV_DEFECT_LIST.Items.Count)
            {
                LV_DEFECT_LIST.Items[listIndex].Selected = true;
                LV_DEFECT_LIST.Items[listIndex].Focused = true;
                LV_DEFECT_LIST.EnsureVisible(listIndex);
            }
        }

        private void BTN_DEFECT_NAME_Click(object sender, EventArgs e)
        {
            string DefectJudge = CBB_SINGLE_RECIPE.Text;

            LV_DEFECT_LIST.BeginUpdate();
            try
            {
                LV_DEFECT_LIST.Items.Clear();

                for (int i = 0; i < Feature_row.Count; i++)
                {
                    if ((Feature_row[i]["PRE_CLASSIFY_GROUP"]?.ToString() != DefectJudge) && DefectJudge != "ALL")
                        continue;

                    var listViewItem = new System.Windows.Forms.ListViewItem((i + 1).ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["DefectJudge"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["PRE_CLASSIFY_GROUP"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["PIXEL_X"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["PIXEL_Y"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["Area"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["GrayAVG_Pre"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["GrayMin_Pre"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["GrayMax_Pre"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["OriginalImage_Column"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["OriginalImage_Row"]?.ToString());
                    listViewItem.SubItems.Add(Feature_row[i]["PTNType"]?.ToString());
                    LV_DEFECT_LIST.Items.Add(listViewItem);

                }
            }
            finally
            {
                LV_DEFECT_LIST.EndUpdate();
            }

            SyncDefectSelectionAfterFilter();
            BTN_DEFECT_NAME.Text = "Defect Name (" + DefectJudge + ")";
        }

        public Bitmap LoadFileNamesFromBinary(string filePath, int ptn_num, int defect_num) //filepath  배열순서대로 vp0~
        {
            try
            {
                using (FileStream fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                using (BinaryReader reader = new BinaryReader(fs))
                {
                    // 이미지 개수 읽기
                    int vecSize = reader.ReadInt32();
                    // 이미지 헤더 읽기 (첫 번째 이미지의 헤더만 읽음)
                    int width = reader.ReadInt32();
                    int height = reader.ReadInt32();
                    int size = reader.ReadInt32();

                    Bitmap_Crop = new Bitmap(width, height);
                    for (int i = 0; i < vecSize; i++)
                    {

                        // 메타데이터 읽기
                        int patternIndex = reader.ReadInt32();
                        int defectIndex = reader.ReadInt32();
                        uint lenFileName = reader.ReadUInt32();
                        byte[] fileNameBytes = reader.ReadBytes((int)lenFileName);
                        if (i == defect_num)
                        {
                            //if (ptn_num == patternIndex && defect_num == defectIndex)

                            // Bitmap 객체 생성
                            using (Bitmap bitmap = new Bitmap(width, height))
                            {
                                byte[] img = reader.ReadBytes(size);
                                // 픽셀 값 설정
                                for (int i_h = 0; i_h < height; i_h++)
                                {
                                    for (int j = 0; j < width; j++)
                                    {
                                        // 픽셀 값 가져오기
                                        byte pixelValue = img[i_h * width + j];
                                        // 픽셀 설정
                                        //bitmap.SetPixel(j, i_h, Color.FromArgb(pixelValue, pixelValue, pixelValue));
                                        Bitmap_Crop.SetPixel(j, i_h, Color.FromArgb(pixelValue, pixelValue, pixelValue));
                                    }
                                }
                                SetPictureBoxImage(PB_DEFECT_ARRAY, Bitmap_Crop);
                                originalImage = Bitmap_Crop;
                                Clipboard.SetImage(Bitmap_Crop);
                                return Bitmap_Crop;
                                break;

                            }

                        }
                        else  // 이미지 데이터 건너뛰기
                        {
                            reader.BaseStream.Seek(size, SeekOrigin.Current);
                        }
                    }
                }

            }
            catch
            {
                //Console.Write("왜죽늬?");
            }
            Clipboard.SetImage(Bitmap_Crop);
            return Bitmap_Crop;


        }

        private class ConfigState
        {
            public List<VP_INFO> VpInfoList = new List<VP_INFO>();
            public string ModelName;
            public int ViewMode;
            public List<string> ViewDefectFeatureNames = new List<string>();
            public string PositionFeatureNameX;
            public string PositionFeatureNameY;
            public int SwapX;
            public int SwapY;
            public int ViewFeatureNum;
            public int DefectMapRowCount = 10;
            public int DefectMapColCount = 10;
            public int DefectMapRowColCount = 1;
            public int DefectMapRowSwap = 1;
            public int DefectMapColSwap = 1;
            public string[] IgnoreIndexArray = new string[0];
            public string DiskBase = @"D:\\";
            public string InspectionPathBase = @"LGD_AMI\\Inspection\\";
            public string RecipePathBase = @"LGD_AMI\\Recipe\\";
            public string ResultPathBase = @"LGD_AMI\\RESULT\\";
            public string LogPathBase = @"LGD_AMI\\Log\\";
            public string SimulatorConfigPath = @"..\Simul_Config.ini";
        }

        private class SearchState
        {
            public string RecipePath;
            public DateTime StartDate;
            public DateTime EndDate;
            public List<string> MatchingFiles = new List<string>();
            public List<InspectionData> InspectionDataList = new List<InspectionData>();
            public bool PanelIdSearch = true;
            public List<(string FolderName, string FolderPath)> MatchingFolders = new List<(string FolderName, string FolderPath)>();
        }

        private class DefectState
        {
            public List<Dictionary<string, object>> FeatureRow = new List<Dictionary<string, object>>();
            public List<Dictionary<string, object>> FeatureRowPost = new List<Dictionary<string, object>>();
            public List<string> CropBinPathPre = new List<string>();
            public List<string> CropBinPathPost = new List<string>();
            public List<Classify_info> ClassifyInfos = new List<Classify_info>();
            public List<Classify_Pre_info> ClassifyPreInfos = new List<Classify_Pre_info>();
            public List<Classify_Post_info> ClassifyPostInfos = new List<Classify_Post_info>();
            public List<string> UsedFeaturePre = new List<string>();
            public List<string> UsedFeaturePost = new List<string>();
            public List<System.Drawing.Point> DefectPosition = new List<System.Drawing.Point>();
            public List<bool> DefectPositionJudge = new List<bool>();
            public List<int> VpDefectNum = new List<int>();
        }
    }

    public class ListViewSorter : IComparer
    {
        private int columnIndex;
        private SortOrder sortOrder;

        public ListViewSorter(int columnIndex, SortOrder sortOrder)
        {
            this.columnIndex = columnIndex;
            this.sortOrder = sortOrder;
        }

        public int Compare(object x, object y)
        {
            System.Windows.Forms.ListViewItem item1 = x as System.Windows.Forms.ListViewItem;
            System.Windows.Forms.ListViewItem item2 = y as System.Windows.Forms.ListViewItem;

            string value1 = item1.SubItems[columnIndex].Text;
            string value2 = item2.SubItems[columnIndex].Text;

            // 숫자 비교
            if (int.TryParse(value1, out int num1) && int.TryParse(value2, out int num2))
            {
                return sortOrder == SortOrder.Ascending ? num1.CompareTo(num2) : num2.CompareTo(num1);
            }

            // 문자열 비교
            return sortOrder == SortOrder.Ascending ? string.Compare(value1, value2) : string.Compare(value2, value1);
        }

    }



}
