  using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.Linq;
using System.Text;
using System.Globalization;
using System.Windows.Controls;
using System.Windows.Data;
using System.Collections;
using System.Xml;
using System.Windows.Media;
using System.Collections.ObjectModel;
using System.Windows;
using System.IO;
using System.Reflection;
using Microsoft.Win32;
using Abt.Controls.SciChart.Visuals;
using System.Text.RegularExpressions;
using System.ComponentModel;
using System.Threading.Tasks;
using RDotNet;
using System.Data.Entity;
using Abt.Controls.SciChart.Model.DataSeries;

namespace K_Treasury
{
    /// <summary>
    ///  This class is working with data management. Binding with data base and so on.
    /// </summary>
    public static class K_DataManager
    {
        /// <summary>
        /// R engine.
        /// </summary>
        public static REngine engine = REngine.GetInstance();

        /// <summary>
        /// Connection string to data base.
        /// </summary>
        public static string ConnectionString = Properties.Settings.Default.ConnectionString;

        /// <summary>
        /// ODBC Data Source for R.
        /// </summary>
        public static string DataSource = "K";

        /// <summary>
        /// Data tables for comboBoxes.
        /// </summary>
        private static Dictionary<string, DataTable> DataTables = new Dictionary<string, DataTable>();

        /// <summary>
        /// Path for scripts.
        /// </summary>
        public static readonly string ScriptFolder = @"R scripts";

        /// <summary>
        /// Path for histogram script.
        /// </summary>
        public static string LinkHistogram = @"R scripts\Market Data\Histogram\histogram.r";

        /// <summary>
        /// Path for statistics script.
        /// </summary>
        public static string LinkStatistics = @"R scripts\Market Data\Curves\statistics.r";

        /// <summary>
        /// Path for statistics script.
        /// </summary>
        public static string LinkSaveSensitivityData = @"R scripts\Financial Risks\Sensitivity\sensitivity data saver.r";

        /// <summary>
        /// Path for statistics script.
        /// </summary>
        public static string LinkSaveTornadoDate = @"R scripts\Financial Risks\Tornado\tornado data saver.r";

        /// <summary>
        /// Path for item calculation script.
        /// </summary>
        public static string LinkItemsScript = @"R scripts\Forecasting\items.r";

        /// <summary>
        /// Path for item calculation GetValue function.
        /// </summary>
        public static string LinkGetValueScript = @"R scripts\Forecasting\getValue.r";

        /// <summary>
        /// Path for financial instruments script.
        /// </summary>
        public static string LinkFinancialInstruments = @"R scripts\Financial Instruments\";

        /// <summary>
        /// Path for financial instruments script.
        /// </summary>
        public static string LinkSpecialFunctions = @"R scripts\Functions\special functions.r";

        /// <summary>
        /// Path for item data load.
        /// </summary>
        public static string LinkItemDataLoad = @"R scripts\Forecasting\data loader.r";

        /// <summary>
        /// Path for item data save.
        /// </summary>
        public static string LinkItemDataSave = @"R scripts\Forecasting\data saver.r";

        /// <summary>
        /// Path for item scenario data save.
        /// </summary>
        public static string LinkItemDataScenarioSave = @"R scripts\Forecasting\data scenario saver.r";

        /// <summary>
        /// Path for market data save.
        /// </summary>
        public static string LinkMarketDataHistorySave = @"R scripts\Market data\Curves\market data saver.r";

        /// <summary>
        /// Path for histogram data save.
        /// </summary>
        public static string LinkHistogramDataSave = @"R scripts\Market data\Histogram\histogram data saver.r";

        /// <summary>
        /// Path for market data save.
        /// </summary>
        public static string LinkMarketDataScenarioSave = @"R scripts\Market data\Forecasting\scenario data saver.r";

        /// <summary>
        /// Path for index create script.
        /// </summary>
        public static string LinkIndexCreate = @"R scripts\Market data\Data Update\index create.r";

        /// <summary>
        /// Table for budget item scenario values.
        /// </summary>
        public static string BudgetItemScenarioTable = @"BudgetItemScenarioValues";

        /// <summary>
        /// Table for budget item history values.
        /// </summary>
        public static string BudgetItemHistoryTable = @"BudgetItemHistoryValues";

        /// <summary>
        /// Table for sensitivity budget item values.
        /// </summary>
        public static string SensitivityBudgetItemScenarioTable = @"SensitivityBudgetItemScenarioValues";

        /// <summary>
        /// Path for script to load correlations.
        /// </summary>
        public static string LinkGetCorrelationMatrix = @"R scripts\Market data\Joint forecasting\GetCorrelationMatrix.R";

        /// <summary>
        /// Path for script to load correlations.
        /// </summary>
        public static string LinkProductFlow = @"R scripts\Forecasting\product flows.R";

        /// <summary>
        /// Separator for series.
        /// </summary>
        public static string SeriesSeparator = @".";

        /// <summary>
        /// Separator for series.
        /// </summary>
        public static DateTime DefaultDate = DateTime.ParseExact("2000-01-01", "yyyy-MM-dd", CultureInfo.InvariantCulture);

        /// <summary>
        /// Default DateTime
        /// </summary>
        public static DateTime DefaultDateHistory = DateTime.ParseExact("2000-01-01", "yyyy-MM-dd", CultureInfo.InvariantCulture);

        /// <summary>
        /// Represents date format yyyy-MM-dd.
        /// </summary>
        public static string FormatDateYYYYMMDD = @"yyyy-MM-dd";

        /// <summary>
        /// Represents date format dd.MM.yyyy
        /// </summary>
        public static string FormatDateDDMMYYYY = @"dd.MM.yyyy";

        /// <summary>
        /// List of result types.
        /// </summary>
        public enum ResultTypes { Chart = 1, DescriptiveStatistics = 2, Map = 3, Table = 4 };

        /// <summary>
        /// List of stage types.
        /// </summary>
        public enum StageTypes { Independant = 1, Correlated = 2, Formula = 3 };

        /// <summary>
        /// List of value types.
        /// </summary>
        public enum ValueTypes { Absolute = 1, Log = 2 };

        /// <summary>
        /// List of axes.
        /// </summary>
        public enum Axes { X = 1, Y = 2 };

        /// <summary>
        /// Dictionary of sensitivity/tornado indexes. key:SeriesName, value:SeriesID
        /// </summary>
        public static Dictionary<string, int> sensitivityIndexes = new Dictionary<string, int>();

        /// <summary>
        /// Determines whether market data updates chart and statistics.
        /// </summary>
        public static bool IsChartUpdating = false;

        /// <summary>
        /// Determines ID of the current company.
        /// </summary>
        public static int CurrentCompanyID = -1;

        /// <summary>
        /// Determines parent ID of the current company.
        /// </summary>
        public static int CurrentParentID = 0;

        /// <summary>
        /// Determines ID of the current budget item.
        /// </summary>
        public static int CurrentBudgetItemID = -1;

        /// <summary>
        /// Determines ID of the current tree type.
        /// </summary>
        public static int CurrentTree = -1;

        /// <summary>
        /// Determines ID of the current user.
        /// </summary>
        public static int CurrentUserID = -1;

        /// <summary>
        /// Determines ID of the current project.
        /// </summary>
        public static int CurrentProjectID = -1;

        /// <summary>
        /// Contains list of companies for item forecasting.
        /// </summary>
        public static string CompanyList = String.Empty;


        public static List<int> CompanyIDsList = new List<int>();

        /// <summary>
        /// Determines length of history sample.
        /// </summary>
        public static int HistoryLength = 365;

        /// <summary>
        /// Random variable for colors
        /// </summary>
        public static Random RandomColor = new Random(10);

        /// <summary>
        /// Contains data 
        /// </summary>
        public static DataFrame CorrelationMatrix = null;

        /// <summary>
        /// Quantiles for portfolio optimization
        /// </summary>
        public static double[] PortfolioOptimizationQuantiles = new double[] { 0.05, 0.5, 0.95 };

        /// <summary>
        /// Dictionary of optimized set of instruments 
        /// </summary>
        public static Dictionary<Tuple<int/*Asset ID*/, int/*Type ID*/>, InstrumentEntity> po_p_assetIDDict = new Dictionary<Tuple<int, int>, InstrumentEntity>();

        /// <summary>
        /// Database model
        /// </summary>
        public static K_Model K_DataModel = new K_Model();

        /// <summary>
        /// Current User ID
        /// </summary>
        public static int K_UserID = 1; //1 - admin, 0 - not existed user

        /// <summary>
        /// Company filter condition string
        /// </summary>
        public static string CompanyFilter = "";

        public static string SeriesFamilyFilter = "";

        public static Dictionary<int, HashSet<int>> LinkedIndicesByIndex = new Dictionary<int, HashSet<int>>();
        
        /// <summary>
        /// Used in Financial Risks Tab
        /// </summary>

        public static Dictionary<string, List<int>> LinkedIndexIDsByItemCode = new Dictionary<string,List<int>>();

        /// <summary>
        /// Key1 - Item Code, Key2 - Asset type, Value - List of asset's ids
        /// </summary>
        public static Dictionary<string, Dictionary<string, List<int>>> AssetsByItemCode = new  Dictionary<string,Dictionary<string,List<int>>>();

        /// <summary>
        /// Key1 - Asset type, Key2 - AssetId, Value - List of index's ids
        /// </summary>
        public static Dictionary<string, Dictionary<int, HashSet<int>>> IndexIDsByAsset = new Dictionary<string, Dictionary<int, HashSet<int>>>();

        public static double[] Quantiles = new double[] { 0.01, 0.05, 0.5, 0.95, 0.99 };
        /// <summary>
        /// Current Language ID
        /// </summary>
        public static int LanguageID { get; set; }


        public struct OuterData
        {
            public SciChartSurface Chart;
            public List<string> Items;
            public ListView List;
            public GridView Grid;
            public string StartDateForecast;
            public string EndDateForecast;
            public string StartDateHistory;
            public string EndDateHistory;
            public DateTime StartDate;
            public DateTime EndDate;
            public string SeriesName;
            public string BudgetItemTable;
            public int FrequencyID;
            public int ValueType;
            public int ScenarioNumber;
            public int SeriesID;
            public int ModelID;
            public int History;
            public bool Estimate;
            public bool nullSigma;
            public bool IsSensitivity;
            public bool IsHistory;
            public string formula;
            public DataFrame Parameters;
        }

        public class InstrumentEntity
        {
            public int _ID { get; set; }
            public int _TypeID { get; set; }
            public int _CompanyID { get; set; }
            public int _BankID { get; set; }
            public int _GroupID { get; set; }
            public DateTime _OpenDate { get; set; }
            public DateTime? _CloseDate { get; set; }
            public string _CurrencyCode { get; set; }
            public string _Label { get; set; }
            public int _Duration { get; set; }

            public InstrumentEntity(int asset_id, int type_id, int comp_id, string curr_code, int bank_id, int group_id, DateTime start_date, DateTime? end_date, string label = "unNamed", int duration = 0)
            {
                this._ID = asset_id; this._TypeID = type_id;
                this._CompanyID = comp_id;
                this._CurrencyCode = curr_code;
                this._BankID = bank_id;
                this._GroupID = group_id;
                this._OpenDate = start_date;
                this._CloseDate = end_date;
                this._Label = label;
                this._Duration = duration;
            }
        }

        public static Dictionary<string,string> SortItemsByFormula(Dictionary<string,string> Items)
        {

            return null;
        }

        public static class FlowModel
        {
            public static DataFrame E;
            public static DataFrame R;
            public static DataFrame H;
            public static DataFrame L;
        }


        public static string MakeArrayR<T>(T[] Values)
        {
            return "c(" + MakeCommaSeparatedString(Values) + ")";
        }

        public static string MakeCommaSeparatedString<T>(T[] Values)
        {
            string IDs = "";
            if(typeof(T) != typeof(string))
            {
                foreach (var id in Values)
                {
                    IDs += id.ToString() + ",";
                }
            }
            else
            {
                foreach (var id in Values)
                {
                    IDs += "'" + id + "',";
                }

            }

            return IDs.TrimEnd(',');
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="tree">Tree to be filled in.</param>
        /// <param name="parentID">Parent ID of the company.</param>
        /// <param name="flow">Cashflow or company tree.</param>
        /// <param name="cashflow">Cashflow tree or Income tree.</param>
        /// <param name="IsLabel">Labels (true) or checkboxes (false).</param>
        [Obsolete("Use custom functions: CashFlowTreeBuild & CompanyTreeBuild methods.")]
        public static void TreeBuild(TreeView tree, int parentID, bool flow, int cashflow, bool IsLabel = false, RoutedPropertyChangedEventHandler<object> handler = null)
        {
            //if (handler == null)
            //{
            //    handler = delegate { };
            //}

            if (handler != null) tree.SelectedItemChanged -= handler;
            TreeViewItem[] items = new TreeViewItem[] { };
            tree.Items.Clear();
            if (handler != null) tree.SelectedItemChanged += handler;

            K_DataManager.CurrentTree = cashflow;

            string[] s;
            string query;
            int i = 0;

            using (SqlConnection conn = new SqlConnection(K_DataManager.ConnectionString))
            {
                conn.Open();

                if (flow)
                {
                    query = String.Format(@"select BudgetItemID as ID, BudgetItemName AS Name, Levels from BudgetItems where IsCashFlow = {0} and CompanyID = {1} order by LeftID", cashflow, K_DataManager.CurrentCompanyID);
                }
                else
                {
                    query = String.Format(@"select CompanyID as ID, CompanyName AS Name, Levels from Companies where ParentID = {0} order by LeftID", parentID);
                }

                SqlCommand cmd = new SqlCommand(query, conn);
                SqlDataReader r = cmd.ExecuteReader();

                while (r.Read())
                {
                    s = r["Levels"].ToString().Split('.');

                    if (s.Length > items.Length)
                    {
                        Array.Resize<TreeViewItem>(ref items, s.Length);
                    }

                    i = s.Length - 1;

                    if (s[i] != "")
                    {
                        if (IsLabel)
                        {
                            items[i] = new TreeViewItem() { Header = new Label() { Content = r["Name"] }, Tag = new Label() { Content = r["ID"] } };
                        }
                        else
                        {
                            items[i] = new TreeViewItem() { Header = new CheckBox() { Content = r["Name"] }, Tag = new Label() { Content = r["ID"] } };
                        }

                        if (i == 0)
                            tree.Items.Add(items[i]);
                        else
                            items[i - 1].Items.Add(items[i]);
                    }
                }
            }

            //if (cashflow == 0)
            try
            {
                ((tree.Items[0] as TreeViewItem).Items[0] as TreeViewItem).IsExpanded = false;
            }
            catch (Exception) { }
        }

        /// <summary>
        /// Создание дерева статей
        /// </summary>
        /// <param name="tree">Дерево для заполнения</param>
        /// <param name="cashflow">Индикатор денежного потока</param>
        public static void CashFlowTreeBuild(TreeView tree, int cashflow)
        {
            K_DataManager.CurrentTree = cashflow;
            var query = String.Format(@"select BudgetItemID as ID, BudgetItemName AS Name, Levels from BudgetItems where IsCashFlow = {0} and CompanyID = {1} order by LeftID", cashflow, K_DataManager.CurrentCompanyID);

            TreeBuildNew(tree, query);

            if (cashflow == 0)
            {
                ((tree.Items[0] as TreeViewItem).Items[0] as TreeViewItem).IsExpanded = false;
            }

        }

        /// <summary>
        /// Создание дерева компаний.
        /// </summary>
        /// <param name="tree">Дерево для заполнения</param>
        /// <param name="parentID">Идентификатор корня дерева</param>
        /// <param name="filter">Условие фильтрации компаний</param>
        public static void CompanyTreeBuild(TreeView tree, int parentID, string filter = "")
        {
            string condition = "";
            if (filter.Length > 0)
            {
                condition = " and CompanyID not in (" + filter + ") ";
            }

            var view = GetTableByName("Companies").DefaultView;
            view.RowFilter = "ParentID = " + parentID + condition;
            view.Sort = "LeftID";
            TreeBuildFromView(tree, view);

            //Collapse all subtrees of second level
            foreach (TreeViewItem sublevel in (tree.Items[0] as TreeViewItem).Items)
                sublevel.IsExpanded = false;
        }

        public static void TreeBuildFromView(TreeView tree, DataView view)
        {
            TreeViewItem[] items = new TreeViewItem[] { };
            tree.Items.Clear();

            string[] s;
            int i = 0;
            foreach (DataRowView r in view)
            {
                s = r.Row["Levels"].ToString().Split('.');

                if (s.Length > items.Length)
                {
                    Array.Resize<TreeViewItem>(ref items, s.Length);
                }

                i = s.Length - 1;

                if (s[i] != "")
                {
                    items[i] = new TreeViewItem() { Header = new Label() { Content = r.Row[1] }, Tag = new Label() { Content = r.Row[0] } };

                    if (i == 0)
                        tree.Items.Add(items[i]);
                    else if (items[i - 1] != null)
                        items[i - 1].Items.Add(items[i]);
                }
            }
        }

        public static void TreeBuildNew(TreeView tree, string query)
        {
            TreeViewItem[] items = new TreeViewItem[] { };
            tree.Items.Clear();

            string[] s;
            int i = 0;

            using (SqlConnection conn = new SqlConnection(K_DataManager.ConnectionString))
            {
                conn.Open();

                SqlCommand cmd = new SqlCommand(query, conn);
                SqlDataReader r = cmd.ExecuteReader();

                while (r.Read())
                {
                    s = r["Levels"].ToString().Split('.');

                    if (s.Length > items.Length)
                    {
                        Array.Resize<TreeViewItem>(ref items, s.Length);
                    }

                    i = s.Length - 1;

                    if (s[i] != "")
                    {
                        items[i] = new TreeViewItem() { Header = new Label() { Content = r["Name"] }, Tag = new Label() { Content = r["ID"] } };

                        if (i == 0)
                            tree.Items.Add(items[i]);
                        else if (items[i - 1] != null)
                            items[i - 1].Items.Add(items[i]);
                    }
                }
            }

        }


        //TODO: Вынести сохранение последних выбранных значений на выход из программы
        [Obsolete("Use FilterComboBox")]
        public static void FillComboBox(ComboBox comboBox, string member, string value, string table, int languageID = 1, string where = "")
        {
            comboBox.ItemsSource = null;
            comboBox.Items.Clear();

            if (!DataTables.ContainsKey(comboBox.Name))
            {
                DataTables.Add(comboBox.Name, new DataTable(table));
            }

            if (where == "" || where.Length > 0 && where.Trim()[where.Trim().Length - 1] != '=')
            {
                object temp = comboBox.Tag;
                comboBox.Tag = null;
                var dt = GetDataTableAsync(member, value, table, ref languageID, where);

                if (dt.Rows.Count > 0)
                {
                    AddTableToDictionary(dt);

                    comboBox.ItemsSource = dt.DefaultView;
                    comboBox.DisplayMemberPath = dt.Columns[member].ToString();
                    comboBox.SelectedValuePath = dt.Columns[value].ToString();
                }
                comboBox.Tag = temp;

                try
                {
                    if ((int)Properties.Settings.Default[comboBox.Name] == -1 && comboBox.Items.Count > 0)
                    {
                        comboBox.SelectedIndex = 0;
                    }
                    else
                    {
                        comboBox.SelectedIndex = (int)Properties.Settings.Default[comboBox.Name] - 1 >= comboBox.Items.Count ? comboBox.Items.Count - 1 : (int)Properties.Settings.Default[comboBox.Name];
                    }
                }
                catch (Exception)
                {
                    if (comboBox.Items.Count > 0)
                    {
                        comboBox.SelectedIndex = 0;
                    }
                    else
                    {
                        comboBox.SelectedIndex = -1;
                    }

                    try
                    {
                        SettingsProperty property = new SettingsProperty(comboBox.Name);
                        property.Attributes.Add(typeof(UserScopedSettingAttribute), new UserScopedSettingAttribute());
                        property.Provider = Properties.Settings.Default.Providers["LocalFileSettingsProvider"];
                        property.PropertyType = typeof(int);
                        property.SerializeAs = SettingsSerializeAs.String;
                        property.DefaultValue = -1;
                        Properties.Settings.Default.Properties.Add(property);
                        Properties.Settings.Default.Save();

                        comboBox.SelectedIndex = (int)Properties.Settings.Default[comboBox.Name];

                        if (comboBox.SelectedIndex == -1)
                        {
                            comboBox.SelectedIndex = 0;
                        }
                    }
                    catch (Exception)
                    { }
                }
            }
        }

        [Obsolete("Use TableView poxy")]
        private static DataTable GetDataTableAsync(string member, string value, string table, ref int languageID, string where)
        {
            if (where.StartsWith("AND", StringComparison.OrdinalIgnoreCase))
            {
                where = where.Substring(4);// Trim "AND" in the beginning of where clause
            }
            if (where.StartsWith(" AND", StringComparison.OrdinalIgnoreCase))
            {
                where = where.Substring(5);// Trim "AND" in the beginning of where clause
            }

            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                string query = String.Format("Select {0}, {1} from {2}", member, value, table);
                if (languageID == -1 && where.Length > 0)
                {
                    query += " where " + where;
                }
                else
                {
                    if (languageID != 1 && (int)K_DataManager.ExecuteSQLQuery(String.Format(@"select count(1) from {0} where LanguageID = {1}", table, languageID)) == 0)
                    {
                        languageID = 1;
                    }

                    query += " where LanguageID = " + languageID;
                    if (where.Length > 0)
                        query += " and " + where;
                }
                query += " order by 2";

                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                DataTable dt = new DataTable(table);
                dt.Clear();
                try
                {
                    da.Fill(dt);
                }
                catch (Exception exc)
                {
                    MessageBox.Show(caption: "Cann't load table: " + table, messageBoxText: exc.Message);
                }
                return dt;
            }
        }


        public static void FilterComboBox(ComboBox comboBox, string sourceTableName, string selectedMemberPath, string displayMemberPath, string orderByPath = "", int languageID = 1, string clause = "")
        {
            var view = GetTableView(sourceTableName);

            if (view == null)
            {
                return;
            }

            if (languageID == -1)
            {
                if (clause.Length > 0)
                    view.RowFilter = clause;
            }
            else if (clause.Length > 0)
            {
                view.RowFilter = clause + " AND LanguageID=" + languageID;
            }
            else
            {
                view.RowFilter = "LanguageID=" + languageID;
            }

            if (view.Count == 0)
            {
                view.RowFilter = clause;
            }

            if (orderByPath.Length > 0)
            {
                view.Sort = orderByPath;
            }
            else
            {
                int a = 0;
            }

            comboBox.ItemsSource = null;
            comboBox.ItemsSource = view;
            comboBox.DisplayMemberPath = displayMemberPath;
            comboBox.SelectedValuePath = selectedMemberPath;

            if (comboBox.Items.Count > 0)
            {
                comboBox.SelectedIndex = 0;
            }
        }

        public static void TreeBuildJoint(TreeView tree, object schemeID)
        {
            try
            {
                string query = string.Empty;

                tree.Items.Clear();

                using (SqlConnection conn = new SqlConnection(K_DataManager.ConnectionString))
                {
                    conn.Open();

                    query = String.Format(@"select seriesIDs+'-'+cast(stagetypeid as varchar)+'-'+SavedModelIDs AS Name, StageID as ID from stages
                                        where SchemeID = {0}
                                        order by StageID", schemeID);

                    SqlCommand cmd = new SqlCommand(query, conn);
                    SqlDataReader r = cmd.ExecuteReader();

                    while (r.Read())
                    {
                        tree.Items.Add(new TreeViewItem() { Header = new Label() { Content = r["Name"] }, Tag = new Label() { Content = r["ID"] } });
                    }
                }

                tree.Tag = 1;
            }
            catch (Exception) { }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="id"></param>
        /// <returns></returns>
        public static IEnumerable<K_Chart.VarPoint> GetQuantileData(int id, string date)
        {

            using (SqlConnection conn = new SqlConnection(K_DataManager.ConnectionString))
            {
                conn.Open();
                //SqlCommand cmd = new SqlCommand(String.Format(@"select Date, Value from BudgetItemHistoryValues where BudgetItemID = {0} order by Date", id), conn);
                SqlCommand cmd = new SqlCommand(String.Format(@"select Date, Value from BudgetItemHistoryValues where BudgetItemID = {0} and date >= '{1}' union select distinct Date, NULL Value from BudgetItemScenarioValues where BudgetItemID = {0} order by Date, Value DESC", id, date), conn);
                SqlDataReader r = cmd.ExecuteReader();

                K_DataManager.engine.Evaluate("library(RODBC)");
                K_DataManager.engine.Evaluate(String.Format(@"db_conn<-odbcConnect(""{0}"")", K_DataManager.DataSource));

                while (r.Read())
                {
                    double data5 = double.NaN;
                    double data25 = double.NaN;
                    double data50 = double.NaN;
                    double data75 = double.NaN;
                    double data95 = double.NaN;
                    double[] data;

                    #region comment
                    //try
                    //{
                    //    K_DataManager.engine.Evaluate(String.Format(@"q<-sqlQuery(db_conn,""select * from budgetitemscenariovalues where budgetItemID = {0} and date='{1}' order by Value asc"")", id, Convert.ToDateTime(r["Date"]).ToString("yyyy-MM-dd")));
                    //}
                    //catch (Exception)
                    //{
                    //    K_DataManager.engine.Evaluate("close(db_conn)");
                    //    K_DataManager.engine.Evaluate(@"db_conn<-odbcConnect(""K"")");
                    //}

                    //if (K_DataManager.engine.Evaluate("nrow(q)").AsNumeric()[0] > 0)
                    //{
                    //    data5 = K_DataManager.engine.Evaluate(@"quantile(q$Value, 0.05)").AsNumeric()[0];
                    //    data25 = K_DataManager.engine.Evaluate(@"quantile(q$Value, 0.25)").AsNumeric()[0];
                    //    data50 = K_DataManager.engine.Evaluate(@"quantile(q$Value, 0.50)").AsNumeric()[0];
                    //    data75 = K_DataManager.engine.Evaluate(@"quantile(q$Value, 0.75)").AsNumeric()[0];
                    //    data95 = K_DataManager.engine.Evaluate(@"quantile(q$Value, 0.95)").AsNumeric()[0];
                    //}
                    #endregion

                    data = K_Chart.GetQuantiles(String.Format(@"Select * from BudgetItemScenarioValues where BudgetItemID = {0} and date='{1}' order by Value asc", id, Convert.ToDateTime(r["Date"]).ToString("yyyy-MM-dd")), new double[] { 0.05, 0.95, 0.5, 0.25, 0.75 });

                    if (data != null)
                    {
                        data5 = data[0];
                        data95 = data[1];
                        data50 = data[2];
                        data25 = data[3];
                        data75 = data[4];
                    }

                    bool IsNaN = false;

                    try
                    {
                        double NaN = Convert.ToDouble(r["Value"]);
                    }
                    catch (Exception)
                    {
                        IsNaN = true;
                    }

                    if (!IsNaN)
                    {
                        yield return new K_Chart.VarPoint(Convert.ToDateTime(r["Date"]), Convert.ToDouble(r["Value"]), data5, data95, data50, data25, data75);
                    }
                    else
                    {
                        yield return new K_Chart.VarPoint(Convert.ToDateTime(r["Date"]), double.NaN, data5, data95, data50, data25, data75);
                    }
                }
            }
        }

        /// <summary>
        /// Executes scalar SQL query.
        /// </summary>
        /// <param name="query">SQL query string.</param>
        /// <returns>Returns one object.</returns>
        public static object ExecuteSQLQuery(string query)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(ConnectionString))
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(query, conn);
                    return cmd.ExecuteScalar();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(String.Format("Query: {0} Error: {1}", query, ex.Message), "Sql query FAILED!");
                return null;
            }
        }

        /// <summary>
        /// Executes scalar SQL query with parameters.
        /// </summary>
        /// <param name="query">SQL query string.</param>
        /// <param name="parameters">Parameters for the query</param>
        /// <returns>Returns one object.</returns>
        public static object ExecuteSQLQuery(string query, params object[] parameters)
        {
            using (SqlConnection conn = new SqlConnection(ConnectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(String.Format(query, parameters), conn);
                return cmd.ExecuteScalar();
            }
        }

        /// <summary>
        /// Searches for dependent items.
        /// </summary>
        /// <param name="equation">Item formula.</param>
        /// <returns>Returns list of item IDs.</returns>
        public static int[] GetDependentItems(string equation, string entity)
        {
            List<int> entities = new List<int>();

            Regex reg = new Regex(String.Format(@"({0}\.)(?<CompanyID>([0-9]*))\.(?<{0}>([a-zA-Z0-9_]*))", entity), RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace);
            MatchCollection Elements = reg.Matches(equation);

            foreach (Match element in Elements)
            {
                string code = element.Groups[entity].Value.ToString().ToLower();
                string companyID = element.Groups["CompanyID"].Value.ToString();
                var ID = K_DataManager.K_DataModel.BudgetItems.Where(b => b.CompanyID.ToString() == companyID && b.BudgetItemCode.ToLower() == code).Select(b => b.BudgetItemID).FirstOrDefault();
                entities.Add(ID);
            }

            return entities.ToArray();
        }

        public static string[] GetDependentIndex(string equation, string entity)
        {
            List<string> entities = new List<string>();

            //FI
            Regex reg = new Regex(String.Format(@"({0}\.)(?<{0}>([a-zA-Z0-9_]*))", entity), RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace);
            MatchCollection Elements = reg.Matches(equation);

            foreach (Match element in Elements)
            {
                entities.Add(element.Groups[entity].Value);
            }

            return entities.ToArray();
        }

        /// <summary>
        /// Method for release unused memory using GarbageCollector.
        /// </summary>
        public static void MemoryRelease()
        {
            if (K_DataManager.engine != null)
            {
                K_DataManager.engine.Evaluate("rm(list=ls(all=TRUE))");
                K_DataManager.engine.Evaluate("gc()");
                K_DataManager.engine.ClearGlobalEnvironment(true, true);
            }
        }

        /// <summary>
        /// Method for release unused memory using GarbageCollector.
        /// </summary>
        public static void ClearListGridVIew(ListView lv = null, GridView gv = null)
        {
            if (lv != null && lv.Items.Count > 0)
            {
                lv.Items.Clear();
            }

            if (gv != null && gv.Columns.Count > 0)
            {
                gv.Columns.Clear();
            }

            K_DataManager.MemoryRelease();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <param name="startDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static bool CheckDates(string startDate, string endDate)
        {
            if (startDate != "" && endDate != "")
            {
                if (Convert.ToDateTime(startDate) > Convert.ToDateTime(endDate))
                {
                    return false;
                }
            }

            if (startDate != "" && endDate == "")
            {
                if (Convert.ToDateTime(startDate) > DateTime.Now)
                {
                    return false;
                }
            }

            return true;
        }

        public static bool CheckNumber(object number)
        {
            double d = 0;
            return Double.TryParse(number.ToString(), out d);
        }

        public static Color GetNextColor()
        {
            var properties = typeof(Brushes).GetProperties();
            var count = properties.Count();

            var colour = properties
                        .Select(x => new { Property = x, Index = RandomColor.Next(count) })
                        .OrderBy(x => x.Index)
                        .First();

            return ((SolidColorBrush)colour.Property.GetValue(colour, null)).Color;
        }

        //TODO: DO NOT reload existing data. Hack - get data through R, get statistics and pass data to SciChart
        public static void UpdateDescriptiveStatistics(object _outerData)
        {
            if (!File.Exists(K_DataManager.LinkStatistics))
                return;

            K_DataManager.OuterData outerData = (K_DataManager.OuterData)_outerData;
            K_DataManager.ClearListGridVIew(outerData.List, outerData.Grid);

            K_DataManager.LinkStatistics = K_DataManager.LinkStatistics.Replace(@"\", @"/");

            var seriesIDs   = "params.IDs <- "      + MakeArrayR(K_Chart.md_c_seriesIDDict.Values.ToArray());
            var names       = "params.names <- "    + MakeArrayR(K_Chart.md_c_seriesIDDict.Keys.ToArray());

            try
            {
                #region Load series values
                K_DataManager.engine.Evaluate("if(!exists('generalDictionary')){generalDictionary<- list()}");
                foreach (var ID in K_Chart.md_c_seriesIDDict.Values)
                {
                    var dataValues = K_DataManager.K_DataModel.SeriesHistoryValues.Where(s => s.SeriesID == ID && s.FrequencyID == outerData.FrequencyID &&
                        s.Date >= outerData.StartDate && s.Date <= outerData.EndDate).OrderBy(s => s.Date).Select(s => s.Value).AsEnumerable();
                    var vecValues = K_DataManager.engine.CreateNumericVector(dataValues);
                    K_DataManager.engine.SetSymbol("vec", vecValues);
                    K_DataManager.engine.Evaluate(String.Format("generalDictionary[[{0}]] <- vec", ID));
                }
                #endregion

                K_DataManager.engine.Evaluate(String.Format(@"params.DataSource <-""{0}""", K_DataManager.DataSource));
                K_DataManager.engine.Evaluate(String.Format(@"params.where <- c(""{0}"", ""{1}"")", outerData.StartDate.ToString(FormatDateYYYYMMDD), outerData.EndDate.ToString(FormatDateYYYYMMDD)));
                K_DataManager.engine.Evaluate(String.Format(@"params.frequencyID <- {0}", outerData.FrequencyID));
                K_DataManager.engine.Evaluate(seriesIDs);
                K_DataManager.engine.Evaluate(names);
                K_DataManager.engine.Evaluate(String.Format(@"source('{0}')", K_DataManager.LinkStatistics));
                DataFrame data = K_DataManager.engine.Evaluate("stats").AsDataFrame();

                K_DataManager.MemoryRelease();

                outerData.Grid.Columns.Add(new GridViewColumn { Header = "Statistics", DisplayMemberBinding = new Binding("[0]") });

                for (int i = 0; i < data.ColumnCount; i++)
                {
                    outerData.Grid.Columns.Add(new GridViewColumn { Header = data.ColumnNames[i], DisplayMemberBinding = new Binding(String.Format("[{0}]", i+1)) });
                }

                for (int j = 0; j < data.RowCount; j++)
                {
                    Collection<object> items = new Collection<object>();
                    items.Add(data.RowNames[j]);

                    for (int i = 0; i < data.ColumnCount; i++)
                    {
                        items.Add(data[j, i]);
                    }

                    outerData.List.Items.Add(items);
                }

                foreach (var column in outerData.Grid.Columns)
                {
                    column.Width = Double.NaN;
                    column.ClearValue(GridViewColumn.WidthProperty);
                }
            }
            catch (Exception exc)
            {
                //MessageBox.Show(caption: "Can't get descriptive statistics via R.", messageBoxText: exc.Message);
            }
        }

        public static void ClearOuterData(ref OuterData outerData)
        {
            //K_DataManager.OuterData outerData = (K_DataManager.OuterData)_outerData;
            outerData.Chart = null;
            outerData.Items = null;
            outerData.List = null;
            outerData.Grid = null;
            outerData.StartDateForecast = String.Empty;
            outerData.EndDateForecast = String.Empty;
            outerData.StartDateHistory = String.Empty;
            outerData.EndDateHistory = String.Empty;
            outerData.SeriesName = String.Empty;
            K_DataManager.MemoryRelease();
        }

        public static void InitData()
        {
            //Company filter
            var companies = K_DataManager.K_DataModel.DisabledObjects.Where(obj => obj.UserID == 1 && obj.ObjectType == DisabledObject.Company.ToString()).Select(obj => obj.ObjectName);
            K_DataManager.CompanyFilter = String.Join(",", companies.ToArray());

            // SeriesFamily filter
            var seriesFamily = K_DataManager.K_DataModel.DisabledObjects.Where(obj => obj.UserID == 1 && obj.ObjectType == DisabledObject.SeriesFamily.ToString()).Select(obj => obj.ObjectName);
            K_DataManager.SeriesFamilyFilter = String.Join(",", seriesFamily.ToArray());

            try
            {
                K_DataManager.UpdateTableInDictionary("AggregationTypes");
                K_DataManager.UpdateTableInDictionary("AssetAccounts");
                K_DataManager.UpdateTableInDictionary("AssetTypes");
                K_DataManager.UpdateTableInDictionary("Axes");
                K_DataManager.UpdateTableInDictionary("AxisTypes");
                K_DataManager.UpdateTableInDictionary("Banks");
                K_DataManager.UpdateTableInDictionary("BankGroups");
                K_DataManager.UpdateTableInDictionary("BudgetItems");
                K_DataManager.UpdateTableInDictionary("Commodities");
                K_DataManager.UpdateTableInDictionary("Companies");
                K_DataManager.UpdateTableInDictionary("CorrelationMatrix");
                K_DataManager.UpdateTableInDictionary("Currencies");
                K_DataManager.UpdateTableInDictionary("ForecastModels");
                K_DataManager.UpdateTableInDictionary("Frequencies");

                K_DataManager.UpdateTableInDictionary("NetworkElements");
                K_DataManager.UpdateTableInDictionary("ResultTypes");
                K_DataManager.UpdateTableInDictionary("SavedModels");
                K_DataManager.UpdateTableInDictionary("Schemes");
                K_DataManager.UpdateTableInDictionary("Series");
                K_DataManager.UpdateTableInDictionary("SeriesFamilies");
                K_DataManager.UpdateTableInDictionary("SeriesTypes");
                K_DataManager.UpdateTableInDictionary("StageTypes");
                K_DataManager.UpdateTableInDictionary("StringDictionary");
                K_DataManager.UpdateTableInDictionary("TreeTypes");
                K_DataManager.UpdateTableInDictionary("Units");
                K_DataManager.UpdateTableInDictionary("ValueTypes");
                K_DataManager.UpdateTableInDictionary("ForecastTypes");
                K_DataManager.UpdateTableInDictionary("Stages");
            }
            catch (Exception exc)
            {
                MessageBox.Show(caption: "Main Init FAILED!!! Check DB connection", messageBoxText: (exc.InnerException ?? exc).Message);
                return;
            }
        }

        public static void AddTableToDictionary(DataTable Table)
        {
            if (!DataTables.ContainsKey(Table.TableName))
            {
                DataTables.Add(Table.TableName, Table);
            }
        }

        public static void UpdateTableInDictionary(string tableName)
        {
            if (DataTables.ContainsKey(tableName))
            {
                DataTables[tableName] = GetTableByName(tableName);
            }
            else
            {
                DataTables.Add(tableName, GetTableByName(tableName));
            }
        }

        public static DataView GetTableView(string TableName)
        {
            if (DataTables.ContainsKey(TableName))
            {
                return new DataView(DataTables[TableName]);
            }

            return null;
        }

        public static DataTable GetTableByName(string tableName)
        {
            switch (tableName)
            {
                default: MessageBox.Show(caption: "Add table mapping", messageBoxText: "Function: GetTableByName"); break;

                case "AggregationTypes": return ToDataTable(K_DataModel.AggregationTypes.ToList(), "AggregationTypes");
                case "AssetAccounts": return ToDataTable(K_DataModel.AssetAccounts.ToList(), "AssetAccounts");
                case "AssetTypes": return ToDataTable(K_DataModel.AssetTypes.ToList(), "AssetTypes");
                case "Axes": return ToDataTable(K_DataModel.Axes.ToList(), "Axes");
                case "AxisTypes": return ToDataTable(K_DataModel.AxisTypes.ToList(), "AxisTypes");
                case "Banks": return ToDataTable(K_DataModel.Banks.ToList(), "Banks");
                case "BankGroups": return ToDataTable(K_DataModel.BankGroups.ToList(), "BankGroups");
                case "BudgetItems": return ToDataTable(K_DataModel.BudgetItems.ToList(), "BudgetItems");
                case "Commodities": return ToDataTable(K_DataModel.Commodities.ToList(), "Commodities");
                case "CommodityExtractions": return ToDataTable(K_DataModel.CommodityExtractions.ToList(), "CommodityExtractions");
                case "CommodityHubs": return ToDataTable(K_DataModel.CommodityHubs.ToList(), "CommodityHubs");
                case "CommodityRefineries": return ToDataTable(K_DataModel.CommodityRefineries.ToList(), "CommodityRefineries");
                case "CommodityStorages": return ToDataTable(K_DataModel.CommodityStorages.ToList(), "CommodityStorages");
                case "Companies": return ToDataTable(K_DataModel.Companies.ToList(), "Companies");
                case "CorrelationMatrix": return ToDataTable(K_DataModel.CorrelationMatrices.ToList(), "CorrelationMatrix");
                case "Currencies": return ToDataTable(K_DataModel.Currencies.ToList(), "Currencies");
                case "ForecastModels": return ToDataTable( K_DataModel.ForecastModels.ToList(), "ForecastModels");
                case "Frequencies": return ToDataTable( K_DataModel.Frequencies.ToList(), "Frequencies");
                case "NetworkEdges": return ToDataTable(K_DataModel.NetworkEdges.ToList(), "NetworkEdges");
                case "NetworkElements": return ToDataTable(K_DataModel.NetworkElements.ToList(), "NetworkElements");
                case "ResultTypes": return ToDataTable(K_DataModel.ResultTypes.ToList(), "ResultTypes");
                case "SavedModels": return ToDataTable(K_DataModel.SavedModels.ToList(), "SavedModels");
                case "Schemes": return ToDataTable(K_DataModel.Schemes.ToList(), "Schemes");
                case "Series": return ToDataTable(K_DataModel.Series.ToList(), "Series");
                case "SeriesFamilies": return ToDataTable(K_DataModel.SeriesFamilies.ToList(), "SeriesFamilies");
                case "SeriesTypes": return ToDataTable( K_DataModel.SeriesTypes.ToList(), "SeriesTypes");
                case "StageTypes": return ToDataTable(K_DataModel.StageTypes.ToList(), "StageTypes");
                case "StringDictionary": return ToDataTable(K_DataModel.StringDictionary.ToList(), "StringDictionary");
                case "TreeTypes": return ToDataTable(K_DataModel.TreeTypes.ToList(), "TreeTypes");
                case "Units": return ToDataTable(K_DataModel.Units.ToList(), "Units");
                case "ValueTypes": return ToDataTable(K_DataModel.ValueTypes.ToList(), "ValueTypes");
                case "BudgetItemQuantiles": return ToDataTable(K_DataModel.BudgetItemQuantiles.ToList(), "BudgetItemQuantiles");
                case "ForecastTypes": return ToDataTable(K_DataModel.ForecastTypes.ToList(), "ForecastTypes");
                case "Stages": return ToDataTable(K_DataModel.Stages.ToList(), "Stages");
            }
            return null;
        }
        
        public static bool IsCorrMatrixConsistent(string seriesIDs, string corrID)
        {
            bool conformable = true;

            using (SqlConnection conn = new SqlConnection(K_DataManager.ConnectionString))
            {
                conn.Open();
                string query = String.Format(@"select distinct(seriesid1) as a from correlations
                                                            where correlationid={0}
                                                            union 
                                                            select distinct(seriesid2) as a from correlations
                                                            where correlationid={0}", corrID);

                SqlCommand cmd = new SqlCommand(query, conn);
                SqlDataReader r = cmd.ExecuteReader();

                while (r.Read() && conformable)
                {
                    if (!seriesIDs.Split(',').Contains(r["a"].ToString())) conformable = false;
                }
            }

            return conformable;
        }
        public static void GetFlowModelResult( int CompanyID)
        {
            var RLink = LinkProductFlow;
            if (File.Exists(String.Format("{0}", RLink)))
            {
                RLink = RLink.Replace(@"\", @"/");
                try
                {
                    K_DataManager.engine.Evaluate(String.Format(@"DataSource <- ""{0}""", K_DataManager.DataSource));
                    K_DataManager.engine.Evaluate(String.Format(@"companyID <- {0}", CompanyID));
                    K_DataManager.engine.Evaluate(String.Format(@"source('{0}')", RLink));

                    var E = K_DataManager.engine.Evaluate("na.omit(E)").AsDataFrame();
                    if (E.RowCount > 0) FlowModel.E = E;

                    var R = K_DataManager.engine.Evaluate("na.omit(R)").AsDataFrame();
                    if (R.RowCount > 0) FlowModel.R = R;

                    var H = K_DataManager.engine.Evaluate("na.omit(D)").AsDataFrame();
                    if (H.RowCount > 0) FlowModel.H = H;

                    var L = K_DataManager.engine.Evaluate("na.omit(L)").AsDataFrame();
                    if (L.RowCount > 0) FlowModel.L = L;

                }
                catch (Exception exc)
                {
                }
            }
        }

        public static int InsertUpdateHistoryForIndex(int seriesID, string RLink, string RParameter, string series, DateTime startDate, bool synthetic = false)
        {
            int rows = 0;
            bool evaluated = false;
            int ID = (int)K_DataManager.ExecuteSQLQuery("select isnull(max(ID),0) from logs") + 1;
            int baseFreqID = 1;

            try
            {
                K_DataManager.ExecuteSQLQuery(String.Format("Delete SeriesHistoryValues where seriesID = {0} and Date > '{1}'", seriesID, startDate.ToString("yyyy-MM-dd")));

                if (File.Exists(String.Format("{0}", RLink)))
                {
                    RLink = RLink.Replace(@"\", @"/");

                    try
                    {
                        K_DataManager.engine.Evaluate(String.Format(@"DataSource <- ""{0}""", K_DataManager.DataSource));
                        K_DataManager.engine.Evaluate(String.Format(@"start_date <- ""{0}""", startDate.ToString(K_DataManager.FormatDateDDMMYYYY)));
                        K_DataManager.engine.Evaluate(String.Format(@"seriesID <- {0}", seriesID));
                        K_DataManager.engine.Evaluate(String.Format(@"item <- ""{0}""", RParameter));
                        K_DataManager.engine.Evaluate(String.Format(@"login <- """""));
                        K_DataManager.engine.Evaluate(String.Format(@"pwd <- """""));
                        K_DataManager.engine.Evaluate(String.Format(@"source('{0}')", RLink));

                        var res = K_DataManager.engine.Evaluate("na.omit(res)").AsDataFrame();
                        if (res.RowCount > 0)
                        {
                            DataTable table = new DataTable("SeriesHistoryValues");
                            RDataFrameToDataTable(res, table);
                            DataTableBulkCopy(table);
                        }

                        rows = res.RowCount;
                        evaluated = true;

                        baseFreqID = (int)K_DataManager.engine.Evaluate("min(res$FrequencyID)").AsNumeric()[0];
                    }
                    catch (Exception exc)
                    {
                        MessageBox.Show(caption: "Can't get history series via R.", messageBoxText: exc.Message);
                    }
                }
                else
                {
                    if (synthetic)
                    {
                        K_DataManager.ExecuteSQLQuery(String.Format("Insert into Logs Select {0}, N'Series ''{1}'': script not found', N'{2}','{3}'", ID++, series, "Try evaluation", DateTime.Now));
                    }

                    if (File.Exists(K_DataManager.LinkIndexCreate))
                    {
                        try
                        {
                            K_DataManager.engine.Evaluate(String.Format(@"start_date <- ""{0}""", startDate.ToString("dd.MM.yyyy")));
                            K_DataManager.engine.Evaluate(String.Format(@"seriesID <- {0}", seriesID));
                            K_DataManager.engine.Evaluate(String.Format(@"DataSource <- ""{0}""", K_DataManager.DataSource));
                            K_DataManager.engine.Evaluate(String.Format(@"source('{0}')", K_DataManager.LinkIndexCreate.Replace(@"\", @"/")));

                            var res = K_DataManager.engine.Evaluate("na.omit(res)").AsDataFrame();
                            if (res != null && res.RowCount > 0)
                            {
                                DataTable table = new DataTable("SeriesHistoryValues");
                                RDataFrameToDataTable(res, table);
                                DataTableBulkCopy(table);
                                rows = res.RowCount;
                                baseFreqID = (int)K_DataManager.engine.Evaluate("min(res$FrequencyID)").AsNumeric()[0];
                                evaluated = true;
                            }
                            else
                            {
                                rows = 0;
                            }

                        }
                        catch (Exception exc)
                        {
                            MessageBox.Show(caption: "Can't calculate history series via R.", messageBoxText: exc.Message);
                        }
                    }
                }

                if (evaluated)
                {
                    if (rows > 0)
                    {
                        for (int freqID = baseFreqID + 1; freqID <= 5; freqID++) //week=2..year=5
                        {
                            K_DataManager.InsertHistoryFrequencyValues(seriesID, freqID, baseFreqID);
                        }
                    }
                }
            }
            catch (Exception exc)
            {
                MessageBox.Show(caption: "Can't update history series.", messageBoxText: exc.Message);
            }
            finally
            {
                K_DataManager.MemoryRelease();
            }

            return rows;
        }

        public static Task UpdateMarketData(/*object loader*/)
        {
            DateTime startDate;
            string seriesID;
            string series;
            string RLink;
            string RParameter;
            string log = "";
            int rows = 0;

            return Task.Run(() =>
            {
                //AdornedControl.AdornedControl Loader = (AdornedControl.AdornedControl)loader;
                K_DataManager.ExecuteSQLQuery("Delete Logs");

                using (SqlConnection conn = new SqlConnection(K_DataManager.ConnectionString))
                {
                    try
                    {
                        //Loader.IsAdornerVisible = true;

                        conn.Open();
                        SqlCommand cmd = new SqlCommand(@"select s.seriesID, s.seriesName, s.RLink, s.RParameter, isnull(max(v.Date),'2000-01-01') lastDate from series s
                                                    left join seriesHistoryValues v
                                                    on s.seriesID=v.seriesID
                                                    where LanguageID = 1
                                                    group by s.seriesID, s.seriesName, s.RLink, s.RParameter
                                                    order by 1", conn);

                        SqlDataReader r = cmd.ExecuteReader();

                        while (r.Read())
                        {
                            startDate = Convert.ToDateTime(r["lastDate"]).AddDays(1);
                            seriesID = r["seriesID"].ToString();
                            series = r["seriesName"].ToString();
                            RLink = r["RLink"].ToString();
                            RParameter = r["RParameter"].ToString();

                            if (startDate > DateTime.Now || (RLink == "" && RParameter == ""))
                            {
                                continue;
                            }

                            rows += InsertUpdateHistoryForIndex(Convert.ToInt32(seriesID), RLink, RParameter, series, startDate, (RParameter != "" && RLink == ""));
                        }

                        log = String.Format("Market data updated!\r\nAdded {0} values.\r\nSee log for details.", rows);

                        MessageBox.Show(log, "Information!", MessageBoxButton.OK, MessageBoxImage.Information);
                    }
                    catch (Exception e) { }
                    finally
                    {
                        //Loader.IsAdornerVisible = false;
                    }
                }

                K_DataManager.ExecuteSQLQuery("delete SeriesHistoryValues where SeriesID = 0");
                K_DataManager.ExecuteSQLQuery(@"WITH CTE AS(SELECT seriesid, frequencyid, date,value,
                                            RN = ROW_NUMBER()OVER(PARTITION BY SeriesID, frequencyid, date ORDER BY SeriesID,frequencyid, date)  FROM serieshistoryvalues)
                                            DELETE FROM CTE WHERE RN > 1");
            });
        }

        public static void InsertHistoryFrequencyValues(int seriesID, int frequencyID, int baseFreqID = 1)
        {
            DateTime minDate = DateTime.Now;
            string[] frequencies = new string[] { "0", "day", "week", "month", "quarter", "year" };

            using (SqlConnection conn = new SqlConnection(K_DataManager.ConnectionString))
            {
                try
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand(String.Format(@"select max(date) from serieshistoryvalues
                                                                    where seriesID={0} and frequencyid={1}
                                                                    group by seriesid", seriesID, frequencyID), conn);

                    try
                    {
                        minDate = Convert.ToDateTime(cmd.ExecuteScalar());

                        if (minDate.ToString("dd.MM.yyyy") == "01.01.0001")
                        {
                            cmd.CommandText = String.Format(@"select min(date) from serieshistoryvalues
                                                              where seriesID={0} and frequencyid={1}
                                                              group by seriesid", seriesID, baseFreqID);
                            minDate = Convert.ToDateTime(cmd.ExecuteScalar());
                        }
                    }
                    catch (Exception)
                    {
                        minDate = DateTime.Now.AddYears(-20);
                    }

                    cmd.CommandText = string.Format(@"DELETE serieshistoryvalues WHERE seriesID = {0} and frequencyid = {1} and date>='{2}'", seriesID, frequencyID, minDate.ToString("yyyy-MM-dd"));
                    cmd.ExecuteNonQuery();

                    cmd.CommandText = String.Format(@"INSERT INTO serieshistoryvalues
                                                    SELECT {0}, {1} FrequencyID, CONVERT(DATE,DATEADD({3}, DATEDIFF({3},0,date), 0),101), avg(value), 0 IsAuthentic FROM serieshistoryvalues
                                                    WHERE seriesID={0} and frequencyid = {4} and date>='{2}'
                                                    GROUP BY seriesID, CONVERT(DATE,DATEADD({3}, DATEDIFF({3},0,date), 0),101)
                                                    ORDER BY 1,3", seriesID, frequencyID, minDate.ToString("yyyy-MM-dd"), frequencies[frequencyID], baseFreqID);
                    cmd.ExecuteNonQuery();
                    cmd.Dispose();
                }
                catch (Exception) { }
            }
        }
        private static void CastDataTypeByTable(DataTable dt)
        {
            if (dt.TableName == "SeriesHistoryValues")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.Int32");
                dt.Columns[1].DataType = System.Type.GetType("System.Int32");
                dt.Columns[2].DataType = System.Type.GetType("System.DateTime");
                dt.Columns[3].DataType = System.Type.GetType("System.Double");
                dt.Columns[4].DataType = System.Type.GetType("System.Int32");
            }
            else if (dt.TableName == "AssetHistoryValues")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.Int32");
                dt.Columns[1].DataType = System.Type.GetType("System.Int32");
                dt.Columns[2].DataType = System.Type.GetType("System.Int32");
                dt.Columns[3].DataType = System.Type.GetType("System.DateTime");
                dt.Columns[4].DataType = System.Type.GetType("System.Double");
            }
            else if (dt.TableName == "AssetScenarioValues")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.Int32");
                dt.Columns[1].DataType = System.Type.GetType("System.Int32");
                dt.Columns[2].DataType = System.Type.GetType("System.Int32");
                dt.Columns[3].DataType = System.Type.GetType("System.DateTime");
                dt.Columns[4].DataType = System.Type.GetType("System.Double");
                dt.Columns[5].DataType = System.Type.GetType("System.Int32");
                dt.Columns[6].DataType = System.Type.GetType("System.Int32");
            }
            else if (dt.TableName == "SeriesScenarioValues")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.Int32");
                dt.Columns[1].DataType = System.Type.GetType("System.Int32");
                dt.Columns[2].DataType = System.Type.GetType("System.DateTime");
                dt.Columns[3].DataType = System.Type.GetType("System.Int32");
                dt.Columns[4].DataType = System.Type.GetType("System.Double");
                dt.Columns[5].DataType = System.Type.GetType("System.Double");
            }
            else if (dt.TableName == "SeriesQuantiles")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.Int32");
                dt.Columns[1].DataType = System.Type.GetType("System.Int32");
                dt.Columns[2].DataType = System.Type.GetType("System.DateTime");
                dt.Columns[3].DataType = System.Type.GetType("System.Double");
                dt.Columns[4].DataType = System.Type.GetType("System.Double");
            }
            else if (dt.TableName == "BudgetItemHistoryValues")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.Int32");
                dt.Columns[1].DataType = System.Type.GetType("System.Int32");
                dt.Columns[2].DataType = System.Type.GetType("System.DateTime");
                dt.Columns[3].DataType = System.Type.GetType("System.Double");
            }
            else if (dt.TableName == "BudgetItemScenarioValues")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.Int32");
                dt.Columns[1].DataType = System.Type.GetType("System.Int32");
                dt.Columns[2].DataType = System.Type.GetType("System.DateTime");
                dt.Columns[3].DataType = System.Type.GetType("System.Int32");
                dt.Columns[4].DataType = System.Type.GetType("System.Double");
            }
            else if (dt.TableName == "BudgetItemQuantiles")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.Int32");
                dt.Columns[1].DataType = System.Type.GetType("System.DateTime");
                dt.Columns[2].DataType = System.Type.GetType("System.Double");
                dt.Columns[3].DataType = System.Type.GetType("System.Double");
                dt.Columns[4].DataType = System.Type.GetType("System.Double");
                dt.Columns[5].DataType = System.Type.GetType("System.Double");
                dt.Columns[6].DataType = System.Type.GetType("System.Double");
            }
            else if (dt.TableName == "Optimization")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.DateTime");
                for (int i = 1; i < dt.Columns.Count; i++)
                {
                    dt.Columns[i].DataType = System.Type.GetType("System.Double");
                }
            }
            else if (dt.TableName == "SensitivityBudgetItemScenarioValues")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.Int32");
                dt.Columns[1].DataType = System.Type.GetType("System.Int32");
                dt.Columns[2].DataType = System.Type.GetType("System.DateTime");
                dt.Columns[3].DataType = System.Type.GetType("System.Int32");
                dt.Columns[4].DataType = System.Type.GetType("System.Double");
            }
            else if (dt.TableName == "Sensitivity")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.Int32");
                dt.Columns[1].DataType = System.Type.GetType("System.Int32");
                dt.Columns[2].DataType = System.Type.GetType("System.DateTime");
                dt.Columns[3].DataType = System.Type.GetType("System.Double");
                dt.Columns[4].DataType = System.Type.GetType("System.Double");
                dt.Columns[5].DataType = System.Type.GetType("System.Double");
            }
            else if (dt.TableName == "Tornado")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.Int32");
                dt.Columns[1].DataType = System.Type.GetType("System.Int32");
                dt.Columns[2].DataType = System.Type.GetType("System.DateTime");
                dt.Columns[3].DataType = System.Type.GetType("System.Double");
                dt.Columns[4].DataType = System.Type.GetType("System.Double");
                dt.Columns[5].DataType = System.Type.GetType("System.Double");
                dt.Columns[6].DataType = System.Type.GetType("System.Double");
                dt.Columns[7].DataType = System.Type.GetType("System.Double");
            }
            else if (dt.TableName == "Series")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.Int32");
                dt.Columns[1].DataType = System.Type.GetType("System.Int32");
                dt.Columns[2].DataType = System.Type.GetType("System.String");
                dt.Columns[3].DataType = System.Type.GetType("System.String");
                dt.Columns[4].DataType = System.Type.GetType("System.Int32");
                dt.Columns[5].DataType = System.Type.GetType("System.String");
                dt.Columns[6].DataType = System.Type.GetType("System.String");
                dt.Columns[7].DataType = System.Type.GetType("System.String");
                dt.Columns[8].DataType = System.Type.GetType("System.String");
                dt.Columns[9].DataType = System.Type.GetType("System.Int32");
                dt.Columns[10].DataType = System.Type.GetType("System.String");
                dt.Columns[11].DataType = System.Type.GetType("System.Int32");
            }
            else if (dt.TableName == "BudgetItems")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.Int32");
                dt.Columns[1].DataType = System.Type.GetType("System.String");
                dt.Columns[2].DataType = System.Type.GetType("System.String");
                dt.Columns[3].DataType = System.Type.GetType("System.Int32");
                dt.Columns[4].DataType = System.Type.GetType("System.Int32");
                dt.Columns[5].DataType = System.Type.GetType("System.String");
                dt.Columns[6].DataType = System.Type.GetType("System.Int32");
                dt.Columns[7].DataType = System.Type.GetType("System.Int32");
                dt.Columns[8].DataType = System.Type.GetType("System.String");
                dt.Columns[9].DataType = System.Type.GetType("System.Int32");
                dt.Columns[10].DataType = System.Type.GetType("System.Int32");
                dt.Columns[11].DataType = System.Type.GetType("System.Int32");
            }
            else if (dt.TableName == "Statistics")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.String");
            }
            else if (dt.TableName == "BudgetItemsData")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.DateTime");
                for (int i = 1; i < dt.Columns.Count; i++)
                {
                    dt.Columns[i].DataType = System.Type.GetType("System.Double");
                }
            }
            else if (dt.TableName == "Correlations")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.Int32");
                dt.Columns[1].DataType = System.Type.GetType("System.Int32");
                dt.Columns[2].DataType = System.Type.GetType("System.Int32");
                dt.Columns[3].DataType = System.Type.GetType("System.Double");
                dt.Columns[4].DataType = System.Type.GetType("System.Int32");
                dt.Columns[5].DataType = System.Type.GetType("System.Int32");
            }
            else if (dt.TableName == "choices")
            {
                dt.Columns[0].DataType = System.Type.GetType("System.Double");
                dt.Columns[1].DataType = System.Type.GetType("System.String");
                dt.Columns[2].DataType = System.Type.GetType("System.DateTime");
                dt.Columns[3].DataType = System.Type.GetType("System.DateTime");
                dt.Columns[4].DataType = System.Type.GetType("System.Double");
                dt.Columns[5].DataType = System.Type.GetType("System.Double");
            }

            
        }

        public static DataTable RDataFrameToDataTable(DataFrame resultsMatrix, DataTable dt)
        {
            var columns = new DataColumn[resultsMatrix.ColumnCount];

            for (int i = 0; i < resultsMatrix.ColumnCount; i++)
            {
                columns[i] = new DataColumn(resultsMatrix.ColumnNames[i], typeof(double));
            }

            #region TODO: Add dynamic type cast
            dt.Columns.AddRange(columns);
            CastDataTypeByTable(dt);
            #endregion
            for (int y = 0; y < resultsMatrix.RowCount; y++)
            {
                var dr = dt.NewRow();
                for (int x = 0; x < resultsMatrix.ColumnCount; x++)
                {
                    try
                    {
                        dr[x] = (resultsMatrix[y, x].ToString() == "NA" ? DBNull.Value : resultsMatrix[y, x]) ?? DBNull.Value;
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, String.Format("Column {0} binding error!", resultsMatrix.ColumnNames[x]), MessageBoxButton.OK, MessageBoxImage.Error);
                        throw;
                    }
                }

                dt.Rows.Add(dr);
            }

            return dt;
        }

        public static List<BudgetItemScenarioValue> RDataFrameToEntities(DataFrame resultsMatrix)
        {
            var result = new List<BudgetItemScenarioValue>(resultsMatrix.RowCount);
            for (int y = 0; y < resultsMatrix.RowCount; y++)
            {
                var line = new BudgetItemScenarioValue();
                try
                {
                    line.BudgetItemID       = Int32.Parse(resultsMatrix[y, 0].ToString());
                    line.FrequencyID        = Int32.Parse(resultsMatrix[y, 1].ToString());
                    line.Date               = DateTime.Parse(resultsMatrix[y, 2].ToString());
                    line.ScenarioNumber     = Int32.Parse(resultsMatrix[y, 3].ToString());
                    line.Value              = Double.Parse(resultsMatrix[y, 4].ToString());
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, String.Format("Line {0} parse error!", y), MessageBoxButton.OK, MessageBoxImage.Error);
                }
                result.Add(line);
            }

            return result;
        }

        public static void DataTableBulkCopy(DataTable table)
        {
            using (var bulkCopy = new SqlBulkCopy(K_DataManager.ConnectionString, SqlBulkCopyOptions.KeepIdentity))
            {
                foreach (DataColumn col in table.Columns)
                {
                    bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                }

                bulkCopy.BulkCopyTimeout = 600;
                bulkCopy.DestinationTableName = table.TableName;
                bulkCopy.WriteToServer(table);
            }
        }

        async public static void DataTableBulkCopyAsync(DataTable table)
        {
            using (var bulkCopy = new SqlBulkCopy(K_DataManager.ConnectionString, SqlBulkCopyOptions.KeepIdentity))
            {
                foreach (DataColumn col in table.Columns)
                {
                    bulkCopy.ColumnMappings.Add(col.ColumnName, col.ColumnName);
                }

                bulkCopy.BulkCopyTimeout = 600;
                bulkCopy.DestinationTableName = table.TableName;
                await bulkCopy.WriteToServerAsync(table);
            }
        }



        public static void CalculateFinancialInstruments(int ID, string assetName, DateTime forecastStartDate, int scenarioNumber = 1, int freqID = 1)
        {
            try
            {
                string link = K_DataManager.LinkFinancialInstruments + assetName.ToLower() + ".r";

                if (File.Exists(String.Format("{0}", link)))
                {
                    link = link.Replace(@"\", @"/");

                    var assetTypeID = K_DataModel.AssetTypes.Where(a => a.AssetTypeName == assetName && a.LanguageID == 1).Select(a => a.AssetTypeID).SingleOrDefault();

                    K_DataManager.K_DataModel.Database.ExecuteSqlCommand(String.Format("Delete from AssetHistoryValues where AssetTypeID = {0} and AssetID = {1}", assetTypeID, ID));
                    K_DataManager.K_DataModel.Database.ExecuteSqlCommand(String.Format("Delete from AssetScenarioValues where AssetTypeID = {0} and AssetID = {1}", assetTypeID, ID));
                    #region Preload indecies
                    var indexIDs = K_DataManager.GetIndexIDsByAsset(assetName, ID);
                    // History part

                    // Forecast part
                    #endregion
                    #region Load assetAccounts

                    var colNames = new string[] { "AssetAccountCode", "AssetAccountID" };
                    var listAccounts = K_DataManager.K_DataModel.AssetAccounts.Where(a => a.LanguageID == 1);
                    var data = new IEnumerable[2];
                    var ids = new List<int>();
                    var codes = new List<string>();
                    listAccounts.ToList().ForEach(l => { ids.Add(l.AssetAccountID); codes.Add(l.AssetAccountCode); });
                    data[0] = codes;
                    data[1] = ids;

                    var df = K_DataManager.engine.CreateDataFrame(data, columnNames: colNames);
                    K_DataManager.engine.SetSymbol("assetaccounts", df);
                    K_DataManager.engine.Evaluate("if(!exists('generalDictionary')){generalDictionary<- list()}");
                    K_DataManager.engine.Evaluate(String.Format("generalDictionary[['{0}']] <- {0}", "assetaccounts"));
                    #endregion

                    K_DataManager.engine.Evaluate(String.Format(@"DataSource <- ""{0}""", K_DataManager.DataSource));
                    K_DataManager.engine.Evaluate(String.Format(@"{0}ID <- {1}", assetName.ToLower(), ID));
                    K_DataManager.engine.Evaluate(String.Format(@"forecast_start_date <- as.Date(""{0}"")", forecastStartDate.ToString(K_DataManager.FormatDateYYYYMMDD)));
                    K_DataManager.engine.Evaluate(String.Format(@"ir_year <- {0}", 365));
                    K_DataManager.engine.Evaluate(String.Format(@"frequencyID <- {0}", freqID));
                    K_DataManager.engine.Evaluate(String.Format(@"mc_num <- {0}", scenarioNumber));
                    K_DataManager.engine.Evaluate(String.Format(@"source('{0}',chdir=T)", link));

                    var history_res = K_DataManager.engine.Evaluate("history_res").AsDataFrame();
                    if (history_res.ColumnCount > 0)
                    {
                        DataTable htable = new DataTable("AssetHistoryValues");
                        RDataFrameToDataTable(history_res, htable);
                        DataTableBulkCopy(htable);
                    }

                    var forecast_res = K_DataManager.engine.Evaluate("forecast_res").AsDataFrame();
                    if (forecast_res.ColumnCount > 0)
                    {
                        DataTable ftable = new DataTable("AssetScenarioValues");
                        RDataFrameToDataTable(forecast_res, ftable);
                        DataTableBulkCopy(ftable);
                    }

                    K_DataManager.MemoryRelease();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, String.Format("Instrument {0} not evaluated!", ID), MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }

        public static void ForecastIndex(OuterData outerData)
        {
            K_DataManager.LoadRFunctions();
            try
            {
                string link = K_DataManager.ExecuteSQLQuery(String.Format("select Link from forecastmodels where modelID = {0} and languageID = 1", outerData.ModelID)).ToString();
                string type = K_DataManager.ExecuteSQLQuery(String.Format("select Parameters from forecastmodels where modelID = {0} and languageID = 1", outerData.ModelID)).ToString();

                if (File.Exists(String.Format("{0}", link)))
                {
                    link = link.Replace(@"\", @"/");

                    K_DataManager.ExecuteSQLQuery(String.Format("delete SeriesScenarioValues where SeriesID = {0}", outerData.SeriesID));

                    K_DataManager.engine.Evaluate(String.Format(@"input.DataSource <- ""{0}""", K_DataManager.DataSource));
                    K_DataManager.engine.Evaluate(String.Format(@"input.forecast_start_date <- as.Date(""{0}"")", Convert.ToDateTime(outerData.StartDateForecast).ToString("yyyy-MM-dd")));
                    K_DataManager.engine.Evaluate(String.Format(@"input.forecast_end_date <- as.Date(""{0}"")", Convert.ToDateTime(outerData.EndDateForecast).ToString("yyyy-MM-dd")));
                    K_DataManager.engine.Evaluate(String.Format(@"input.estimate<-{0}", outerData.Estimate.ToString().ToUpper()));
                    K_DataManager.engine.Evaluate(String.Format(@"input.seriesID <- {0}", outerData.SeriesID));
                    K_DataManager.engine.Evaluate(String.Format(@"input.frequencyID <- {0}", outerData.FrequencyID));
                    K_DataManager.engine.Evaluate(String.Format(@"input.scenarioNumber <- {0}", outerData.ScenarioNumber));
                    K_DataManager.engine.Evaluate(String.Format(@"input.history <- {0}", outerData.History));
                    K_DataManager.engine.Evaluate(String.Format(@"input.type <- ""{0}""", type));
                    K_DataManager.engine.Evaluate(String.Format(@"input.nullSigma <- {0}", outerData.nullSigma.ToString().ToUpper()));
                    K_DataManager.engine.Evaluate(String.Format(@"input.formula <- ""{0}""", outerData.formula));
                    K_DataManager.engine.Evaluate(String.Format(@"source('{0}')", link));

                    var res = K_DataManager.engine.Evaluate("forecast").AsDataFrame();
                    if (res != null && res.RowCount > 0)
                    {
                        DataTable table = new DataTable("SeriesScenarioValues");
                        K_DataManager.RDataFrameToDataTable(res, table);
                        K_DataManager.DataTableBulkCopy(table);
                        //K_DataManager.K_DataModel.UpdateContext(K_DataManager.K_DataModel.SeriesScenarioValues.AsEnumerable());
                        K_DataManager.K_DataModel.SeriesScenarioValues.Local.Clear();
                    }
                }
                else
                {
                    MessageBox.Show(String.Format("Series '{0}': script not found\r\n", link));
                }
            }
            catch (Exception)
            { }
        }
        public static List<InstrumentEntity> SynthesFinancialInstruments(List<InstrumentEntity> BaseInstruments)
        {
            var allInstruments = new List<InstrumentEntity>();
            try
            {
                string link = @"R scripts\Portfolio Optimization\Synth.R";

                if (File.Exists(String.Format("{0}", link)))
                {
                    link = link.Replace(@"\", @"/");
                    foreach (var instr in BaseInstruments)
                    {
                        var assetName = K_DataManager.GetAssetName(instr._TypeID);
                        allInstruments.Add(instr);
                        K_DataManager.engine.Evaluate(String.Format(@"DataSource <- ""{0}""", K_DataManager.DataSource));
                        K_DataManager.engine.Evaluate(String.Format(@"origID <- {0}", instr._ID));
                        K_DataManager.engine.Evaluate(String.Format(@"typeID <- {0}", instr._TypeID));
                        K_DataManager.engine.Evaluate(String.Format(@"source('{0}')", link));

                        var synthIDs = K_DataManager.engine.Evaluate("res[,1]").AsVector().ToArray();
                        foreach (var ID in synthIDs)
                        {
                            allInstruments.Add(new InstrumentEntity(Convert.ToInt32(ID), instr._TypeID, instr._CompanyID, "RUB", 200, instr._GroupID, new DateTime(), new DateTime()));
                            //CalculateFinancialInstruments(Convert.ToInt32(ID), assetName);
                        }
                        K_DataManager.MemoryRelease();
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Synthes failed!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            return allInstruments;
        }
        public static DataTable POResult = new DataTable("POResult");
        public static class POData
        {
            public static int itemID;
            public static int[] assetIDs;
            public static int[] indexIDs;
            public static int mc_num;
            public static int frequencyID;
            public static DateTime startDate;
            public static DateTime endDate;
        }

        public static DataTable PortfolioOptimization()
        {
            try
            {
                string link = @"R scripts\Portfolio Optimization\PO.R";

                if (File.Exists(String.Format("{0}", link)))
                {
                    link = link.Replace(@"\", @"/");
                    K_DataManager.engine.Evaluate(String.Format(@"DataSource    <- ""{0}""", K_DataManager.DataSource));
                    K_DataManager.engine.Evaluate(String.Format(@"assetIDs      <- {0}", MakeArrayR(POData.assetIDs)));         // --идентификаторы инструментов 
                    K_DataManager.engine.Evaluate(String.Format(@"itemID        <- {0}", POData.itemID));
                    K_DataManager.engine.Evaluate(String.Format(@"mc_num        <- {0}", POData.mc_num));
                    K_DataManager.engine.Evaluate(String.Format(@"frequencyID   <- {0}", POData.frequencyID));
                    K_DataManager.engine.Evaluate(String.Format(@"start_date    <- '{0}'", POData.startDate.ToString(K_DataManager.FormatDateYYYYMMDD)));
                    K_DataManager.engine.Evaluate(String.Format(@"end_date      <- '{0}'", POData.endDate.ToString(K_DataManager.FormatDateYYYYMMDD)));
                    K_DataManager.engine.Evaluate(String.Format(@"indexIDs      <- {0}", MakeArrayR(POData.indexIDs)));
                    K_DataManager.engine.Evaluate(String.Format(@"source('{0}')", link));

                    POResult = new DataTable("choices");
                    var choices = K_DataManager.engine.Evaluate("choices").AsDataFrame();
                    RDataFrameToDataTable(choices, POResult);

                    var allocations = K_DataManager.engine.Evaluate("allocations").AsDataFrame();
                    var dt = new DataTable("Optimization");
                    RDataFrameToDataTable(allocations, dt);

                    K_DataManager.MemoryRelease();
                    return dt;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Optimization failed!", MessageBoxButton.OK, MessageBoxImage.Warning);
            }
            return new DataTable();
        }

        /// <summary>
        /// 
        /// </summary>
        public static void LoadRFunctions()
        {
            using (SqlConnection conn = new SqlConnection(K_DataManager.ConnectionString))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(@"Select * from FunctionList", conn);
                SqlDataReader r = cmd.ExecuteReader();

                K_DataManager.engine.Evaluate(String.Format(@"param.DataSource <- ""{0}""", K_DataManager.DataSource));
                while (r.Read())
                {
                    string link = r["Link"].ToString();

                    if (File.Exists(String.Format("{0}", link)))
                    {
                        link = link.Replace(@"\", @"/");
                        K_DataManager.engine.Evaluate(String.Format(@"source('{0}')", link));
                    }
                }
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="tree"></param>
        /// <returns></returns>
        public static List<string> RunThroughTree(ItemCollection tree)
        {
            List<string> list = new List<string>();

            foreach (TreeViewItem item in tree)
            {
                if (item.Items.Count == 0)
                {
                    list.Add((item.Header as Label).Content.ToString());
                }
                else
                {
                    list.Add((item.Header as Label).Content.ToString());
                    list.AddRange(RunThroughTree(item.Items));
                }
            }

            return list;
        }

        public static void SetSeriesQuantiles(int seriesID, int frequencyID, double[] quantiles)
        {
            try
            {
                K_DataManager.engine.Evaluate("library(RODBC)");
                K_DataManager.engine.Evaluate("library(plyr)");
                K_DataManager.engine.Evaluate(String.Format(@"conn <- odbcConnect(""{0}"")", K_DataManager.DataSource));
                K_DataManager.engine.Evaluate(String.Format(@"data <- sqlQuery(conn, ""select * from SeriesScenarioValues where SeriesID = {0} and frequencyID = {1} order by 1,3,4"", errors = T)", seriesID, frequencyID));
                K_DataManager.engine.Evaluate(String.Format(@"sqlQuery(conn, ""delete from SeriesQuantiles where SeriesID = {0} and frequencyID = {1}"", errors = T)", seriesID, frequencyID));
                K_DataManager.engine.Evaluate(@"close(conn)");
                K_DataManager.engine.Evaluate(@"q <- na.omit(data.frame(""SeriesID"" = NA, ""FrequencyID"" = NA, ""Date""= NA, ""Quantile"" = NA, ""Value"" = NA))");

                foreach (double q in quantiles)
                {
                    //K_DataManager.engine.Evaluate(String.Format(@"q <- rbind(q, data.frame(""SeriesID"" = {0}, ""FrequencyID"" = {1}, ddply(data, c(""Date""), summarise, Quantile = {2}, Value=quantile(Value, {3}))))", seriesID, frequencyID, q * 100, q));
                    K_DataManager.engine.Evaluate(String.Format(@"tmp2 <- as.data.frame(ddply(data, c(""Date""), summarise, Quantile = {0}, Value=quantile(Value, {1})))", q * 100, q));
                    K_DataManager.engine.Evaluate(String.Format(@"tmp <- data.frame(""SeriesID"" = rep({0}, nrow(tmp2)), ""FrequencyID"" = rep({1}, nrow(tmp2)), ""Date"" = as.character(tmp2$Date), ""Quantile"" = tmp2$Quantile, ""Value"" = tmp2$Value)", seriesID, frequencyID));
                    K_DataManager.engine.Evaluate(String.Format(@"q <- rbind(q, tmp)"));
                }

                //K_DataManager.engine.Evaluate(@"sqlSave(conn, q, ""SeriesQuantiles"", append=T, rownames=F)");
                //K_DataManager.engine.Evaluate(@"close(conn)");

                var res = K_DataManager.engine.Evaluate("as.data.frame(q)").AsDataFrame();
                if (res != null && res.RowCount > 0)
                {
                    DataTable table = new DataTable("SeriesQuantiles");
                    K_DataManager.RDataFrameToDataTable(res, table);
                    K_DataManager.DataTableBulkCopy(table);
                }
            }
            catch (Exception) { }
            finally
            {
                K_DataManager.MemoryRelease();
            }
        }

        public static void SetSeriesQuantiles(string data, double[] quantiles)
        {
            try
            {
                K_DataManager.engine.Evaluate("library(RODBC)");
                K_DataManager.engine.Evaluate("library(plyr)");
                K_DataManager.engine.Evaluate(String.Format(@"data <- {0}", data)); // передаем переменную R, которая создается до вызова метода
                K_DataManager.engine.Evaluate(@"q <- na.omit(data.frame(""SeriesID"" = NA, ""FrequencyID"" = NA, ""Date""= NA, ""Quantile"" = NA, ""Value"" = NA))");

                string seriesIDs = K_DataManager.engine.Evaluate(@"paste(unique(data$SeriesID), collapse = "","")").AsCharacter()[0];
                string frequencyID = K_DataManager.engine.Evaluate(@"max(data$FrequencyID)").AsNumeric()[0].ToString();

                K_DataManager.ExecuteSQLQuery(@"delete from SeriesQuantiles where SeriesID in ({0}) and frequencyID = {1}", seriesIDs, frequencyID);

                foreach (double q in quantiles)
                {
                    K_DataManager.engine.Evaluate(String.Format(@"tmp2 <- as.data.frame(ddply(data, c(""SeriesID"", ""FrequencyID"", ""Date""), summarise, Quantile = {0}, Value=quantile(Value, {1})))", q * 100, q));
                    K_DataManager.engine.Evaluate(String.Format(@"tmp <- data.frame(""SeriesID"" = tmp2$SeriesID, ""FrequencyID"" = tmp2$FrequencyID, ""Date"" = as.character(tmp2$Date), ""Quantile"" = tmp2$Quantile, ""Value"" = tmp2$Value)"));
                    K_DataManager.engine.Evaluate(String.Format(@"q <- rbind(q, tmp)"));
                }

                var res = K_DataManager.engine.Evaluate("as.data.frame(q)").AsDataFrame();
                if (res != null && res.RowCount > 0)
                {
                    DataTable table = new DataTable("SeriesQuantiles");
                    K_DataManager.RDataFrameToDataTable(res, table);
                    K_DataManager.DataTableBulkCopy(table);
                }
            }
            catch (Exception e) { }
        }



        public static XyDataSeries<DateTime, double> GetAssetDataForChart(int AssetTypeID, int AssetID, int AAccountID)
        {
            try
            {
                string query = String.Format(@"select Date, Value from AssetHistoryValues where AssetTypeID = {0} and AssetID = {1} and AssetAccountID = {2} union select [Date], Value from AssetScenarioValues where AssetTypeID = {0} and AssetID = {1} and AssetAccountID = {2} and ScenarioNumber=1 order by Date",
                    AssetTypeID, AssetID, AAccountID);
                var dataAdapter = new SqlDataAdapter(query, K_DataManager.ConnectionString);
                var dt = new DataTable("Asset Account");
                dt.Clear();
                dataAdapter.Fill(dt);

                var dataSeries = new XyDataSeries<DateTime, double> { SeriesName = "Asset Account" };
                dataSeries.Append(dt.AsEnumerable().Select(row => row.Field<DateTime>("Date")).ToArray(),
                                    dt.AsEnumerable().Select(row => row.Field<double>("Value")).ToArray());
                return dataSeries;
            }
            catch (Exception exc)
            {
                return null;
            }
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="series"></param>
        /// <returns></returns>
        public static bool CheckSeries(string series)
        {
            if (series.Contains('.'))
            {
                Regex reg = new Regex(String.Format(@"({0}\.)([0-9\.]*)(?<{0}>([a-zA-Z0-9_]*))", "index"), RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace);
                MatchCollection Elements = reg.Matches(series);

                foreach (Match element in Elements)
                {
                    //string[] tmp = series.ToLower().Split('.');

                    if ((int)K_DataManager.ExecuteSQLQuery(String.Format(@"select isnull(seriesID,-1) from Series where lower(SeriesCode) = '{0}'", element.Groups["index"].Value)) == -1)
                    {
                        return false;
                    }
                }

                return true;
            }

            return false;
        }

        /// <summary>
        /// Enumerate all the descendants (children) of a visual object.
        /// </summary>
        /// <param name="parent">Starting visual (parent).</param>
        public static void InsertControls(Visual parent)
        {
            string name;
            string type;
            string text = "";

            try
            {
                var children = LogicalTreeHelper.GetChildren(parent);

                foreach (var childVisual in children)
                {
                    var frameworkElement = childVisual as FrameworkElement;

                    type = childVisual.GetType().ToString();

                    if (childVisual is Button)
                    {
                        name = ((Button)childVisual).Name;
                        text = ((Button)childVisual).Content.ToString();
                        type = "Button";
                        InsertControl(name, text, type);
                    }

                    if (childVisual is Label)
                    {
                        if (((Label)childVisual).HasContent)
                        {
                            name = ((Label)childVisual).Name;
                            text = ((Label)childVisual).Content.ToString();
                            type = "Label";
                            InsertControl(name, text, type);
                        }
                    }

                    if (childVisual is RadioButton)
                    {
                        name = ((RadioButton)childVisual).Name;
                        text = ((RadioButton)childVisual).Content.ToString();
                        type = "RadioButton";
                        InsertControl(name, text, type);
                    }

                    if (childVisual is TabItem)
                    {
                        name = ((TabItem)childVisual).Name;
                        text = ((TabItem)childVisual).Header.ToString();
                        type = "TabItem";
                        InsertControl(name, text, type);
                    }

                    if (childVisual is MainWindow)
                    {
                        name = ((MainWindow)childVisual).Name;
                        text = ((MainWindow)childVisual).Title;
                        type = "MainWindow";
                        InsertControl(name, text, type);
                    }

                    // Recursively enumerate children of the child visual object.
                    if (childVisual is Visual)
                    {
                        InsertControls((Visual)childVisual);
                    }
                }
            }
            catch (Exception)
            {
                return;
            }
        }

        public static void InsertControl(string name, string text, string type)
        {
            try
            {
                if ((int)K_DataManager.ExecuteSQLQuery(String.Format(@"select count(1) from Controls where ControlName = '{0}'", name)) == 0)
                {
                    K_DataManager.ExecuteSQLQuery(String.Format(@"Insert into Controls select '{0}', '{1}', N'{2}', 1, {3}", name, type, text, 1));
                    K_DataManager.ExecuteSQLQuery(String.Format(@"Insert into Controls select '{0}', '{1}', N'{2}', 1, {3}", name, type, text, 2));
                }
                else
                {
                    K_DataManager.ExecuteSQLQuery(String.Format(@"update Controls set Exist = 1 where ControlName = '{0}'", name));
                }
            }
            catch (Exception)
            {
            }
        }

        public static DataTable ToDataTable<T>(IList<T> data, string tableName = "")
        {
            PropertyDescriptorCollection properties = TypeDescriptor.GetProperties(typeof(T));

            DataTable table = new DataTable(tableName);

            foreach (PropertyDescriptor prop in properties)
                table.Columns.Add(prop.Name, Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType);

            foreach (T item in data)
            {
                DataRow row = table.NewRow();

                foreach (PropertyDescriptor prop in properties)
                    row[prop.Name] = prop.GetValue(item) ?? DBNull.Value;

                table.Rows.Add(row);
            }

            return table;
        }

        public static TreeViewItem GetCompanyByID(TreeViewItem parent, int ID)
        {
            if (parent != null && ID != -1)
            {
                if ((int)(parent.Tag as Label).Content == ID)
                    return parent;

                foreach (var item in parent.Items)
                {
                    TreeViewItem child = item as TreeViewItem;
                    if (GetCompanyByID(child, ID) != null)
                        return child;
                }
            }

            return null;
        }

        public static List<int> GetCompanyList(TreeViewItem treeItem)
        {
            var companyTreeIDs = new List<int>();

            if (!treeItem.IsExpanded)
            {
                companyTreeIDs = GetChildren(treeItem);
            }
            else
            {
                companyTreeIDs.Add(K_DataManager.CurrentCompanyID);
            }

            return companyTreeIDs;
        }

        /// <summary>
        /// Get asset name by id
        /// </summary>
        /// <param name="assetID"></param>
        /// <returns></returns>
        public static string GetAssetName(int assetID)
        {
            var assetName = K_DataManager.K_DataModel.AssetTypes.Find(assetID, 1).AssetTypeName.ToString();
            if (assetID == 4 || assetID == 5)
                assetName = "Swaps";

            assetName = assetName.Replace(" ", "");
            return assetName;
        }

        /// <summary>
        /// Recursively get Tag value from parent Tree Node
        /// </summary>
        /// <param name="parent"></param>
        /// <returns></returns>
        public static List<int> GetChildren(TreeViewItem parent)
        {
            var children = new List<int>();

            if (parent != null)
            {
                try
                {
                    children.Add(Int32.Parse((parent.Tag as Label).Content.ToString()));
                    foreach (var item in parent.Items)
                    {
                        TreeViewItem child = item as TreeViewItem;
                        children.AddRange(GetChildren(child));
                    }
                }
                catch (Exception exc)
                {
                }
            }

            return children;
        }

        /// <summary>
        /// Get all series ids from source
        /// </summary>
        /// <param name="source"></param>
        /// <returns></returns>
        public static HashSet<int> GetItemIndexIDs(string source, bool isRecursive = true)
        {
            var res = new HashSet<int>();

            Regex reg = new Regex(@"(index\.)(?<IndexCode>([a-zA-Z0-9_]*))", RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace);
            MatchCollection Elements = reg.Matches(source);
            foreach (Match elem in Elements)
            {
                try
                {
                    var seriesCode = elem.Groups["IndexCode"].Value.ToString();
                    var series = K_DataModel.Series.Where(s => s.SeriesCode == seriesCode).First();
                    res.Add(series.SeriesID);
                    if (isRecursive)
                    {
                        var LinkedIndices = GetItemIndexIDs(series.Parameters);
                        if (LinkedIndices.Count > 0)
                        {
                            LinkedIndicesByIndex[series.SeriesID] = LinkedIndices;
                            res.UnionWith(LinkedIndices);
                        }
                    }
                }
                catch (Exception exc)
                {
                }
            }

            return res;
        }

        /// <summary>
        /// Get child item's codes by item.formula
        /// </summary>
        /// <param name="source">item.formula</param>
        /// <returns></returns>
        public static List<string> GetItemsCodesByItem(string source)
        {
            var res = new List<string>();

            Regex reg = new Regex(@"(item\.)[0-9]+\.(?<ItemCode>([a-zA-Z0-9_]*))", RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace);
            MatchCollection Elements = reg.Matches(source);
            foreach (Match elem in Elements)
            {
                try
                {
                    var itemCode = elem.Groups["ItemCode"].Value.ToString();
                    res.Add(itemCode);
                }
                catch (Exception exc)
                {
                }
            }

            return res;
        }

        /// <summary>
        /// Get all dependent indices from asset
        /// </summary>
        /// <param name="assetName"></param>
        /// <param name="assetID"></param>
        /// <returns></returns>
        public static List<int> GetIndexIDsByAsset(string assetName, int assetID)
        {
            var indices = new List<int>();
            switch (assetName.ToLower())
            {
                case "accounts":
                    {
                        var asset = K_DataManager.K_DataModel.Accounts.Find(assetID);
                        if (asset != null)
                        {
                            if (asset.InterestRateType == 1 && asset.BaseInterestRateID > 0)
                            {
                                indices.Add((int)asset.BaseInterestRateID);
                            }

                            if (asset.CreditInterestRateType == 1 && asset.CreditBaseInterestRateID > 0)
                            {
                                indices.Add((int)asset.CreditBaseInterestRateID);
                            }
                        }
                        break;
                    }

                case "bonds":
                    {
                        var asset = K_DataManager.K_DataModel.Bonds.Find(assetID);
                        if (asset != null)
                        {
                            if (asset.RiskFreeRateType == 1 && asset.RiskFreeRateID > 0)
                            {
                                indices.Add((int)asset.RiskFreeRateID);
                            }

                            if (asset.CouponRateType == 1 && asset.CouponRateID > 0)
                            {
                                indices.Add((int)asset.CouponRateID);
                            }
                        }
                        break;
                    }

                case "contracts":
                    {
                        var asset = K_DataManager.K_DataModel.Contracts.Find(assetID);
                        if (asset != null)
                        {
                            //TODO: Calculate budget item  for volumes 
                            //indices.Add(asset.PlannedVolumeID);
                            indices.Add(asset.SpotPriceID);
                            indices.AddRange(GetItemIndexIDs(asset.Formula));
                        }
                        break;
                    }

                case "commoditySwaps":
                    {
                        var asset = K_DataManager.K_DataModel.CommoditySwaps.Find(assetID);
                        if (asset != null)
                        {
                            indices.Add(asset.FloatPriceID);
                        }
                        break;
                    }

                case "credits":
                    {
                        var asset = K_DataManager.K_DataModel.Credits.Find(assetID);
                        if (asset != null)
                        {
                            if (asset.InterestRateType == 1 && asset.BaseInterestRateID > 0)
                            {
                                indices.Add((int)asset.BaseInterestRateID);
                            }
                        }
                        break;
                    }

                case "deposits":
                    {
                        var asset = K_DataManager.K_DataModel.Deposits.Find(assetID);
                        if (asset != null)
                        {
                            if (asset.InterestRateType == 1 && asset.BaseInterestRateID > 0)
                            {
                                indices.Add((int)asset.BaseInterestRateID);
                            }
                        }
                        break;
                    }

                case "forwards":
                    {
                        var asset = K_DataManager.K_DataModel.Forwards.Find(assetID);
                        if (asset != null)
                        {
                            indices.Add(asset.RiskFreeAccID);
                            indices.Add(asset.RiskFreeBaseID);
                        }
                        break;
                    }

                case "swaps":
                    {
                        var asset = K_DataManager.K_DataModel.Swaps.Find(assetID);
                        if (asset != null)
                        {
                            indices.Add(asset.RiskFreeAccID);
                            indices.Add(asset.RiskFreeBaseID);
                        }
                        break;
                    }

                default:
                    break;
            }

            return indices;
        }

        /// <summary>
        /// Get all asset IDs from itemFormula group by asset type name
        /// </summary>
        /// <param name="itemFormula"></param>
        /// <returns></returns>
        public static Dictionary<string, List<int>> GetAssetIDsByType(string itemFormula)
        {
            Dictionary<string, List<int>> entities = new Dictionary<string, List<int>>();

            Regex reg = new Regex(@"(asset\.)([0-9\.]+)([a-zA-Z]+)\.([a-zA-Z0-9_]+)", RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace);
            var Elements = reg.Matches(itemFormula).OfType<Match>().Select(m => m.Value).Distinct();
            
            foreach (var element in Elements)
            {
                var token = element.Split('.');

                int assetID = 0;
                var assetType = token[2];
                if (Int32.TryParse(token[3], out assetID))
                {
                    if (entities.ContainsKey(assetType))
                    {
                        entities[assetType].Add(assetID);
                    }
                    else
                    {
                        entities.Add(assetType, new List<int>() { assetID });
                    }
                }
                else if (token[3].Length == 3)
                {
                    //Aggregation by currency
                    var query = String.Format(@"Select AssetID From {0} Where CompanyID = {1} And CurrencyCode = '{2}'", assetType, token[1], token[3]);
                    var assetByCurrency = K_DataManager.K_DataModel.Database.SqlQuery<int>(query).ToList();
                    if (entities.ContainsKey(assetType))
                    {
                        entities[assetType].AddRange(assetByCurrency);
                    }
                    else
                    {
                        entities.Add(assetType, assetByCurrency);
                    }
                }
            }

            return entities;
        }

        public static string GetTruncatedDate(ref string date, int frequencyID)
        {
            string[] freqNames = new string[] {"0", "day", "week", "month", "quarter", "year" };

            date = Convert.ToDateTime(K_DataManager.ExecuteSQLQuery(String.Format(@"select CONVERT(DATE,DATEADD({0}, DATEDIFF({0}, 0, '{1}'), 0),101)", freqNames[frequencyID], date))).ToString(K_DataManager.FormatDateYYYYMMDD).ToString();

            return freqNames[frequencyID];
        }

        public static DateTime GetTruncatedDate(DateTime date, int frequencyID)
        {
            string[] freqNames = new string[] {"0", "day", "week", "month", "quarter", "year" };

            return Convert.ToDateTime(K_DataManager.ExecuteSQLQuery(@"select CONVERT(DATE,DATEADD({0}, DATEDIFF({0}, 0, '{1}'), 0),101)", freqNames[frequencyID], date.ToString(FormatDateYYYYMMDD)));
        }

        /// <summary>
        /// Get all linked with item indices 
        /// </summary>
        /// <param name="itemFormula"></param>
        /// <returns></returns>
        public static List<int> GetIndexListByItem(string itemCode, string itemFormula)
        {
            var res = new List<int>();

            //Parse index links (Recursive)
            res.AddRange(GetItemIndexIDs(itemFormula));

            var dict = GetAssetIDsByType(itemFormula);
            if (dict.Count > 0)
            {
                try
                {
                    K_DataManager.AssetsByItemCode.Add(itemCode, dict);
                    dict.ToList().ForEach(r => r.Value.ForEach(a =>
                    {
                        var indexIDsByAsset = new HashSet<int>(GetIndexIDsByAsset(r.Key, a));
                        if (indexIDsByAsset.Count > 0)
                        {
                            if (K_DataManager.IndexIDsByAsset.ContainsKey(r.Key))
                            {
                                if (K_DataManager.IndexIDsByAsset[r.Key].ContainsKey(a))
                                {
                                    K_DataManager.IndexIDsByAsset[r.Key][a].UnionWith(indexIDsByAsset);
                                }
                                else
                                {
                                    K_DataManager.IndexIDsByAsset[r.Key].Add(a, indexIDsByAsset);
                                }
                            }
                            else
                            {
                                K_DataManager.IndexIDsByAsset.Add(r.Key, new Dictionary<int, HashSet<int>>() { { a, indexIDsByAsset } });
                            }
                            res.AddRange(indexIDsByAsset);
                        }
                    }));
                }
                catch (Exception exc)
                {
                }
            }
            return res;
        }

        /// <summary>
        /// Get list of unique index IDs by dependent items and asset for company.
        /// </summary>
        /// <param name="companyID"></param>)
        /// <returns></returns>
        public static List<int> GetIndexIDsByCompany(int companyID)
        {
           var items = K_DataManager.K_DataModel.BudgetItems.Where(r => r.CompanyID == companyID && r.Formula != "").OrderBy(r => r.LeftID).ToList();
            //Only items for company 
            var res = new List<int>();
            var iter = 0;
            items.ForEach(i => 
            {
                try
                {
                    iter++;
                    var  indexIDsByItem = GetIndexListByItem(i.BudgetItemCode, i.Formula);
                    if(indexIDsByItem.Count > 0)
                    {
                        K_DataManager.LinkedIndexIDsByItemCode.Add(i.BudgetItemCode, indexIDsByItem);
                        res.AddRange(indexIDsByItem);
                    }
                }
                catch (Exception exc)
                {
                }
            });

            return res.Distinct().ToList();
        }

        public static HashSet<int> GetLinkedIndicies(int ID)
        {
            var linkedIndicies = new HashSet<int>();
            foreach (var indexPair in LinkedIndicesByIndex)
            {
                if (indexPair.Value.Contains(ID))
                {
                    linkedIndicies.Add(indexPair.Key);
                    linkedIndicies.UnionWith(GetLinkedIndicies(indexPair.Key));
                }
            }
            return linkedIndicies;
        }

        public static void CalculateLinkedIndicies(K_DataManager.OuterData outerData)
        {
            var indexID = outerData.SeriesID;
            var linkedIndicies = GetLinkedIndicies(indexID);

            linkedIndicies.ToList().ForEach(id => { outerData.SeriesID = id; K_Forecast.CalculateIndex(outerData); });
        }


        public static InstrumentEntity GetInstrumentByTypeID(int typeID, int ID)
        {
            switch (typeID)
            {
                case (int)AssetTypes.Deposits:
                    {
                        var asset = K_DataManager.K_DataModel.Deposits.Find(ID);
                        if (asset != null)
                        {
                            return new K_DataManager.InstrumentEntity(asset.AssetID, (int)AssetTypes.Deposits, CurrentCompanyID, asset.CurrencyCode, asset.BankID, asset.GroupID, asset.OpenDate, asset.CloseDate, asset.Label,duration:asset.IsSynthetic ?? 0);
                        }
                        break;
                    }
                case (int)AssetTypes.Accounts:
                    {
                        var asset = K_DataManager.K_DataModel.Accounts.Find(ID);
                        if (asset != null)
                        {
                            return new K_DataManager.InstrumentEntity(asset.AssetID, (int)AssetTypes.Accounts, CurrentCompanyID, asset.CurrencyCode, asset.BankID, asset.GroupID, asset.OpenDate, asset.CloseDate, asset.Label);
                        }
                        break;
                    }
                case (int)AssetTypes.Credits:
                    {
                        var asset = K_DataManager.K_DataModel.Credits.Find(ID);
                        if (asset != null)
                        {
                            return new K_DataManager.InstrumentEntity(asset.AssetID, (int)AssetTypes.Credits, CurrentCompanyID, asset.CurrencyCode, asset.BankID, asset.GroupID, asset.OpenDate, asset.CloseDate, asset.Label);
                        }
                        break;
                    }
                case (int)AssetTypes.Bonds:
                    {
                        var asset = K_DataManager.K_DataModel.Bonds.Find(ID);
                        if (asset != null)
                        {
                            return new K_DataManager.InstrumentEntity(asset.AssetID, (int)AssetTypes.Bonds, CurrentCompanyID, asset.CurrencyCode, asset.BankID, asset.GroupID, asset.OpenDate, asset.CloseDate, asset.Label);
                        }
                        break;
                    }
                default:
                    break;
            }
        return null;
        }

        /// <summary>
        /// Convert dataTable with fields :Date, ScenarioNumber, Value  to IEnumerable[]
        /// </summary>
        /// <param name="table"></param>
        /// <returns></returns>
        public static IEnumerable[] ConvertDataTableToDataFrame(DataTable table)
        {
            var cols = new IEnumerable[3];
            cols[0] = table.AsEnumerable().Select(t => t.Field<DateTime>("Date").ToString(FormatDateYYYYMMDD)).ToArray();
            cols[1] = table.AsEnumerable().Select(t => t.Field<Int32>("ScenarioNumber")).ToArray();
            cols[2] = table.AsEnumerable().Select(t => t.Field<Double>("Value")).ToArray();

            return cols;
        }

        public  static void RDataFramePassToR<T>(string dataframeName, List<T> table, string tableName, string[] colNames)
        {
            var data = ConvertListToDataFrame(table, tableName);
            
            if(data.Length == 0)
            {
                MessageBox.Show("No data rows", "Information!", MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var df = engine.CreateDataFrame(data, columnNames: colNames);
            engine.SetSymbol(dataframeName.ToLower(), df);
            engine.Evaluate(String.Format("dict[['{0}']] <- {0}", dataframeName.ToLower()));
        }

        public static IEnumerable[] ConvertListToDataFrame<T>( List<T> list, string tableName)
        {
            var cols = new IEnumerable[0];
            switch (tableName)
            {
                case "BudgetItemScenarioValues":
                    {
                        var typedList = list as List<K_Treasury.BudgetItemScenarioValue>;
                        if (typedList == null) return cols;
                        cols = new IEnumerable[3];
                        cols[0] = typedList.Select(l => l.Date.ToString(FormatDateYYYYMMDD)).ToArray();
                        cols[1] = typedList.Select(l => l.ScenarioNumber).ToArray();
                        cols[2] = typedList.Select(l => l.Value).ToArray();
                        break;
                    }

                case "AssetScenarioValues":
                    {
                        var typedList = list as List<AssetScenarioValue>;
                        if (typedList == null) return cols;
                        cols = new IEnumerable[3];
                        cols[0] = typedList.Select(l => l.Date.ToString(FormatDateYYYYMMDD)).ToArray();
                        cols[1] = typedList.Select(l => l.ScenarioNumber).ToArray();
                        cols[2] = typedList.Select(l => l.Value).ToArray();
                        break;
                    }

                case "SeriesScenarioValues":
                    {
                        var typedList = list as List<K_Treasury.SeriesScenarioValue>;
                        if (typedList == null) return cols;
                        cols = new IEnumerable[3];
                        cols[0] = typedList.Select(l => l.Date.ToString(FormatDateYYYYMMDD)).ToArray();
                        cols[1] = typedList.Select(l => l.ScenarioNumber).ToArray();
                        cols[2] = typedList.Select(l => l.Value).ToArray();
                        break;
                    }

                case "SeriesHistoryValues":
                    {
                        var typedList = list as List<K_Treasury.SeriesHistoryValue>;
                        if (typedList == null) return cols;
                        cols = new IEnumerable[2];
                        cols[0] = typedList.Select(l => l.Date.ToString(FormatDateYYYYMMDD)).ToArray();
                        cols[1] = typedList.Select(l => l.Value).ToArray();
                        break;
                    }
                 case "PreparedForecastData":
                    {
                        var typedList = list as List<PreparedForecastData>;
                        if (typedList == null) return cols;
                        cols = new IEnumerable[3];
                        cols[0] = typedList.Select(l => l.Date.ToString(FormatDateYYYYMMDD)).ToArray();
                        cols[1] = typedList.Select(l => l.ScenarioNumber).ToArray();
                        cols[2] = typedList.Select(l => l.Value).ToArray();
                        break;
                    }
                default:
                    break;
            }

            return cols;
        }

        public static void PreloadAssets(string formula, DateTime startDate, DateTime endDate)
        {
            Regex reg = new Regex(@"(asset.([0-9])+.([a-zA-Z])+.([a-zA-Z0-9.]*)([a-zA-Z0-9])+)", RegexOptions.IgnoreCase | RegexOptions.IgnorePatternWhitespace);
            MatchCollection Elements = reg.Matches(formula);
            var colNames = new string[] { "Date", "ScenarioNumber", "Value" };

            foreach (Match element in Elements)
            {
                var token = element.Value.Split('.');
                var entity = token[0];
                var companyID = token[1];
                var assetType = token[2];
                int assetID;
                string query = "";

                if (Int32.TryParse(token[3], out assetID) && token.Length == 5) // No aggregation
                {
                    var accountCode = token[4];
                    query = String.Format(@"SELECT * from AssetScenarioValues" +
                                 " where AssetID = {0}" +
                                 " and AssetAccountID = (SELECT AssetAccountID from AssetAccounts WHERE lower(AssetAccountCode) = '{1}' and LanguageID = 1)" +
                                 " and AssetTypeID = (SELECT AssetTypeID from AssetTypes WHERE lower(AssetTypeName) = '{2}' and LanguageID = 1)" +
                                 " and Date >= '{3}' and Date <='{4}'", assetID, accountCode, assetType, startDate.ToString(K_DataManager.FormatDateYYYYMMDD), endDate.ToString(K_DataManager.FormatDateYYYYMMDD));
                    var list = K_DataManager.K_DataModel.AssetScenarioValues.SqlQuery(query).AsNoTracking().Select(l => new PreparedForecastData { Date = l.Date, Value = l.Value, ScenarioNumber = l.ScenarioNumber }).ToList();
                    K_DataManager.RDataFramePassToR(element.Value, list, "PreparedForecastData", colNames);

                }
                else if (assetType == "contracts" && token.Length == 4) // Aggregation by Company
                {
                    query = String.Format(@"SELECT * from AssetScenarioValues" +
                                 " where AssetTypeID = 9" +
                                 " and AssetID = (SELECT AssetID from Contracts WHERE CompanyID = {0})" +
                                 " and AssetAccountID = (SELECT AssetAccountID from AssetAccounts WHERE lower(AssetAccountCode) = '{1}' and LanguageID = 1)" +
                                 " and Date >= '{2}' and Date <='{3}'", companyID, token[3], startDate.ToString(K_DataManager.FormatDateYYYYMMDD), endDate.ToString(K_DataManager.FormatDateYYYYMMDD));
                    var list = K_DataManager.K_DataModel.AssetScenarioValues.SqlQuery(query).AsNoTracking().GroupBy(l => new { l.Date, l.ScenarioNumber }).Select(l =>
                        new PreparedForecastData { Date = l.Key.Date, Value = l.Sum(v => v.Value), ScenarioNumber = l.Key.ScenarioNumber }).ToList();
                    K_DataManager.RDataFramePassToR(element.Value, list, "PreparedForecastData", colNames);

                }

                else if (assetType == "contracts" && token[3] == "product") // Aggregation by Product
                {
                    query = String.Format(@"SELECT * from AssetScenarioValues"+
                                 " where AssetTypeID = 9"+
                                 " and AssetID = (SELECT AssetID from Contracts WHERE CompanyID = {0} AND ProductTypeID = {1})" +
                                 " and AssetAccountID = (SELECT AssetAccountID from AssetAccounts WHERE lower(AssetAccountCode) = '{2}' and LanguageID = 1)"+
                                 " and Date >= '{3}' and Date <='{4}'", companyID, token[4], token[5], startDate.ToString(K_DataManager.FormatDateYYYYMMDD), endDate.ToString(K_DataManager.FormatDateYYYYMMDD));
                    var list = K_DataManager.K_DataModel.AssetScenarioValues.SqlQuery(query).AsNoTracking().GroupBy(l => new { l.Date, l.ScenarioNumber }).Select(l =>
                        new PreparedForecastData { Date = l.Key.Date, Value = l.Sum(v => v.Value), ScenarioNumber = l.Key.ScenarioNumber }).ToList();
                    K_DataManager.RDataFramePassToR(element.Value, list, "PreparedForecastData", colNames);

                }

                else if (token[3].Length == 3 && token.Length == 5) // Aggregation by Currency
                {
                    var curencyCode = token[3];
                    var accountCode = token[4];
                    query = String.Format(@"SELECT * from AssetScenarioValues" +
                                 " where AssetID in (SELECT AssetID from {1} WHERE lower(CurrencyCode) = '{2}' and CompanyID = {3})" +
                                 " and AssetAccountID = (SELECT AssetAccountID from AssetAccounts WHERE lower(AssetAccountCode) = '{4}' and LanguageID = 1)" +
                                 " and AssetTypeID = (SELECT AssetTypeID from AssetTypes WHERE lower(AssetTypeName) = '{1}' and LanguageID = 1)" +
                                 " and Date >= '{5}' and Date <='{6}'", element.Value, assetType, curencyCode, companyID, accountCode, startDate.ToString(K_DataManager.FormatDateYYYYMMDD), endDate.ToString(K_DataManager.FormatDateYYYYMMDD));
                    var list = K_DataManager.K_DataModel.AssetScenarioValues.SqlQuery(query).AsNoTracking().GroupBy(l => new { l.Date, l.ScenarioNumber }).Select(l =>
                        new PreparedForecastData { Date = l.Key.Date, Value = l.Sum(v => v.Value), ScenarioNumber = l.Key.ScenarioNumber }).ToList();
                    K_DataManager.RDataFramePassToR(element.Value, list, "PreparedForecastData", colNames);

                }
                var data = new[] { "a", null };
                data[1] = "ddd";
            }
        }


        public class PreparedForecastData
        {
            public DateTime Date;
            public double Value;
            public int ScenarioNumber;
        }
    }

    /// <summary>
    /// Descriptive statistics items.
    /// </summary>
    public class DescriptiveStatistics
    {
        public string Data { get; set; }
        public double Maximum { get; set; }
        public double Minimum { get; set; }
        public double Mean { get; set; }
        public double Median { get; set; }
        public double Mode { get; set; }
        public double Procentile_5 { get; set; }
        public double Quartile_1st { get; set; }
        public double Quartile_3rd { get; set; }
        public double Procentile_95 { get; set; }
        public double Autocorrelation { get; set; }
        public double NormalityTest { get; set; }
    }
    
}
