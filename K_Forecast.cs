 using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Data.SqlClient;
using System.Windows;
using System.Data;
using RDotNet;

namespace K_Treasury
{
    public static class K_Forecast
    {
        /// <summary>
        /// Initialize/Finalize Sensivity table
        /// </summary>
        /// <param name="Indexes"></param>
        /// <param name="indexUpdate"></param>
        public static void UpdateInsertSensitivityData(int[] Indexes, int indexUpdate = -1)
        {
            try
            {
                if (indexUpdate == -1)
                {
                   //K_DataManager.ExecuteSQLQuery("delete Sensitivity where BudgetItemID in (select BudgetItemID from BudgetItems where CompanyID in ({0}))", K_DataManager.CompanyList);
                    K_DataManager.K_DataModel.Sensitivities.RemoveRange(K_DataManager.K_DataModel.Sensitivities.Where(s => 
                        K_DataManager.K_DataModel.BudgetItems.Where(b => K_DataManager.CompanyList.Contains(b.CompanyID.ToString())).Select(b => b.BudgetItemID).Contains(s.BudgetItemID)));
                    K_DataManager.K_DataModel.SaveChanges();
                }

                K_DataManager.K_DataModel.SensitivityBudgetItemScenarioValues.Where(b => 
                    K_DataManager.K_DataModel.BudgetItems.Where(i => K_DataManager.CompanyIDsList.Contains((int)i.CompanyID)).Select(i=>i.BudgetItemID).Contains((int)b.BudgetItemID)).ToList().ForEach(b => 
                {
                    if (indexUpdate == -1)
                    {
                        foreach (int seriesID in Indexes)
                        {
                            var sensivity = K_DataManager.K_DataModel.Sensitivities.Find(b.BudgetItemID, seriesID, b.Date);
                            if (sensivity != null)
                            {
                                sensivity.OldValue = b.Value;
                                sensivity.NewValue = 0;
                                sensivity.Sensitivity1 = 0;
                                sensivity.h = 0;
                            }
                            else
                            {
                                K_DataManager.K_DataModel.Sensitivities.Add(
                                    new Sensitivity()
                                    {
                                        BudgetItemID = b.BudgetItemID,
                                        SeriesID = seriesID,
                                        Date = b.Date,
                                        OldValue = b.Value,
                                        NewValue = 0,
                                        Sensitivity1 = 0,
                                        h = 0
                                    });
                            }
                        }
                    }
                    else
                    {
                        var sensivity = K_DataManager.K_DataModel.Sensitivities.Find(b.BudgetItemID, indexUpdate, b.Date);
                        sensivity.NewValue = b.Value;
                    }
                });

                K_DataManager.K_DataModel.SaveChanges();

            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        /// <summary>
        /// Initialize/Finalize Tornado table
        /// </summary>
        /// <param name="Indexes"></param>
        /// <param name="indexUpdate"></param>
        public static void UpdateInsertTornadoData( int quantile, int indexUpdate)
        {
            try
            {
                var sensivities = K_DataManager.K_DataModel.SensitivityBudgetItemScenarioValues.Where(b =>
                   K_DataManager.K_DataModel.BudgetItems.Where(i => K_DataManager.CompanyIDsList.Contains((int)i.CompanyID)).Select(i => i.BudgetItemID).Contains((int)b.BudgetItemID)).ToList();
                foreach (var b in sensivities)
                {
                    var tornado = K_DataManager.K_DataModel.Tornadoes.Find(b.BudgetItemID, indexUpdate, b.Date);
                    if (tornado == null) // First quantile update, create new record
                    {
                        tornado = new Tornado()
                               {
                                   BudgetItemID = b.BudgetItemID,
                                   SeriesID = indexUpdate,
                                   Date = b.Date
                               };
                        K_DataManager.K_DataModel.Tornadoes.Add(tornado);
                    }
                    switch (quantile)
                    {
                        case 1:
                            tornado.Fq1 = b.Value; break;
                        case 5:
                            tornado.Fq5 = b.Value; break;
                        case 50:
                            tornado.Fq50 = b.Value; break;
                        case 95:
                            tornado.Fq95 = b.Value; break;
                        case 99:
                            tornado.Fq99 = b.Value; break;
                        default:
                            break;
                    }
                }

                K_DataManager.K_DataModel.SaveChanges();
            }
            catch (Exception exc)
            {
                MessageBox.Show(exc.Message);
            }
        }

        public static DataFrame ForecastModel(object _outerData, bool updateChart = false)
        {
            string link = string.Empty;
            string type = string.Empty;

            K_DataManager.LoadRFunctions();
            K_DataManager.OuterData outerData = (K_DataManager.OuterData)_outerData;

            try
            {
                link = K_DataManager.ExecuteSQLQuery(String.Format("select Link from forecastmodels where modelID = {0} and languageID = 1", outerData.ModelID)).ToString();
                type = K_DataManager.ExecuteSQLQuery(String.Format("select Parameters from forecastmodels where modelID = {0} and languageID = 1", outerData.ModelID)).ToString();

                if (File.Exists(String.Format("{0}", link)))
                {
                    link = link.Replace(@"\", @"/");

                    if (outerData.ScenarioNumber > 0)
                    {
                        K_DataManager.ExecuteSQLQuery(String.Format("delete SeriesScenarioValues where SeriesID = {0} and FrequencyID = {1}", outerData.SeriesID, outerData.FrequencyID));
                    }

                    K_DataManager.engine.Evaluate(String.Format(@"input.DataSource <- ""{0}""", K_DataManager.DataSource));
                    K_DataManager.engine.Evaluate(String.Format(@"input.forecast_start_date <- as.Date(""{0}"")", Convert.ToDateTime(outerData.StartDateForecast).ToString(K_DataManager.FormatDateYYYYMMDD)));
                    K_DataManager.engine.Evaluate(String.Format(@"input.forecast_end_date <- as.Date(""{0}"")", Convert.ToDateTime(outerData.EndDateForecast).ToString(K_DataManager.FormatDateYYYYMMDD)));
                    K_DataManager.engine.Evaluate(String.Format(@"input.estimate<-{0}", outerData.Estimate.ToString().ToUpper()));
                    K_DataManager.engine.Evaluate(String.Format(@"input.seriesID <- {0}", outerData.SeriesID));
                    K_DataManager.engine.Evaluate(String.Format(@"input.frequencyID <- {0}", outerData.FrequencyID));
                    K_DataManager.engine.Evaluate(String.Format(@"input.scenarioNumber <- {0}", outerData.ScenarioNumber));
                    K_DataManager.engine.Evaluate(String.Format(@"input.history <- {0}", outerData.History));
                    K_DataManager.engine.Evaluate(String.Format(@"input.type <- ""{0}""", type));
                    K_DataManager.engine.Evaluate(String.Format(@"input.nullSigma <- {0}", outerData.nullSigma.ToString().ToUpper()));
                    K_DataManager.engine.Evaluate(String.Format(@"input.formula <- ""{0}""", outerData.formula));
                    K_DataManager.engine.Evaluate(String.Format(@"input.modelID <- -1"));
                    K_DataManager.engine.Evaluate(String.Format(@"input.IsForward <- FALSE"));
                    K_DataManager.engine.Evaluate(String.Format(@"source('{0}')", link));

                    outerData.Parameters = K_DataManager.engine.Evaluate("output.parameters").AsDataFrame();

                    if (outerData.ScenarioNumber != -2)
                    {
                        var res = K_DataManager.engine.Evaluate("forecast").AsDataFrame();
                        if (res != null && res.RowCount > 0)
                        {
                            DataTable table = new DataTable("SeriesScenarioValues");
                            K_DataManager.RDataFrameToDataTable(res, table);
                            K_DataManager.DataTableBulkCopy(table);
                            K_DataManager.K_DataModel.SeriesScenarioValues.Local.Clear();

                            K_DataManager.SetSeriesQuantiles("forecast", K_DataManager.PortfolioOptimizationQuantiles);
                        }

                        if (updateChart)
                        {
                            K_Chart.UpdateForecastLineChart(outerData);
                        }
                    }
                }
                else
                {
                    MessageBox.Show(String.Format("Series '{0}': script not found\r\n", link));
                }
            }
            catch (Exception e)
            {
                MessageBox.Show(String.Format(e.Message));
            }

            return outerData.Parameters;
        }

        public static void CalculateIndex(K_DataManager.OuterData outerData)
        {
                outerData.History = 0;
                outerData.nullSigma = true;

                try
                {
                    outerData.ModelID = (int)K_DataManager.ExecuteSQLQuery(String.Format(@"select ModelID from Series where LanguageID = 1 and SeriesID = {0}", outerData.SeriesID));
                    if (outerData.ModelID == 11)
                    {
                        outerData.formula = K_DataManager.ExecuteSQLQuery(String.Format(@"select parameters from Series where LanguageID = 1 and SeriesID = {0}", outerData.SeriesID)).ToString();
                    }

                    outerData.Estimate = false;
                }   
                catch (Exception)
                {
                    outerData.ModelID = 1; //base model
                    outerData.Estimate = true;
                }

                K_Forecast.ForecastModel(outerData);
        }
        
        /// <summary>
        /// 
        /// </summary>
        /// <param name="outerData"></param>
        /// <param name="link"></param>
        public static void CalculateItem(K_DataManager.OuterData outerData)
        {
            var link = K_DataManager.LinkItemsScript;
            if (File.Exists(link) && File.Exists(K_DataManager.LinkGetValueScript))
            {
                link = link.Replace(@"\", @"/");

                K_DataManager.ExecuteSQLQuery("delete from {0} where BudgetItemID in (select BudgetItemID from BudgetItems where CompanyID in ({1}))", outerData.BudgetItemTable, K_DataManager.CompanyList);
                var items = new Dictionary<int, string>();
                K_DataManager.K_DataModel.BudgetItems.Where(b => K_DataManager.CompanyIDsList.Contains((int)b.CompanyID) && b.Formula != "").OrderByDescending(b => b.BudgetItemID).ToList().ForEach(b => items.Add(b.BudgetItemID, b.Formula));

                var iterNum = 0;
                var startDate = Convert.ToDateTime(outerData.StartDateForecast);
                var endDate = Convert.ToDateTime(outerData.EndDateForecast);
                // Для записи не рассчитанных статей.
                var calclulatedItems = new List<int>();

                // Need preload data from db
                //K_DataManager.engine.Evaluate(String.Format(@"source('{0}')", K_DataManager.LinkGetValueScript.Replace(@"\", @"/")));
                K_DataManager.engine.Evaluate(String.Format(@"source('{0}')", @"R scripts\Forecasting\getValueSpeedUp.R".Replace(@"\", @"/")));

                K_DataManager.engine.Evaluate(String.Format(@"DataSource <- ""{0}""", K_DataManager.DataSource));
                K_DataManager.engine.Evaluate(String.Format(@"f_s <- as.Date(""{0}"")", outerData.StartDateForecast));
                K_DataManager.engine.Evaluate(String.Format(@"f_e <- as.Date(""{0}"")", outerData.EndDateForecast));
                K_DataManager.engine.Evaluate(String.Format(@"mc_num <- {0}", outerData.ScenarioNumber));
                K_DataManager.engine.Evaluate(String.Format(@"frequencyID <- {0}", outerData.FrequencyID));
                K_DataManager.engine.Evaluate(String.Format(@"item_table <- ""{0}""", outerData.BudgetItemTable));
                K_DataManager.engine.Evaluate(@"library(RODBC)");
                K_DataManager.engine.Evaluate(@"conn <- odbcConnect(DataSource)");

                // Пробегаем по всем статьям, чтобы рассчитать.
                while (items.Count > 0)
                {
                    var notCalculated = new Dictionary<int, string>();
                    iterNum++;

                    foreach (var itemID in items.Keys)
                    {
                        try
                        {
                            // Обработка базовых статей
                            var IsAllBaseItemsCalculated = true;
                            foreach (var baseItemID in K_DataManager.GetDependentItems(items[itemID], "item"))
                            {
                                if (calclulatedItems.Contains(baseItemID) || baseItemID == itemID) //Статья ссылается на саму себя
                                {
                                    continue;
                                }
                                else// Базовая статья не рассчитана
                                {
                                    if (!items.ContainsKey(baseItemID)) continue;// Базовая статья не принадлежит текущей компании, Пропускаем её

                                    notCalculated.Add(itemID, items[itemID]);
                                    IsAllBaseItemsCalculated = false;
                                    break;
                                }
                            }

                            if (IsAllBaseItemsCalculated)
                            {
                                #region Calculate Item
                                try
                                {
                                    var item = K_DataManager.K_DataModel.BudgetItems.Local.Where(b => b.BudgetItemID == itemID).First();

                                    #region Preload Indices to dataframe
                                    var indices = K_DataManager.GetItemIndexIDs( item.Formula, isRecursive: false);
                                    foreach (var indexID in indices)
                                    {
                                        var indexCode = K_DataManager.K_DataModel.Series.Find(indexID).SeriesCode;
                                        var list = K_DataManager.K_DataModel.SeriesScenarioValues.Where(i => i.SeriesID == indexID && i.FrequencyID == outerData.FrequencyID &&
                                            i.Date >= startDate && i.Date <= endDate && i.ScenarioNumber <= outerData.ScenarioNumber).ToList();
                                        var colNames = new string[] { "Date", "ScenarioNumber", "Value" };
                                        if (list.Count > 0)
                                        {
                                            K_DataManager.RDataFramePassToR("index." + indexCode.ToLower(), list, "SeriesScenarioValues", colNames);
                                        }
                                        else
                                        {
                                            var listHistory = K_DataManager.K_DataModel.SeriesHistoryValues.Where(i => i.SeriesID == indexID && i.FrequencyID == outerData.FrequencyID &&
                                            i.Date >= startDate && i.Date <= endDate).ToList();
                                            colNames = new string[] { "Date", "Value" };
                                            K_DataManager.RDataFramePassToR("index." + indexCode.ToLower(), listHistory, "SeriesHistoryValues", colNames);
                                        }
                                    }
                                    #endregion

                                    #region Preload Assets to dataframe
                                    K_DataManager.PreloadAssets(item.Formula, startDate, endDate);

                                    #endregion
                                    K_DataManager.engine.Evaluate(String.Format(@"item_path <- '{0}'", String.Format(@"item.{0}.{1}", item.CompanyID, item.BudgetItemCode).ToLower()));
                                    K_DataManager.engine.Evaluate(String.Format(@"item_formula <- '{0}'", item.Formula));
                                    K_DataManager.engine.Evaluate(String.Format(@"item_id <- {0}", item.BudgetItemID));
                                    K_DataManager.engine.Evaluate(String.Format(@"source('{0}')", link));

                                    DataFrame res = K_DataManager.engine.Evaluate("forecast").AsDataFrame();
                                    if (res.RowCount > 0)
                                    {
                                        K_DataManager.engine.Evaluate(String.Format(@"dict[['{0}']] <- forecast[,c(""Date"",""ScenarioNumber"",""Value"")]", item.BudgetItemCode.ToLower()));
                                        #region Save data to DB overhead ~0
                                        DataTable table = new DataTable(outerData.BudgetItemTable);
                                        K_DataManager.RDataFrameToDataTable(res, table);
                                        K_DataManager.DataTableBulkCopyAsync(table);
                                        //K_DataManager.DataTableBulkCopy(table);
                                        #endregion
                                    }
                                }
                                catch (Exception exc)
                                {
                                    MessageBox.Show(exc.Message);
                                }
                                #endregion

                                calclulatedItems.Add(itemID);
                            }
                        }
                        catch (Exception exc)
                        {
                            MessageBox.Show(exc.Message);
                        }
                    }

                    items = notCalculated;
                }
            }
            K_DataManager.engine.Evaluate(@"close(conn)");
            K_DataManager.MemoryRelease();
        }

        /// <summary>
        /// Calculate all assets linked with items
        /// </summary>
        /// <param name="outerData"></param>
        /// <param name="items"></param>
        public static void CalculateLinkedAssets(K_DataManager.OuterData outerData, List<string> items, bool isSaveRestore = false)
        {
            var calculatedAssetIDsByType = new Dictionary<string, HashSet<int>>();

            foreach (var item in items)
            {
                if (!K_DataManager.AssetsByItemCode.ContainsKey(item))
                    continue;
                foreach (var assetType in K_DataManager.AssetsByItemCode[item])
                {
                    var calculatedIDs = new HashSet<int>();
                    foreach (var assetID in assetType.Value)
                    {
                        if (calculatedAssetIDsByType.ContainsKey(assetType.Key) && calculatedAssetIDsByType[assetType.Key].Contains(assetID))
                            continue;

                        K_DataManager.CalculateFinancialInstruments(assetID, assetType.Key, DateTime.Parse(outerData.StartDateForecast), 1, outerData.FrequencyID);
                        calculatedIDs.Add(assetID);
                    }

                    if (calculatedIDs.Count == 0)
                        continue;

                    if (calculatedAssetIDsByType.ContainsKey(assetType.Key))
                    {
                        calculatedAssetIDsByType[assetType.Key].UnionWith(calculatedIDs);
                    }
                    else
                    {
                        calculatedAssetIDsByType.Add(assetType.Key, calculatedIDs);
                    }
                }
            }
            if (isSaveRestore)
            {
                K_Forecast.SaveAssetRestoreValues(calculatedAssetIDsByType, outerData.StartDateForecast, outerData.EndDateForecast);
            }
        }

        /// <summary>
        /// Calculate all assets linked with indices
        /// </summary>
        /// <param name="sensivityIndexID"></param>
        /// <param name="startDateForecast"></param>
        /// <param name="frequencyID"></param>
        public static Dictionary<string, HashSet<int>> CalculateAssetsByIndex(List<int> sensivityIndexIDs, K_DataManager.OuterData outerData)
        {
            var calculatedAssets = new Dictionary<string, HashSet<int>>();

            // All assets by item
            foreach (var assets in K_DataManager.AssetsByItemCode.Values)
            {
                foreach (var assetType in assets.Keys)
                {
                    var calculatedAssetIDs = new HashSet<int>();
                    foreach (var assetID in assets[assetType])
                    {
                        if (calculatedAssets.ContainsKey(assetType) && calculatedAssets[assetType].Contains(assetID))
                            continue;

                        // Check if asset linked with target index
                        if (K_DataManager.IndexIDsByAsset.ContainsKey(assetType) && K_DataManager.IndexIDsByAsset[assetType].ContainsKey(assetID))
                        {
                            if (K_DataManager.IndexIDsByAsset[assetType][assetID].Any(id => sensivityIndexIDs.Contains(id)))
                            {
                                K_DataManager.CalculateFinancialInstruments(assetID, assetType, DateTime.Parse(outerData.StartDateForecast), 1, outerData.FrequencyID);
                                calculatedAssetIDs.Add(assetID);
                            }
                        }
                    }

                    if (calculatedAssetIDs.Count == 0)
                        continue;

                    if (calculatedAssets.ContainsKey(assetType))
                    {
                        calculatedAssets[assetType].UnionWith(calculatedAssetIDs);
                    }
                    else
                    {
                        calculatedAssets.Add(assetType, calculatedAssetIDs);
                    }
                }
            }

            return calculatedAssets;
        }

        public static void SaveAssetRestoreValues(Dictionary<string, HashSet<int>> assets, string startDate, string endDate)
        {
            foreach (var assetType in assets.Keys)
            {
                var assetTypeID = K_DataManager.K_DataModel.AssetTypes.Where(a => a.AssetTypeName == assetType).Select(a => a.AssetTypeID).First();
                foreach (var assetID in assets[assetType])
                {
                    K_DataManager.ExecuteSQLQuery(String.Format(@"delete AssetRestoreValues where AssetTypeID={0} and AssetID={1} and Date >= '{2}' and Date <= '{3}'", assetTypeID, assetID, startDate, endDate));

                    K_DataManager.ExecuteSQLQuery(String.Format(@"insert into AssetRestoreValues select  AssetTypeID, AssetID, AssetAccountID, Date, ScenarioNumber, Value from AssetScenarioValues 
                            where AssetTypeID={0} and AssetID={1} and Date >= '{2}' and Date <= '{3}'", assetTypeID, assetID, startDate, endDate));
                }
            }
        }

        public static void RestoreAssetValues(Dictionary<string, HashSet<int>> assets, string startDate, string endDate)
        {
            foreach (var assetType in assets.Keys)
            {
                var assetTypeID = K_DataManager.K_DataModel.AssetTypes.Where(a => a.AssetTypeName == assetType).Select(a => a.AssetTypeID).First();
                foreach (var assetID in assets[assetType])
                {
                    K_DataManager.ExecuteSQLQuery(String.Format(@"delete AssetScenarioValues where AssetTypeID={0} and AssetID={1} and Date >= '{2}' and Date <= '{3}'", assetTypeID, assetID, startDate, endDate));

                    K_DataManager.ExecuteSQLQuery(String.Format(@"insert into AssetScenarioValues select  AssetTypeID, AssetID, AssetAccountID, Date, Value, ScenarioNumber, [ForecastValueType]=2 from AssetRestoreValues
                            where AssetTypeID={0} and AssetID={1} and Date >= '{2}' and Date <= '{3}'", assetTypeID, assetID, startDate, endDate));
                }
            }
        }

        public static void CalculateSensivity(K_DataManager.OuterData outerData, double h)
        {
            //Sensitivity data initialization
            K_Forecast.UpdateInsertSensitivityData(K_DataManager.sensitivityIndexes.Values.ToArray<int>());

            //Adding h to indexes`
            foreach (int indexID in K_DataManager.sensitivityIndexes.Values)
            {
                //TODO: Maybe h% from AVG value
                var query = String.Format(@"Select AVG(Value) from SeriesScenarioValues where SeriesID = {0} and FrequencyID = {1} and Date >= '{2}' and Value <> 0 group by SeriesID", indexID, outerData.FrequencyID, outerData.StartDateForecast);
                double v0 = Convert.ToDouble(K_DataManager.ExecuteSQLQuery(query));
                double delta = h * (v0 == 0 ? 1 : v0);

                //K_DataManager.ExecuteSQLQuery(String.Format(@"update SeriesScenarioValues set Value = Value + {0} where SeriesID = {1} and FrequencyID = {2} and Date >= '{3}' and Date <= '{4}'",
                //    h, indexID, outerData.FrequencyID, outerData.StartDateForecast, outerData.EndDateForecast));
                
                K_DataManager.K_DataModel.SeriesScenarioValues.Where(s => s.SeriesID == indexID && s.FrequencyID == outerData.FrequencyID && s.Date >= outerData.StartDate && s.Date <= outerData.EndDate).ToList().ForEach(i => i.Value += delta);
               // K_DataManager.K_DataModel.SaveChanges();

                // Save indices restore values
                var linkedIndexIDs = K_DataManager.GetLinkedIndicies(indexID).ToList();

                //Calculate linked indices
                outerData.SeriesID = indexID;
                K_DataManager.CalculateLinkedIndicies(outerData);
                //Calculate assets linked with index
                var calculatedAssets = K_Forecast.CalculateAssetsByIndex(linkedIndexIDs, outerData);

                K_Forecast.CalculateItem(outerData);
                K_Forecast.UpdateInsertSensitivityData(K_DataManager.sensitivityIndexes.Values.ToArray<int>(), indexID);

                K_DataManager.K_DataModel.SeriesScenarioValues.Where(s => s.SeriesID == indexID && s.FrequencyID == outerData.FrequencyID && s.Date >= outerData.StartDate && s.Date <= outerData.EndDate).ToList().ForEach(i => i.Value -= delta);
                K_DataManager.K_DataModel.SaveChanges();

                // Restore linked and base indices values
                if (linkedIndexIDs.Count > 0)
                {
                    K_DataManager.ExecuteSQLQuery(String.Format(@"delete from SeriesScenarioValues where SeriesID in ({0}) and FrequencyID = {1} and Date >= '{2}' and Date <= '{3}'",
                        String.Join(",", linkedIndexIDs), outerData.FrequencyID, outerData.StartDateForecast, outerData.EndDateForecast));
                    query = String.Format(@"insert into SeriesScenarioValues select [SeriesID],[FrequencyID],[Date],[ScenarioNumber],[Value],RestoreValue=NULL  from SeriesRestoreValues where SeriesID in ({0}) and FrequencyID = {1} and Date >= '{2}' and Date <= '{3}'",
                        String.Join(",", linkedIndexIDs), outerData.FrequencyID, outerData.StartDateForecast, outerData.EndDateForecast);
                    K_DataManager.ExecuteSQLQuery(query);
                    // Restore changed assets
                    K_Forecast.RestoreAssetValues(calculatedAssets, outerData.StartDateForecast, outerData.EndDateForecast);
                }

                //K_DataManager.ExecuteSQLQuery(String.Format(@"Update Sensitivity set h = {0} where seriesID = {1}", h, indexID));
                K_DataManager.K_DataModel.Sensitivities.Where(s => s.SeriesID == indexID).ToList().ForEach(i => i.h = delta);
                K_DataManager.K_DataModel.SaveChanges();
            }

            // TODO: Delete restore infor from series assets items
            K_DataManager.K_DataModel.Sensitivities.Where(s =>
                K_DataManager.K_DataModel.BudgetItems.Where(b => K_DataManager.CompanyList.Contains(b.CompanyID.ToString())).Select(b => b.BudgetItemID).Contains(s.BudgetItemID)).ToList().ForEach(i => i.Sensitivity1 = Math.Round((i.NewValue.Value - i.OldValue.Value) / i.h.Value, 3));
            K_DataManager.K_DataModel.SaveChanges();

//            K_DataManager.ExecuteSQLQuery(String.Format(@"update Sensitivity set Sensitivity = round((NewValue - OldValue)/h, 3) 
//                                                        where BudgetItemID in (select BudgetItemID from BudgetItems where CompanyID in ({0}))", K_DataManager.CompanyList));
        }

        public static void CalculateTornado(K_DataManager.OuterData outerData)
        {
            var frequencyID = outerData.FrequencyID;
            var forecastStartDate = outerData.StartDateForecast;
            var forecastEndDate = outerData.EndDateForecast;
            var forecastStartHistoryDate = outerData.StartDateHistory;

            K_DataManager.K_DataModel.Tornadoes.RemoveRange(K_DataManager.K_DataModel.Tornadoes.Where(s => 
                K_DataManager.K_DataModel.BudgetItems.Where(b => K_DataManager.CompanyList.Contains(b.CompanyID.ToString())).Select(b => b.BudgetItemID).Contains(s.BudgetItemID)));
            K_DataManager.K_DataModel.SaveChanges();

            K_DataManager.ExecuteSQLQuery(String.Format(@"update SeriesScenarioValues set RestoreValue = Value where FrequencyID = {0} and Date >= '{1}' and Date <= '{2}'", frequencyID, forecastStartDate, forecastEndDate));

            foreach (int id in K_DataManager.sensitivityIndexes.Values)
            {
                var data = K_Chart.GetQuantiles(String.Format(@"select * from SeriesHistoryValues where SeriesID = {0} and date >= '{1}' and date <= '{2}' order by Value asc", id, forecastStartHistoryDate, forecastStartDate), K_DataManager.Quantiles);

                if (data != null)
                {
                    for (int k = 0; k < data.Length; k++)
                    {
                        K_DataManager.ExecuteSQLQuery(String.Format(@"update SeriesScenarioValues set Value = {0} where SeriesID = {1} and FrequencyID = {2} and Date >= '{3}' and Date <= '{4}'", data[k], id, frequencyID, forecastStartDate, forecastEndDate));
                        K_Forecast.CalculateItem(outerData);
                        K_Forecast.UpdateInsertTornadoData((int)(K_DataManager.Quantiles[k] * 100), id);
                    }
                }
            }

            K_DataManager.ExecuteSQLQuery(String.Format(@"update SeriesScenarioValues set Value = RestoreValue where FrequencyID = {0} and Date >= '{1}' and Date <= '{2}'", frequencyID, forecastStartDate, forecastEndDate));
            K_DataManager.ExecuteSQLQuery(String.Format(@"update SeriesScenarioValues set RestoreValue = NULL where FrequencyID = {0} and Date >= '{1}' and Date <= '{2}'", frequencyID, forecastStartDate, forecastEndDate));

        }

    }
}
