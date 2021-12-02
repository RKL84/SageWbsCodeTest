using Sage.Estimating;
using Sage.Estimating.Data;
using Sage.Estimating.Takeoff;
using System;
using System.Collections.Generic;
using System.Linq;

namespace SageWbsCodeTest
{
    class Program
    {
        private static StandardDBConnectionInfo _standardDbConnectionInfo;
        private static EstimateConnectionInfo _estimateConnectionInfo;
        private static StandardDBEntities _standardDBEntities;

        static void Main(string[] args)
        {
            var estimateName = $"estimate_{(new Random()).Next()}";
            var connectionInfo = InstanceService.ReadActiveInstance();
            var standardDbInfo = new StandardDBConnectionInfo(connectionInfo, "eTakeoff_Bridge_Testing");
            var creationInfo = new EstimateCreationInfo(estimateName, standardDbInfo) { LockOnCreate = true };
            var estimateDbService = new EstimateDBService(EstimateDBService.ReadActiveEstimateDB());
            _estimateConnectionInfo = estimateDbService.CreateEstimate(creationInfo);

            var standardDbLinkService = new StandardDBLinkService(_estimateConnectionInfo);
            var standardDbLinkEntity = standardDbLinkService.ReadActivateStandardDBLink();
            _standardDbConnectionInfo = standardDbLinkEntity.StandardDBConnectionInfo;
            var readOptions = new EntityReadOptions
            {
                ReadWithAll = true
            };

            var standardDBService = new StandardDBService(_standardDbConnectionInfo)
            {
                ReadOptions = readOptions
            };
            _standardDBEntities = standardDBService.ReadEntities();


            TestTakeoffAssemblyWithWbs();

            var estimateService = new EstimateService(_estimateConnectionInfo);
            estimateService.UnlockEstimate();

            Console.WriteLine("Test Completed");
            Console.Read();
        }

        static void TestTakeoffAssemblyWithWbs()
        {
            var assemblyName = AssemblyEntity.FormatName("A2010.10 3200");
            var assembly = _standardDBEntities.Assemblies.FirstOrDefault(a => a.Name == assemblyName);
            var standardDbWbsCodeCollection = GetStandardDbWbsDefinitionCollection();
            var wbsDefinitionCollection = GetWbsEstimateDbDefinitionCollection();

            var defaultWbsCode = "CSI UF10 Level 1";
            var defaultWbsValue = "B";
            var wbsDefinition = wbsDefinitionCollection.FirstOrDefault(wbs => wbs.Description == defaultWbsCode);
            using (var cache = new TakeoffSessionCache(_estimateConnectionInfo, _standardDBEntities))
            {
                using (var session = new AssemblyTakeoffSession(assembly, cache) { WriteZeroQuantityItems = true })
                {
                    foreach (var variable in session.TakeoffVariables)
                        variable.Value = 10; // add value to generate non zero item quantities

                    session.AddPass();
                    var items = session.GetItems();

                    WbsValueEntity wbsValue = null;
                    foreach (var item in items)
                    {
                        var index = item.WbsValues.FindIndex
                            (v => v.ParentEntity.Description == defaultWbsCode);

                        if (index >= 0)
                        {
                            wbsValue = item.WbsValues[index];
                            if (string.Equals(wbsValue.Value.Trim(),
                            defaultWbsValue?.Trim(), StringComparison.InvariantCulture))
                                break;// The value is already setup.
                            item.WbsValues.RemoveAt(index);
                        }

                        var wbsDefValue = wbsDefinition.Values
                            .FirstOrDefault(wv =>
                            string.Equals(wv.Value.Trim(),
                            defaultWbsValue?.Trim(), StringComparison.InvariantCulture));
                        WbsValueEntity newWbsValue;
                        if (wbsDefValue != null)
                        {
                            newWbsValue = wbsDefValue;
                        }
                        else
                        {
                            var formattedValue =
                                WbsValueEntity.FormatValue(defaultWbsValue,
                                wbsDefinition.FormattedWbsValueLength);
                            wbsValue = new WbsValueEntity(formattedValue);
                            if (standardDbWbsCodeCollection != null)
                            {
                                var wbsDbItem = standardDbWbsCodeCollection
                                    .FirstOrDefault(a => a.Name == defaultWbsCode);
                                if (wbsDbItem != null)
                                {
                                    var wbsDbValue = wbsDbItem.WbsValueCollection.
                                        FirstOrDefault(a => string.Equals(a.Id.Trim(),
                                        defaultWbsValue?.Trim(), StringComparison.InvariantCulture));
                                    var description = wbsDbValue?.Description;
                                    if (!string.IsNullOrEmpty(description))
                                        wbsValue.Description = description;
                                }
                            }
                            //check the database values
                            wbsDefinition.Values.Add(wbsValue);
                            newWbsValue = wbsValue;
                        }

                        item.WbsValues.Add(newWbsValue);
                    }

                    session.WriteEntities();
                }
            }
        }

        static List<WbsCode> GetStandardDbWbsDefinitionCollection()
        {
            var standardDbWbsCodeCollection = new List<WbsCode>();
            var wbsStandardDbService = new WbsService(_standardDbConnectionInfo)
            {
                ReadOptions = new EntityReadOptions { ReadWithAll = true }
            };
            var wbsStandardDbWbsDefinitionCollection = wbsStandardDbService.ReadWbsDefinitions().Where
                (dtwd => !string.IsNullOrEmpty(dtwd.Description)
                          && (dtwd.WbsType == WbsType.Detail ||
                          dtwd.WbsType == WbsType.Takeoff)).ToList();

            foreach (var item in wbsStandardDbWbsDefinitionCollection)
            {
                var wbsItem = new WbsCode();
                wbsItem.Name = item.Description;
                foreach (var wbsValue in item.Values)
                {
                    wbsItem.WbsValueCollection.Add(new WbsValue()
                    {
                        Id = wbsValue.Value,
                        Description = wbsValue.Description
                    });
                }
                standardDbWbsCodeCollection.Add(wbsItem);
            }

            return standardDbWbsCodeCollection;
        }

        static List<WbsDefinitionEntity> GetWbsEstimateDbDefinitionCollection()
        {
            var wbsEstimateDbService = new WbsService(_estimateConnectionInfo)
            {
                ReadOptions = new EntityReadOptions { ReadWithAll = true }
            };

            return wbsEstimateDbService.ReadWbsDefinitions().Where
                (dtwd => !string.IsNullOrEmpty(dtwd.Description)
                          && (dtwd.WbsType == WbsType.Detail ||
                          dtwd.WbsType == WbsType.Takeoff)).ToList();
        }

    }

    public class WbsCode
    {
        public WbsCode()
        {
            WbsValueCollection = new List<WbsValue>();
        }

        public string Name { get; set; }
        public List<WbsValue> WbsValueCollection { get; set; }
    }

    public class WbsValue
    {
        public string Id { get; set; }
        public string Description { get; set; }
    }
}
