using System;
using System.Collections.Generic;
using System.Linq;

namespace CTRM.DB
{
    public class DataManager
    {
        // Singleton pattern: One central instance for the whole app
        private static DataManager _instance;
        public static DataManager Instance => _instance ??= new DataManager();

        // The master list of all historical data
        public List<IBV_CallData> AllData { get; private set; } = new List<IBV_CallData>();

        // Property to quickly check if we have data before opening Forecasting
        public bool HasData => AllData != null && AllData.Count > 0;

        public void SetData(List<IBV_CallData> data)
        {
            AllData = data;
        }

        public void ClearData()
        {
            AllData.Clear();
        }
    }
}