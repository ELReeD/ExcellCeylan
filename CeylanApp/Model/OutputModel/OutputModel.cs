using System;
using System.Collections.Generic;
using System.Text;

namespace CeylanApp.Model.OutputModel
{

    public class OutputModel
    {
        public string SupervisorName { get; set; }
        public string Name { get; set; }
        public FoodModels Basmati { get; set; }
        public FoodModels Bic { get; set; }
        public FoodModels Garofalo_Pasta { get; set; }
        public FoodModels Isbre { get; set; }
        public FoodModels Stmichele { get; set; }
        public FoodModels Inkitchen { get; set; }
        public FoodModels Oyuncaqlar { get; set; }
        public FoodModels DrKorner { get; set; }
        public FoodModels GoodDay { get; set; }
        public FoodModels Bahlsen { get; set; }
        public FoodModels Lindt { get; set; }
        public FoodModels Zeni { get; set; }
        public FoodModels Total { get; set; }

    }

    public class FoodModels
    {
        public string Target { get; set; }
        public string Real { get; set; }
        public string RealProcent { get; set; }
        public string Fkms { get; set; }
        public string Trend { get; set; }
        public string TrendProcent { get; set; }
    }

}
