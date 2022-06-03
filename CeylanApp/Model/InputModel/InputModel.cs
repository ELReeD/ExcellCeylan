using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Text;

namespace CeylanApp.Model.InputModel
{
    public class InputModel
    {
        [JsonProperty("ДоговорКонтрагента.Агент")]
        public string AgentName { get; set; }

        [JsonProperty("Basmati Rise, ")]
        public string BasmatiRice { get; set; }

        [JsonProperty("BIC, ")]
        public string Bic { get; set; }

        [JsonProperty("BTC Foods, ")]
        public string BtcFood { get; set; }

        [JsonProperty("GAROFALO PASTA, ")]
        public string GarofaloPasta { get; set; }

        [JsonProperty("ISBRE SU, ")]
        public string Isbre { get; set; }

        [JsonProperty("BTC Foods Sweets, ")]
        public string BtcFoodSweet { get; set; }

        [JsonProperty("STMICHELE PECENIYE, ")]
        public string StmichelePecniye { get; set; }

        [JsonProperty("BTC NonFoods, ")]
        public string BTCNonFood { get; set; }

        [JsonProperty("INKITCHEN, ")]
        public string Inkitchen { get; set; }

        [JsonProperty("OYUNCAQLAR, ")]
        public string Oyuncaqlar { get; set; }

        [JsonProperty("Dr.Korner, ")]
        public string DrKorner { get; set; }

        [JsonProperty("Good Day Coffee, ")]
        public string GoodDayCoffe { get; set; }

        [JsonProperty("Lindt&Bahlsen, ")]
        public string LindtBahlsen { get; set; }

        [JsonProperty("Bahlsen, ")]
        public string Bahlsen { get; set; }

        [JsonProperty("Lindt, ")]
        public string Lindt { get; set; }

        [JsonProperty("ZENI, ")]
        public string Zeni { get; set; }

        [JsonProperty("Итог")]
        public string Sumary { get; set; }
    }
}
