using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LinqToExcel;

namespace SeeSharpLabSecond
{
    public class Threat
    {
        [LinqToExcel.Attributes.ExcelColumn("Идентификатор УБИ")]
        public int Id { set; get; }
        [LinqToExcel.Attributes.ExcelColumn("Наименование УБИ")]
        public string Name { set; get; }
        [LinqToExcel.Attributes.ExcelColumn("Описание")]
        public string Description { set; get; }
        [LinqToExcel.Attributes.ExcelColumn("Источник угрозы (характеристика и потенциал нарушителя)")]
        public string Source { set; get; }
        [LinqToExcel.Attributes.ExcelColumn("Объект воздействия")]
        public string Target { set; get; }
        [LinqToExcel.Attributes.ExcelColumn("Нарушение конфиденциальности")]
        public bool Confidence { set; get; } = false;
        [LinqToExcel.Attributes.ExcelColumn("Нарушение целостности")]
        public bool Integrity { set; get; } = false;
        [LinqToExcel.Attributes.ExcelColumn("Нарушение доступности")]
        public bool Availability { set; get; } = false;
        [LinqToExcel.Attributes.ExcelColumn("Дата включения угрозы в БнД УБИ")]
        public DateTime AddedDate { set; get; }
        [LinqToExcel.Attributes.ExcelColumn("Дата последнего изменения данных")]
        public DateTime ChangedDate { set; get; }

        public override string ToString()
        {
            return "Идентификатор УБИ: " + this.Id + ", Наименование УБИ: " + this.Name + ", Описание: " + this.Description + ", Источник угрозы (характеристика и потенциал нарушителя): " 
                + this.Source + ", Объект воздействия: "+ this.Target + ", Нарушение конфиденциальности: " + this.Confidence + ", Нарушение целостности: " + this.Integrity
                + ", Нарушение доступности: " + this.Availability + ", Дата включения угрозы в БнД УБИ: " + this.AddedDate.ToString("dd/MM/yyyy") + ", Дата последнего изменения данных: " + this.ChangedDate.ToString("dd/MM/yyyy") + "\n";
        }
    }
}
