using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;

namespace mvc.Models
{
    public class Order
    {

        /// <summary>
        /// 訂單編號
        /// </summary>
        [DisplayName("訂單編號")]
        public int OrderId { get; set; }

        /// <summary>
        /// 客戶代號
        /// </summary>
        [DisplayName("客戶代號")]
        public int CustId { get; set; }

        /// <summary>
        /// 客戶公司名稱
        /// </summary>
        [DisplayName("客戶公司名稱")]
        public string CompanyName { get; set; }

        /// <summary>
        /// 業務(員工)代號
        /// </summary>
        [DisplayName("業務(員工)編號")]
        public int EmpId { get; set; }

        /// <summary>
        /// 業務(員工)姓名
        /// </summary>
        [DisplayName("業務(員工)姓名")]
        public string EmpName { get; set; }

        /// <summary>
        /// 訂單日期
        /// </summary>
        [DisplayName("訂單日期")]
        public string OrderDate { get; set; }

        /// <summary>
        /// 需要日期
        /// </summary>
        [DisplayName("需要日期")]
        public string RequireDate { get; set; }

        /// <summary>
        /// 出貨日期
        /// </summary>
        [DisplayName("出貨日期")]
        public string ShippedDate { get; set; }

        /// <summary>
        /// 出貨公司編號
        /// </summary>
        [DisplayName("出貨公司編號")]
        public int ShipperId { get; set; }

        /// <summary>
        /// 出貨公司名稱
        /// </summary>
        [DisplayName("出貨公司名稱")]
        public string ShipperName { get; set; }

        /// <summary>
        /// 運費
        /// </summary>
        [DisplayName("運費")]
        public double Freight { get; set; }

        /// <summary>
        /// 出貨說明
        /// </summary>
        [DisplayName("出貨說明")]
        public string ShipName { get; set; }

        /// <summary>
        /// 出貨地址
        /// </summary>
        [DisplayName("出貨地址")]
        public string ShipAddress { get; set; }

        /// <summary>
        /// 出貨城市
        /// </summary>
        [DisplayName("出貨城市")]
        public string ShipCity { get; set; }

        /// <summary>
        /// 出貨地區
        /// </summary>
        [DisplayName("出貨地區")]
        public string ShipRegion { get; set; }

        /// <summary>
        /// 郵遞區號
        /// </summary>
        [DisplayName("郵遞區號")]
        public string ShipPostalCode { get; set; }

        /// <summary>
        /// 出貨國家
        /// </summary>
        [DisplayName("出貨國家")]
        public string ShipCountry { get; set; }
    }
}
