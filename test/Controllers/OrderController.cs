using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using mvc.Models;


namespace mvc.Controllers
{
    public class OrderController : Controller
    {
        private OrderService orderService = new OrderService();

        /// <summary>
        /// 首頁
        /// </summary>
        /// <returns></returns>
        public ActionResult Index()
        {
            return View(orderService.GetOrders());
        }

        /// <summary>
        /// 新增訂單頁面
        /// </summary>
        /// <returns></returns>
        [HttpGet] 
        public ActionResult Create()
        {
            var CusDropListItem = orderService.GetCusDropListItem();
            var EmpDropListItem = orderService.GetEmpDropListItem();
            var ShipDropListItem = orderService.GetShipDropListItem();
            List<SelectListItem > CustId = new List<SelectListItem>();
            List<SelectListItem> EmpId = new List<SelectListItem>();
            List<SelectListItem> ShipperId = new List<SelectListItem>();
            foreach (var item in CusDropListItem)
            {
                CustId.Add(new SelectListItem() { Text=item.CompanyName,Value=item.CustId.ToString()});
            }
            foreach (var item in EmpDropListItem)
            {                
                EmpId.Add(new SelectListItem() { Text = item.EmpName, Value = item.EmpId.ToString()});
            }
            foreach (var item in ShipDropListItem)
            {
                ShipperId.Add(new SelectListItem() { Text = item.ShipperName, Value = item.ShipperId.ToString() });
            }
            ViewData["CustId"] = CustId;
            ViewData["EmpId"] = EmpId;
            ViewData["ShipperId"] = ShipperId;
            return View();
        }

        /// <summary>
        /// 新增訂單Action
        /// </summary>
        /// <param name="o">整份表單</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Create(Order o){
            if (ModelState.IsValid) {
                orderService.InsertOrder(o);
            }          
            return RedirectToAction("Index");
        }

        /// <summary>
        /// 搜尋訂單頁面
        /// </summary>
        /// <returns></returns>
        [HttpGet]
        public ActionResult Search()
        {
            var CusDropListItem = orderService.GetCusDropListItem();
            var EmpDropListItem = orderService.GetEmpDropListItem();
            var ShipDropListItem = orderService.GetShipDropListItem();
            List<SelectListItem> CustId = new List<SelectListItem>();
            List<SelectListItem> EmpId = new List<SelectListItem>();
            List<SelectListItem> ShipperId = new List<SelectListItem>();
            foreach (var item in CusDropListItem)
            {
                CustId.Add(new SelectListItem() { Text = item.CompanyName, Value = item.CustId.ToString() });
            }
            foreach (var item in EmpDropListItem)
            {
                EmpId.Add(new SelectListItem() { Text = item.EmpName, Value = item.EmpId.ToString() });
            }
            foreach (var item in ShipDropListItem)
            {
                ShipperId.Add(new SelectListItem() { Text = item.ShipperName, Value = item.ShipperId.ToString() });
            }
            ViewData["CustId"] = CustId;
            ViewData["EmpId"] = EmpId;
            ViewData["ShipperId"] = ShipperId;
            return View();
        }

        /// <summary>
        /// 搜尋訂單Action
        /// </summary>
        /// <param name="o"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Search(Order o)
        {           
            return View("SearchResult", orderService.SearchOrder(o));            
        }

        /// <summary>
        /// 顯示訂單詳情
        /// </summary>
        /// <param name="orderId">訂單編號</param>
        /// <returns></returns>
        public ActionResult Detail(int? orderId)
        {
            return View(orderService.GetOrderById(orderId));
        }

        /// <summary>
        /// 編輯訂單頁面
        /// </summary>
        /// <param name="orderId">訂單編號</param>
        /// <returns></returns>
        [HttpGet]
        public ActionResult Edit(int? orderId)
        {
            var CusDropListItem = orderService.GetCusDropListItem();
            var EmpDropListItem = orderService.GetEmpDropListItem();
            var ShipDropListItem = orderService.GetShipDropListItem();
            List<SelectListItem> CustId = new List<SelectListItem>();
            List<SelectListItem> EmpId = new List<SelectListItem>();
            List<SelectListItem> ShipperId = new List<SelectListItem>();
            foreach (var item in CusDropListItem)
            {
                CustId.Add(new SelectListItem() { Text = item.CompanyName, Value = item.CustId.ToString() });
            }
            foreach (var item in EmpDropListItem)
            {
                EmpId.Add(new SelectListItem() { Text = item.EmpName, Value = item.EmpId.ToString() });
            }
            foreach (var item in ShipDropListItem)
            {
                ShipperId.Add(new SelectListItem() { Text = item.ShipperName, Value = item.ShipperId.ToString() });
            }
            ViewData["CustIdItem"] = CustId;
            ViewData["EmpIdItem"] = EmpId;
            ViewData["ShipperIdItem"] = ShipperId;
            return View(orderService.GetOrderById(orderId));
        }

        /// <summary>
        /// 編輯訂單Action
        /// </summary>
        /// <param name="orderId">訂單編號</param>
        /// <param name="form">表單</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Edit(int? orderId, FormCollection form)
        {
            var data = orderService.GetOrderById(orderId);
            if (TryUpdateModel(data, "", form.AllKeys, new string[] {  }))
            {
                orderService.UpdateOrder(data);
            }
            else {
                ModelState.AddModelError("UpdateError", "更新失敗!");
                return View(orderService);
            }
            return RedirectToAction("Index");            
        }

        /// <summary>
        /// 刪除訂單頁面
        /// </summary>
        /// <param name="orderId">訂單編號</param>
        /// <returns></returns>
        [HttpGet]
        public ActionResult Delete(int? orderId)
        {
            return View(orderService.GetOrderById(orderId));
        }

        /// <summary>
        /// 刪除訂單Action
        /// </summary>
        /// <param name="orderId">訂單編號</param>
        /// <param name="col">表單</param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult Delete(int? orderId, FormCollection col)
        {
            orderService.GetOrderById(orderId);
            if (orderService != null)
                orderService.DeleteOrder(orderId);
            return RedirectToAction("Index");
        }

    }


    //    [HttpGet()]
    //    public JsonResult TestJson()
    //    {
    //        var result = new Models.Order()
    //        {
    //            CustId = "ABC",
    //            CustName = "測試JSON"
    //        };
    //        return this.Json(result, JsonRequestBehavior.AllowGet);
    //    }
}
