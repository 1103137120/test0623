using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;


namespace mvc.Models
{
    /// <summary>
    /// 訂單服務
    /// </summary>
    public class OrderService
    {
        /// <summary>
        /// 連線
        /// </summary>
        /// <returns></returns>
        private string GetDBConnectionString()
        {
            return
                System.Configuration.ConfigurationManager.ConnectionStrings["DBConn"].ConnectionString.ToString();
        }
        /// <summary>
        /// 新增訂單
        /// </summary>
        /// <param name="order"></param>
        public void InsertOrder(Models.Order order)
        {
            Models.Order result = new Order();
            DataTable dt = new DataTable();
            String sql = @"INSERT INTO [Sales].[Orders]([CustomerID],[EmployeeID],[OrderDate],[RequiredDate],[ShippedDate],[Freight],[ShipperID],[ShipName],[ShipAddress],[ShipCity],[ShipRegion],[ShipPostalCode],[ShipCountry])
                           VALUES (@CustomerID,@EmployeeID,@OrderDate,@RequiredDate,@ShippedDate,@Freight,@ShipperID,@ShipName,@ShipAddress,@ShipCity,@ShipRegion,@ShipPostalCode,@ShipCountry)";

            using (SqlConnection conn = new SqlConnection(this.GetDBConnectionString()))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.Add(new SqlParameter("@CustomerID",order.CustId));
                cmd.Parameters.Add(new SqlParameter("@EmployeeID", order.EmpId));
                cmd.Parameters.Add(new SqlParameter("@OrderDate", order.OrderDate));
                cmd.Parameters.Add(new SqlParameter("@RequiredDate", order.RequireDate));
                cmd.Parameters.Add(new SqlParameter("@ShippedDate", order.ShippedDate));
                cmd.Parameters.Add(new SqlParameter("@Freight", order.Freight));
                cmd.Parameters.Add(new SqlParameter("@ShipperID", order.ShipperId));
                cmd.Parameters.Add(new SqlParameter("@ShipName", order.ShipName));
                cmd.Parameters.Add(new SqlParameter("@ShipAddress", order.ShipAddress));
                cmd.Parameters.Add(new SqlParameter("@ShipCity", order.ShipCity));
                cmd.Parameters.Add(new SqlParameter("@ShipRegion", order.ShipRegion));
                cmd.Parameters.Add(new SqlParameter("@ShipPostalCode", order.ShipPostalCode));
                cmd.Parameters.Add(new SqlParameter("@ShipCountry", order.ShipCountry));
                SqlDataAdapter sqlAdapter = new SqlDataAdapter(cmd);
                sqlAdapter.Fill(dt);
                conn.Close();
            }          
        }
        /// <summary>
        /// 新增訂單明細
        /// </summary>
        /// <param name="order"></param>
        public void InsertOrderDetail(Models.Order order) { }

        public List<Order> SearchOrder(Models.Order order) {
            List<Order> result = new List<Order>();
            DataTable dt = new DataTable();
            String sqlwhere = "1=1";
            if (order.OrderId != 0)
            {
                sqlwhere = sqlwhere+ "AND OrderID=@OrderID";
            }
            if (order.CustId != 0)
            {
                sqlwhere = sqlwhere + " AND c.CustomerID=@CustomerID";                
            }
            if (order.EmpId != 0)
            {
                sqlwhere = sqlwhere + " AND e.EmployeeID=@EmployeeID";
            }
            if (order.ShipperName != null)
            {
                sqlwhere = sqlwhere + " AND ShipperName=@ShipperName";
            }
            if (order.OrderDate != null)
            {
                sqlwhere = sqlwhere + " AND OrderDate=@OrderDate";
            }
            if (order.RequireDate != null)
            {
                sqlwhere = sqlwhere + " AND RequiredDate=@RequiredDate"; ;
            }
            if (order.ShippedDate != null)
            {
                sqlwhere = sqlwhere + " AND ShippedDate=@ShippedDate"; ;
            }
            String sql = @"SELECT OrderID,c.CompanyName,LastName+' '+FirstName AS EmpName,ship.CompanyName AS ShipperName,convert(char(10),OrderDate,111) AS OrderDate,convert(char(10),OrderDate,111) AS RequiredDate,convert(char(10),ShippedDate,111) AS ShippedDate
                           FROM [Sales].[Orders] AS o                           
	                       JOIN [Sales].[Customers] AS c
	                       ON o.CustomerID=c.CustomerID
						   JOIN [HR].[Employees] AS e
						   ON o.EmployeeID=e.EmployeeID
						   JOIN [Production].[Suppliers] AS ship
						   ON o.[ShipperID]=ship.SupplierID
                           WHERE "+ sqlwhere;

                using (SqlConnection conn = new SqlConnection(this.GetDBConnectionString()))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                if (order.OrderId !=0) { cmd.Parameters.Add(new SqlParameter("@OrderID", order.OrderId));  }
                if (order.CustId != 0) { cmd.Parameters.Add(new SqlParameter("@CustomerID", order.CustId)); }
                if (order.EmpId!= 0) { cmd.Parameters.Add(new SqlParameter("@EmployeeID", order.EmpId));}
                if (order.ShipperName != null) { cmd.Parameters.Add(new SqlParameter("@ShipperName", order.ShipperName)); }
                if (order.OrderDate != null) { cmd.Parameters.Add(new SqlParameter("@OrderDate", order.OrderDate)); }
                if (order.RequireDate != null) { cmd.Parameters.Add(new SqlParameter("@RequiredDate", order.RequireDate));}
                if (order.ShippedDate != null) { cmd.Parameters.Add(new SqlParameter("@ShippedDate", order.ShippedDate)); }
                
                SqlDataAdapter sqlAdapter = new SqlDataAdapter(cmd);
                sqlAdapter.Fill(dt);
                conn.Close();
            }

            result = (from i in dt.AsEnumerable()
                      select new Order()
                      {
                          OrderId = i.Field<int>("OrderID"),
                          CompanyName = i.Field<string>("CompanyName"),
                          EmpName = i.Field<string>("EmpName"),
                          ShipperName = i.Field<string>("ShipperName"),
                          OrderDate = i.Field<string>("OrderDate"),
                          RequireDate = i.Field<string>("RequiredDate"),
                          ShippedDate = i.Field<string>("ShippedDate")

                      }).ToList<Order>();

            return result;
        }



        /// <summary>
        /// 刪除訂單
        /// </summary>
        public void DeleteOrder(int? orderId) {
            DataTable dt = new DataTable();
            String sql = @"DELETE FROM [Sales].[Orders]
                           WHERE [OrderID]=@OrderId";

            using (SqlConnection conn = new SqlConnection(this.GetDBConnectionString()))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.Add(new SqlParameter("@OrderId", orderId));
                SqlDataAdapter sqlAdapter = new SqlDataAdapter(cmd);
                sqlAdapter.Fill(dt);
                conn.Close();
            }
        }
        /// <summary>
        /// 修改訂單
        /// </summary>
        /// <param name="order"></param>
        public void UpdateOrder(Models.Order order) {
            DataTable dt = new DataTable();
            String sql = @"UPDATE [Sales].[Orders]
                           SET [CustomerID]=@CustomerID,[EmployeeID]=@EmployeeID,[OrderDate]=@OrderDate,[RequiredDate]=@RequiredDate,[ShippedDate]=@ShippedDate,[Freight]=@Freight,[ShipperID]=@ShipperID,[ShipName]=@ShipName,[ShipAddress]=@ShipAddress,[ShipCity]=@ShipCity,[ShipRegion]=@ShipRegion,[ShipPostalCode]=@ShipPostalCode,[ShipCountry]=@ShipCountry
                           WHERE [OrderID]=@OrderID";

            using (SqlConnection conn = new SqlConnection(this.GetDBConnectionString()))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.Add(new SqlParameter("@OrderID", order.OrderId));
                cmd.Parameters.Add(new SqlParameter("@CustomerID", order.CustId));
                cmd.Parameters.Add(new SqlParameter("@EmployeeID", order.EmpId));
                cmd.Parameters.Add(new SqlParameter("@OrderDate", order.OrderDate));
                cmd.Parameters.Add(new SqlParameter("@RequiredDate", order.RequireDate));
                cmd.Parameters.Add(new SqlParameter("@ShippedDate", order.ShippedDate));
                cmd.Parameters.Add(new SqlParameter("@Freight", order.Freight));
                cmd.Parameters.Add(new SqlParameter("@ShipperID", order.ShipperId));
                cmd.Parameters.Add(new SqlParameter("@ShipName", order.ShipName));
                cmd.Parameters.Add(new SqlParameter("@ShipAddress", order.ShipAddress));
                cmd.Parameters.Add(new SqlParameter("@ShipCity", order.ShipCity));
                cmd.Parameters.Add(new SqlParameter("@ShipRegion", order.ShipRegion));
                cmd.Parameters.Add(new SqlParameter("@ShipPostalCode", order.ShipPostalCode));
                cmd.Parameters.Add(new SqlParameter("@ShipCountry", order.ShipCountry));
                SqlDataAdapter sqlAdapter = new SqlDataAdapter(cmd);
                sqlAdapter.Fill(dt);
                conn.Close();
            }
        }

        /// <summary>
        /// 取得每一筆訂單明細
        /// </summary>
        /// <param name="orderId">訂單ID</param>
        /// <returns></returns>
        public Models.Order GetOrderById(int? orderId) {
            Models.Order result = new Order();

            DataTable dt = new DataTable();
            String sql = @"SELECT OrderID,o.CustomerID,o.EmployeeID,c.CompanyName,LastName+' '+FirstName AS EmpName,convert(char(10),OrderDate,111) AS OrderDate,convert(char(10),OrderDate,111) AS RequiredDate,convert(char(10),ShippedDate,111) AS ShippedDate,o.ShipperID,ship.CompanyName AS ShipperName,Freight,ShipName,ShipAddress,ShipCity,ShipCountry,ShipPostalCode,ShipRegion
                           FROM [Sales].[Orders] AS o                           
	                       JOIN [Sales].[Customers] AS c
	                       ON o.CustomerID=c.CustomerID
						   JOIN [HR].[Employees] AS e
						   ON o.EmployeeID=e.EmployeeID
						   JOIN [Production].[Suppliers] AS ship
						   ON o.[ShipperID]=ship.SupplierID
                           WHERE o.[OrderID]=@OrderId";

            using (SqlConnection conn = new SqlConnection(this.GetDBConnectionString()))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.Add(new SqlParameter("@OrderId", orderId));
                SqlDataAdapter sqlAdapter = new SqlDataAdapter(cmd);
                sqlAdapter.Fill(dt);                
                conn.Close();                    
            }
            
            foreach (DataRow dr in dt.Rows)
            {
                result.OrderId = Convert.ToInt32(dr["OrderID"]);
                result.CustId = Convert.ToInt32(dr["CustomerID"]);
                result.EmpId = Convert.ToInt32(dr["EmployeeID"]);
                result.CompanyName = Convert.ToString(dr["CompanyName"]);
                result.EmpName= Convert.ToString(dr["EmpName"]);
                result.ShipperId = Convert.ToInt32(dr["ShipperID"]); 
                result.ShipperName = Convert.ToString(dr["ShipperName"]);
                result.OrderDate = Convert.ToString(dr["OrderDate"]);
                result.RequireDate = Convert.ToString(dr["RequiredDate"]);
                result.ShippedDate = Convert.ToString(dr["ShippedDate"]);
                result.Freight = Convert.ToInt32(dr["Freight"]); 
                result.ShipName=Convert.ToString(dr["ShipName"]);             
                result.ShipAddress= Convert.ToString(dr["ShipAddress"]);
                result.ShipCity= Convert.ToString(dr["ShipCity"]);
                result.ShipRegion= Convert.ToString(dr["ShipRegion"]);
                result.ShipPostalCode= Convert.ToString(dr["ShipPostalCode"]);
                result.ShipCountry= Convert.ToString(dr["ShipCountry"]);           
                   
            }
            return result;          
        }

        /// <summary>
        /// 取得訂單列表
        /// </summary>
        /// <returns></returns>
        public List<Order> GetOrders() {
            List<Order> result = new List<Order>();
            DataTable dt = new DataTable();
            String sql = @"SELECT OrderID,c.CompanyName,LastName+' '+FirstName AS EmpName,ship.CompanyName AS ShipperName,convert(char(10),OrderDate,111) AS OrderDate,convert(char(10),OrderDate,111) AS RequiredDate,convert(char(10),ShippedDate,111) AS ShippedDate
                           FROM [Sales].[Orders] AS o                           
	                       JOIN [Sales].[Customers] AS c
	                       ON o.CustomerID=c.CustomerID
						   JOIN [HR].[Employees] AS e
						   ON o.EmployeeID=e.EmployeeID
						   JOIN [Production].[Suppliers] AS ship
						   ON o.[ShipperID]=ship.SupplierID";                         

            using (SqlConnection conn = new SqlConnection(this.GetDBConnectionString()))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);                
                SqlDataAdapter sqlAdapter = new SqlDataAdapter(cmd);
                sqlAdapter.Fill(dt);
                conn.Close();
            }

            result = (from i in dt.AsEnumerable() 
                      select new Order()
                      {
                          OrderId = i.Field<int>("OrderID"),
                          CompanyName = i.Field<string>("CompanyName"),
                          EmpName = i.Field<string>("EmpName"),
                          ShipperName= i.Field<string>("ShipperName"),
                          OrderDate = i.Field<string>("OrderDate"),
                          RequireDate = i.Field<string>("RequiredDate"),
                          ShippedDate = i.Field<string>("ShippedDate")

                      }).ToList<Order>();         

            return result;
        }
        /// <summary>
        /// 取得客戶下拉式選單資料
        /// </summary>
        /// <returns></returns>
        public List<Order> GetCusDropListItem()
        {
            List<Models.Order> MapResult = new List<Order>();
            DataTable dt = new DataTable();
            String sql = @"SELECT CustomerID,CompanyName
                           FROM [Sales].[Customers]

                           SELECT EmployeeID,LastName+' '+FirstName AS EmpName
                           FROM [HR].[Employees]

                           SELECT ShipperID,CompanyName AS ShipperName
                           FROM [Sales].[Shippers]";

            using (SqlConnection conn = new SqlConnection(this.GetDBConnectionString()))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataAdapter sqlAdapter = new SqlDataAdapter(cmd);
                sqlAdapter.Fill(dt);
                conn.Close();
            }

            foreach (DataRow Row in dt.Rows)
            {
                MapResult.Add(new Order()
                {
                    CustId = (int)Row["CustomerID"],
                    CompanyName = Row["CompanyName"].ToString()
                }
                );
            }

            return MapResult;
        }
        /// <summary>
        /// 取得員工下拉式選單資料
        /// </summary>
        /// <returns></returns>
        public List<Order> GetEmpDropListItem()
        {
            List<Models.Order> MapResult = new List<Order>();
            DataTable dt = new DataTable();
            String sql = @"SELECT EmployeeID,LastName+' '+FirstName AS EmpName
                           FROM [HR].[Employees]

                           SELECT ShipperID,CompanyName AS ShipperName
                           FROM [Sales].[Shippers]";

            using (SqlConnection conn = new SqlConnection(this.GetDBConnectionString()))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataAdapter sqlAdapter = new SqlDataAdapter(cmd);
                sqlAdapter.Fill(dt);
                conn.Close();
            }

            foreach (DataRow Row in dt.Rows)
            {
                MapResult.Add(new Order()
                {
                    EmpId = (int)Row["EmployeeID"],
                    EmpName = Row["EmpName"].ToString()
                }
                );
            }
            return MapResult;
        }

        /// <summary>
        /// 取得出貨公司下拉式選單資料
        /// </summary>
        /// <returns></returns>
        public List<Order> GetShipDropListItem()
        {
            List<Models.Order> MapResult = new List<Order>();
            DataTable dt = new DataTable();
            String sql = @"SELECT ShipperID,CompanyName AS ShipperName
                           FROM [Sales].[Shippers]";

            using (SqlConnection conn = new SqlConnection(this.GetDBConnectionString()))
            {
                conn.Open();
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataAdapter sqlAdapter = new SqlDataAdapter(cmd);
                sqlAdapter.Fill(dt);
                conn.Close();
            }

            foreach (DataRow Row in dt.Rows)
            {
                MapResult.Add(new Order()
                {
                    ShipperId = (int)Row["ShipperID"],
                    ShipperName = Row["ShipperName"].ToString()
                }
                );
            }
            return MapResult;
        }

        /// <summary>
        /// 暫無用到的map
        /// </summary>
        /// <param name="OrderData"></param>
        /// <returns></returns>
        public List<Models.Order> MapOrderDropListData(DataTable OrderData)
        {
            List<Models.Order> MapResult = new List<Order>();
            foreach (DataRow Row in OrderData.Rows)
            {
                MapResult.Add(new Order()
                {
                    CustId = (int)Row["CustomerID"],
                    CompanyName = Row["CompanyName"].ToString(),
                    EmpId = (int)Row["EmployeeID"],
                    EmpName = Row["EmpName"].ToString(),
                    Freight = (int)Row["Freight"],
                    OrderDate = Row["OrderDate"].ToString(),
                    OrderId = (int)Row["OrderID"],
                    RequireDate = Row["RequireDate"].ToString(),
                    ShipAddress = Row["ShipAddress"].ToString(),
                    ShipCity = Row["ShipCity"].ToString(),
                    ShipCountry = Row["ShipCountry"].ToString(),
                    ShipName = Row["ShipName"].ToString(),
                    ShippedDate = Row["ShippedDate"].ToString(),
                    ShipperId = (int)Row["ShipperID"],
                    ShipperName = Row["ShipperName"].ToString(),
                    ShipPostalCode = Row["ShipPostalCode"].ToString(),
                    ShipRegion = Row["ShipRegion"].ToString()
                }
                );
            }
            return MapResult;
        }
    }
}