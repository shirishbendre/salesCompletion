using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SalesOrderCompletion;
namespace SalesCompletionUsingCommonDll
{
    class GetPendingProducts
    {
        public static string strPurchaseModuleName = System.Configuration.ConfigurationManager.AppSettings["OrderCompletionModuleName"];
        public static string cs = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString.ToString();
       
        public static SqlConnection sqlConnect = new SqlConnection(cs);

        static void Main(string[] args)
        {            
           try
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetPendingProducts",
                                  "*** Sales Completion Process Starts ***", null, strPurchaseModuleName);

                Console.WriteLine("Sales Completion Process Starts.");

                //Collect Incomplete order from Order Table  
                DataTable dtOrder = GetPendingProducts.GetInCompleteOrders();

                #region If Order Count >0
                //If order count greater than 0 the proceed.
                if (dtOrder.Rows.Count > 0)
                {
                    // This loop for getting all order.
                    for (int orderCount = 0; orderCount < dtOrder.Rows.Count; orderCount++)
                    {
                        int intOrderId = Convert.ToInt32(dtOrder.Rows[orderCount]["Id"]);
                        int intCustomerId = Convert.ToInt32(dtOrder.Rows[orderCount]["CustomerId"]);

                        OrderCompletion OrderCompletion = new OrderCompletion();
                        int intResult = OrderCompletion.orderCompletion(intOrderId, "OFFLINE");

                    }
                }
                else
                {
                    Console.WriteLine("No order for processing.............");

                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetPendingProducts",
                                "No order for processing.", null, strPurchaseModuleName);
                }

                #endregion
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetPendingProducts",
                               "*** Sales Completion Process End ***", null, strPurchaseModuleName);

                Console.WriteLine("Sales Completion Process End.");
            }
            catch (Exception ex)
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetPendingProducts",
                                            ex.Message,
                                            null, strPurchaseModuleName);
            }

        }

        public static DataTable GetInCompleteOrders()
        {
            DataTable dt = new DataTable();
            try
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "callWebservice.GetInCompleteOrders",
                          "executing Stored Procedure:pr_uva_GetOrderData", null, strPurchaseModuleName);

                SqlCommand cmdd = new SqlCommand("pr_uva_GetOrderData", sqlConnect);
                cmdd.CommandType = CommandType.StoredProcedure;

                SqlDataAdapter da = new SqlDataAdapter(cmdd);
                da.Fill(dt);

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "callWebservice.GetInCompleteOrders",
                            "Get InComplete Orders data successfully", null, strPurchaseModuleName);
                return dt;
            }
            catch (Exception ex)
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "callWebservice.GetInCompleteOrders",
                            ex.Message,
                            null, strPurchaseModuleName);
                return dt;
            }

        }

    }
}
