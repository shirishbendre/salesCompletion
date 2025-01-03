using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;

namespace SalesOrderCompletion
{
    public class OrderCompletion
    {
        public string strPurchaseModuleName = System.Configuration.ConfigurationManager.AppSettings["OrderCompletionModuleName"];       
        public static bool boolExexuteOnlyOnce = true;

        public int orderCompletion(int intOrderId, string strCompletedByFlag)
        {

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                        "OrderCompletion",
                        "SalesOrderCompletion Dll code starts",
                        null, strPurchaseModuleName);

            int intOrderCompletionResult = 0;                     
            try
            {
                if (!String.IsNullOrEmpty(strCompletedByFlag))
                {
                    strCompletedByFlag = strCompletedByFlag.ToUpper();
                    if (strCompletedByFlag == "REALTIME" || strCompletedByFlag == "OFFLINE")
                    {
                        GetConnection getProductList = new GetConnection();
                        
                        try
                        {
                            if (strCompletedByFlag == "REALTIME")
                            {
                                int intResult = CheckOrderIdPresentInwmCompletedOrders(intOrderId);

                                if (intResult != 1)
                                {
                                    DataTable dtOrderPlaced = getProductList.GetOrderPlacedDetailforOrderId(intOrderId);
                                    if (dtOrderPlaced.Rows.Count > 0)
                                    {                                       
                                        getProductList.SendMailOrderPlacedReceiptToAdmin(Convert.ToInt32(dtOrderPlaced.Rows[0]["OrderId"]));
                                      
                                    }
                                    else
                                    {
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                                                    "OrderCompletion",
                                                    "No record found for Order Id : " + intOrderId +
                                                    " in wm_OrderReceiptSend table to send order placed receipt to admin",
                                                    null, strPurchaseModuleName);

                                    }

                                    string strsSingleOrderCompletionResult = getProductList.singleOrderCompletion(intOrderId, strCompletedByFlag);

                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning,
                                                "OrderCompletion",
                                                "For OrderId : " + intOrderId + " " + strsSingleOrderCompletionResult + "",
                                                null, strPurchaseModuleName);

                                    intOrderCompletionResult = CheckOrderIdPresentInwmCompletedOrders(intOrderId);
                                }
                                else
                                {
                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning,
                                                 "OrderCompletion",
                                                 "Order ID : " + intOrderId +
                                                 " is not taken for processing because it is already present in WM_CompletedOrders table",
                                                 null, strPurchaseModuleName);
                                }
                            }
                            else
                            {
                                if (boolExexuteOnlyOnce)
                                {
                                    DataTable dtOrderPlaced = getProductList.GetOrderPlacedDetail();
                                    if (dtOrderPlaced.Rows.Count > 0)
                                    {
                                        boolExexuteOnlyOnce = false;
                                        // This loop for getting all order which placed on storefront.
                                        for (int orderCount = 0; orderCount < dtOrderPlaced.Rows.Count; orderCount++)
                                        {
                                            getProductList.SendMailOrderPlacedReceiptToAdmin(Convert.ToInt32(dtOrderPlaced.Rows[orderCount]["OrderId"]));
                                        }
                                    }
                                    else
                                    {
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                                                    "OrderCompletion",
                                                    "No record found in wm_OrderReceiptSend table to send order placed receipt to admin",
                                                    null, strPurchaseModuleName);
                                    }
                                }

                                string strsSingleOrderCompletionResult = getProductList.singleOrderCompletion(intOrderId, strCompletedByFlag);

                            
                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning,
                                                 "OrderCompletion",
                                                 "For OrderId : "+ intOrderId + " " + strsSingleOrderCompletionResult + "",
                                                 null, strPurchaseModuleName);
                                intOrderCompletionResult = CheckOrderIdPresentInwmCompletedOrders(intOrderId);

                            }

                        }
                        catch (Exception ex)
                        {
                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error,
                                        "SendMailOrderPlacedReceiptToAdmin",
                                        "SendMailOrderPlacedReceiptToAdmin region catch block. Error: " + ex.Message,
                                        ex, strPurchaseModuleName);
                        }

                    }
                    else
                    {
                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning,
                                "OrderCompletion",
                                "Invalid Completed By Flag value passed.It should be either Realtime or Offline . strCompletedByFlag value is : " + strCompletedByFlag,
                                null, strPurchaseModuleName);
                    }
                }
                else
                {
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning,
                                "OrderCompletion",
                                "Completed By Flag value empty or null . strCompletedByFlag value is : " + strCompletedByFlag,
                                null, strPurchaseModuleName);

                }                               

            }
            catch (Exception ex)
            {
                intOrderCompletionResult = 0;
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "RealTimeOrderCompletion",
                            "In OrderCompletion methods Catch block. Error: " + ex.Message,
                            ex, strPurchaseModuleName);
            }

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                       "OrderCompletion",
                       "SalesOrderCompletion Dll code Ends",
                       null, strPurchaseModuleName);

            return intOrderCompletionResult;
        }


        private static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }


        /// <summary>
        /// Check Order Id Present In wm_CompletedOrders table.if yes return 1 else return 0
        /// </summary>
        /// <param name="intOrderId"></param>
        /// <returns></returns>
        private int CheckOrderIdPresentInwmCompletedOrders(int intOrderId)
        {
            string strGetconnectionString = ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString.ToString();
            SqlConnection thisConnection = new SqlConnection(strGetconnectionString);
            int intReturnValueId = 0;

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                           "OrderCompletion.CheckOrderIdPresentInwmCompletedOrders",
                           "Checking Order Id : " + intOrderId + " Present In WM_CompletedOrders table",
                           null, strPurchaseModuleName);

            try
            {
                SqlCommand cmdd = new SqlCommand("pr_uva_CheckOrderIdPresentInwmCompletedOrders", thisConnection);
                cmdd.CommandType = CommandType.StoredProcedure;
                cmdd.Parameters.AddWithValue("@OrderId", intOrderId);
                thisConnection.Open();
                intReturnValueId = Convert.ToInt32(cmdd.ExecuteScalar());

                if (intReturnValueId == 1)
                {
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success,
                                "OrderCompletion.CheckOrderIdPresentInwmCompletedOrders",
                                "Order having OrderId =" + intOrderId +
                                " is already Present in wm_CompletedOrders table ",
                                null, strPurchaseModuleName);
                }
                else
                {
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success,
                                "OrderCompletion.CheckOrderIdPresentInwmCompletedOrders",
                                "Order having OrderId =" + intOrderId +
                                " is not Present in wm_CompletedOrders table ",
                                null, strPurchaseModuleName);
                }
                thisConnection.Close();
                return intReturnValueId;
            }
            catch (Exception ex)
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error,
                            "OrderCompletion.CheckOrderIdPresentInwmCompletedOrders",
                             ex.Message, ex, strPurchaseModuleName);

                return intReturnValueId;
            }
            finally
            {
                thisConnection.Close();
            }
        }
    }
}
