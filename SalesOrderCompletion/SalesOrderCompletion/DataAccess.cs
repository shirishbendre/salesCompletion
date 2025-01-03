using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Web;

/// <summary>
/// Summary description for DataAccess
/// </summary>
public class DataAccess
{

    //GetConnection _getConnection = new GetConnection();
    string connection = ConfigurationManager.ConnectionStrings["connectionString"].ConnectionString;

    #region Insert
    /// <summary> 
    /// Inserts a download 
    /// </summary> 
    /// <param name="download">Download</param> 
    public int InsertDownload(Download download)
    {
        if (download == null)
            throw new ArgumentNullException("download");

        try
        {
            int returnVal = 0;
            //Get ConnectionString               
            String ConnectionString = GetConnection.GetconnectionString;
            using (SqlConnection myConnection = new SqlConnection(ConnectionString))
            {
                SqlCommand myCommand = new SqlCommand("pr_uva_InsertDownload", myConnection);
                myCommand.CommandType = CommandType.StoredProcedure;


                myCommand.Parameters.AddWithValue("@DownloadGuid", download.DownloadGuid);
                myCommand.Parameters.AddWithValue("@UseDownloadUrl", download.UseDownloadUrl);
                myCommand.Parameters.AddWithValue("@DownloadUrl", download.DownloadUrl);
                myCommand.Parameters.AddWithValue("@DownloadBinary", download.DownloadBinary);
                myCommand.Parameters.AddWithValue("@ContentType", download.ContentType);
                myCommand.Parameters.AddWithValue("@Filename", download.Filename);
                myCommand.Parameters.AddWithValue("@Extension", download.Extension);
                myCommand.Parameters.AddWithValue("@IsNew", download.IsNew);

                //-- Command Time Out 
                myCommand.CommandTimeout = 200000; // ApplicationSetting.GetSQLCommandTimeOut();

                //Open the Database Connection and Execute Stored Procedure 
                myConnection.Open();
                returnVal = Convert.ToInt32(myCommand.ExecuteScalar());
                myCommand.ResetCommandTimeout();
                //Close Connection 
                myConnection.Close();
            }
            return returnVal;
        }
        catch (Exception ex)
        {
            string strex = ex.ToString();
            return 0;
        }

    }

    #endregion



}