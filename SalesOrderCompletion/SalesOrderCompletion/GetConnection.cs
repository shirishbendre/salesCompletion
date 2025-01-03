
//---
using System;
using System.Collections.Generic;
using System.Collections;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web.Script.Serialization;
using System.Net;
using System.Data.SqlClient;
using System.Configuration;
using System.Data;
using System.IO;
using System.IO.Compression;
using Ionic.Zip;
using System.Net.Mail;
using System.Net.Mime;
//using HtmlAgilityPack;
using System.Text.RegularExpressions;
using System.Security.Cryptography;
//using System.Drawing;
using System.Globalization;
using System.Collections.Specialized;
//using Forio;
using System.Web;
using HtmlAgilityPack;
using ForioEpicenter;
using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;  //add reference of EpicenterAPI namespace
/// <summary>
/// Summary description for GetConnection
/// </summary>
public class GetConnection
{
    #region Public Declaration
    DataAccess _dataAccess = new DataAccess();
    public SqlConnection thisConnection = new SqlConnection(GetconnectionString);
    string displayName = string.Empty;
    //Module Name
    // public string strPurchaseModuleName = System.Configuration.ConfigurationManager.AppSettings["PurchaseModuleName"];
    public string strPurchaseModuleName = System.Configuration.ConfigurationManager.AppSettings["OrderCompletionModuleName"];
    public string intSubscriptionDays = Convert.ToString(ConfigurationManager.AppSettings["SupplementVideoSubscriptionDays"]);
    public string parentFileLocation = Convert.ToString(ConfigurationManager.AppSettings["parentFileLocation"]);
    public string childFileLocation = Convert.ToString(ConfigurationManager.AppSettings["childFileLocation"]);
    public string compressedFilePath = Convert.ToString(ConfigurationManager.AppSettings["compressedFilePath"]);
    public string facultyfileLocation = Convert.ToString(ConfigurationManager.AppSettings["facultyFileLocation"]);
    public string appJSFileExtension = Convert.ToString(ConfigurationManager.AppSettings["JSFileExtention"]);
    public string appPlaylist = Convert.ToString(ConfigurationManager.AppSettings["Playlist"]);
    public string strVideoType = Convert.ToString(ConfigurationManager.AppSettings["VideoType"]);
    public String strMicrosite = Convert.ToString(ConfigurationManager.AppSettings["Microsite"]);
    public String pdfFileExt = Convert.ToString(ConfigurationManager.AppSettings["pdfFileExt"]);
    public String mediaext = Convert.ToString(ConfigurationManager.AppSettings["mediaext"]);
    public String strCase = Convert.ToString(ConfigurationManager.AppSettings["Case"]);
    public String parentfileLoc = Convert.ToString(ConfigurationManager.AppSettings["parentfileLoc"]);
    public String msFileExtention1 = Convert.ToString(ConfigurationManager.AppSettings["MicrositeFileExtention1"]);
    public String msFileExtention2 = Convert.ToString(ConfigurationManager.AppSettings["MicrositeFileExtention2"]);
    public String purchasedMicrositePath = Convert.ToString(ConfigurationManager.AppSettings["purchasedMicrositePath"]);
    public String strMyElementId = Convert.ToString(ConfigurationManager.AppSettings["myElementId"]);
    public string strstartJSPoint = Convert.ToString(ConfigurationManager.AppSettings["startJSPoint"]);
    public string strendJSPoint = Convert.ToString(ConfigurationManager.AppSettings["endJSPoint"]);
    public string strDardenFaculty = Convert.ToString(ConfigurationManager.AppSettings["userTypeFaculty"]);
    public string strNonDardenFaculty = Convert.ToString(ConfigurationManager.AppSettings["userTypeNonDardenFaculty"]);
    public string strAdministrator = Convert.ToString(ConfigurationManager.AppSettings["Administrator"]);
    public string includeFolderName = Convert.ToString(ConfigurationManager.AppSettings["IncludesFolder"]);
    public string strFacultySupport = Convert.ToString(ConfigurationManager.AppSettings["FacultySupportFileLocation"]);
    public String strConference = Convert.ToString(ConfigurationManager.AppSettings["Conference"]);


    public string strScriptTagStart = Convert.ToString(ConfigurationManager.AppSettings["scriptTagStart"]);
    public string strScriptTagEnd = Convert.ToString(ConfigurationManager.AppSettings["scriptTagEnd"]);
    public string strJSScript = Convert.ToString(ConfigurationManager.AppSettings["textJS"]);
    public string strSrc = Convert.ToString(ConfigurationManager.AppSettings["src"]);
    //public string strJWplayer = Convert.ToString(ConfigurationManager.AppSettings["jwplayer"]);
    //public string strkey = Convert.ToString(ConfigurationManager.AppSettings["key"]);
    //public string strJWplayerKey = Convert.ToString(ConfigurationManager.AppSettings["jwplayerkey"]);


    public string stronlyFaculty = Convert.ToString(ConfigurationManager.AppSettings["onlyforfaculty"]);
    public string strClickToDownload = Convert.ToString(ConfigurationManager.AppSettings["textfordownload"]);

    StringBuilder sbJwplayer = new StringBuilder();

    public string strName = string.Empty;
    public string strQuantity = string.Empty;
    public string strId = string.Empty;
    public string strAttributeDescription = string.Empty;
    public string strSimulationPath = string.Empty;
    public string strFirstName = string.Empty;
    public string strLastName = string.Empty;
    public string strEmail = string.Empty;
    public string strProductVariant = string.Empty;
    public string strAttributeXml = string.Empty;
    public string strParentGroupProductID = string.Empty;
    public string strOrderItemGuid = string.Empty;
    public int result = 0;

    //  new variable for realtimmewebservice

    public string strCompletedByFlagValue = string.Empty;


    //Configuration Values
    public string strConfigCopyRightPermission = Convert.ToString(ConfigurationManager.AppSettings["CopyrightPermissions"]);
    public string strConfigMasterHardCopy = Convert.ToString(ConfigurationManager.AppSettings["MasterHardCopy"]);
    public string strConfigStudentHardCopy = Convert.ToString(ConfigurationManager.AppSettings["StudentHardCopy"]);
    public string strConfigCoursePackLicense = Convert.ToString(ConfigurationManager.AppSettings["CoursePackLicense"]);
    public string strConfigCoursePack = Convert.ToString(ConfigurationManager.AppSettings["CoursePack"]);
    public string strConfigStudentLicense = Convert.ToString(ConfigurationManager.AppSettings["StudentLicense"]);
    public string strConfigUserLicense = Convert.ToString(ConfigurationManager.AppSettings["UserLicense"]);
    public string strConfigSimulationSession = Convert.ToString(ConfigurationManager.AppSettings["SimulationSession"]);
    public string strForioType = ConfigurationManager.AppSettings["ForioType"];
    public string strConfigSimulation = Convert.ToString(ConfigurationManager.AppSettings["Simulation"]);
    public string strTechnicalNote = Convert.ToString(ConfigurationManager.AppSettings["TechnicalNote"]);
    public string strExercise = Convert.ToString(ConfigurationManager.AppSettings["Exercise"]);
    public string strBook = Convert.ToString(ConfigurationManager.AppSettings["Book"]);
    public string strResearch = Convert.ToString(ConfigurationManager.AppSettings["Research"]);

    public string strProductVariantBook = Convert.ToString(ConfigurationManager.AppSettings["ProductVariantBook"]);


    public int intProductAttributeId = Convert.ToInt32(ConfigurationManager.AppSettings["ProductAttributeId"]);


    //Forio Epicenter values
    public string strEpicenterPublicKey = Convert.ToString(ConfigurationManager.AppSettings["EpicenterPublicKey"]);
    public string strEpicenterSecretKey = Convert.ToString(ConfigurationManager.AppSettings["EpicenterSecretKey"]);
    public string strFacultyUserRole = Convert.ToString(ConfigurationManager.AppSettings["FacultyUserRole"]);
    public string strStudentUserRole = Convert.ToString(ConfigurationManager.AppSettings["StudentUserRole"]);
    public string strLoginEpicenterUrl = Convert.ToString(ConfigurationManager.AppSettings["ForioEpicenterLoginUrl"]);

    public string strFacultyUser = Convert.ToString(ConfigurationManager.AppSettings["FacultyUser"]);
    public string strStudentUser = Convert.ToString(ConfigurationManager.AppSettings["StudentUser"]);

    string strProjectId = Convert.ToString(ConfigurationManager.AppSettings["ForioEpicenterProjectId"]);
    string strProjectFullName = Convert.ToString(ConfigurationManager.AppSettings["ForioEpicenterProjectFullName"]);
    string strSimulationPlatformName = Convert.ToString(ConfigurationManager.AppSettings["SimulationPlatform"]);
    string strEpicenterPlatformName = Convert.ToString(ConfigurationManager.AppSettings["EpicenterPlatform"]);


    #region forio simulation Common variable
    ArrayList sessions = new ArrayList();
    DataTable dtUserType = new DataTable();
    DataTable dtTarget = new DataTable();
    int hardCopyOrCopyrightPermissions = 0;
    int forioCount = 0;
    int intUserType = 0;
    DirectoryInfo myDirCopyFiles = null;
    DirectoryInfo myDirSupport = null;
    DirectoryInfo myDirFaculty = null;
    DirectoryInfo myDirSubFolder = null;
    string strFileExtension = string.Empty;
    string strCopyZIPFile = string.Empty;
    string strZipPath = string.Empty;
    string userName = string.Empty;
    string totalQuantity = string.Empty;
    string parentDirectoryPath = string.Empty;
    string childDirectoryPath = string.Empty;
    //modifying zip file name with format: uname_purchasedate_ordernumber
    string uName = string.Empty;
    string d = string.Empty;
    string strzipName = string.Empty;
    string strSlug = string.Empty;
    string strProduct = string.Empty;
    string strUniversity = string.Empty;
    string SessionId = string.Empty;
    string description = string.Empty;
    private static bool _hasRun;
    string organization = string.Empty;

    //Forio Epicenter declaration
    string strEpicenterAccount = string.Empty;
    string strSimulationPlatform = string.Empty;
    string strOAuthToken = String.Empty;

    string strStatusCode = string.Empty;

    string strGroupId = string.Empty;
    string strGroupname = string.Empty;
    string strProjectName = string.Empty;
    string strRefreshToken = string.Empty;
    string strResponseUserId = string.Empty;
    string strResponseMessagepassrecovery = string.Empty;
    string strResponseId = string.Empty;
    string strSearchResponseUserId = string.Empty;

    string strMessage = string.Empty;
    string strEmailId = string.Empty;
    string strPassword = string.Empty;
    string strIsSimulationSession = "false";

    string strGroupName = string.Empty;
    string strForioName = string.Empty;
    string strProjectFormatId = string.Empty;


    string postData = string.Empty;

    string allowUpload = "false";
    string maxUserCount = "0";
    string singleSignOnOnly = "false";
    string simulationUrl = string.Empty;
    string strProductType = string.Empty;
    //Microsite variable declaration
    string strMicrositeCopyZIPFile = string.Empty;
    DirectoryInfo myMIcroDirCopyFiles = null;
    int extentCount = 0;
    int SubscriptionCount = 0;
    int fileNotFound = 0;
    int CountOauthToken = 0;

    string strCoursePackSlugName = string.Empty;
    #endregion

    #endregion

    #region  Local Operations

    #region Product Process
    /// <summary>
    /// Below Function is used for Geting a coursepack included products.
    /// </summary>
    /// <param name="OrderId">Pass the Order Id</param>
    /// <param name="intCPId">Pass the Coursepack Id</param>
    /// <returns>Details of Coursepack Products</returns>


    /// <summary>
    /// ProductProcess function is used for both Coursepack included products or Simple Products.
    /// </summary>
    /// <param All>Here we pass parameter that we get from datatable depend on coursepack included product or simple product</param>
    /// <returns>It return 0 if it found extentcount > 0 otherwise 1 </returns>
    public Hashtable ProductProcess(int fileNotFound, string totalQuantity, string userName, DateTime createdDate, string parentFileLocation,
                string childFileLocation, DirectoryInfo myDirCopyFiles, DirectoryInfo myDirSupport, DirectoryInfo myDirFaculty,
                DirectoryInfo myDirSubFolder, string strSlug, string strFileExtension, string strProduct, string strUniversity,
                int intUserType, string parentDirectoryPath, string strCopyZIPFile, int extentCount, int intOrderId, int customerId, string strproductVariant,
                int SubscriptionCount, string strProductType, string displayName, string strzipName, int intOrderItemCount, int intParentgroupedProductId, string strProductNumber)
    {

        //Variable declaration
        string supportingFile = "";
        string extention = string.Empty;
        string universityName = string.Empty;
        string fileName = "";
        string folderName = strSlug;
        string[] lstfile;
        string[] micrositeListFileHTM;
        string[] micrositeListFileHTML;
        List<string> exten = new List<string>();
        Hashtable countHashtable = new Hashtable();

        //Variable declaration for Microsite
        string strMicrositeCopyZIPFile = string.Empty;
        DirectoryInfo myMIcroDirCopyFiles = null;
        int MSFileFoundCount = 0;

        try
        {
            fileNotFound = 0;

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
               "parentDirectoryPath: " + parentDirectoryPath + ",strSlug: " + strSlug, null, strPurchaseModuleName);

            #region  Is folder exist
            if (Directory.Exists(parentDirectoryPath))
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                            "Searching for files having extension: " + Convert.ToString(strFileExtension) + " in intOrderId: "
                            + intOrderId, null, strPurchaseModuleName);

                string extent = "*" + Convert.ToString(strFileExtension).ToLower();
                string MicrositeExtentionHTM = "*" + Convert.ToString(msFileExtention1).ToLower();
                string MicrositeExtentionHTML = "*" + Convert.ToString(msFileExtention2).ToLower();//*.htm
                //Get files existing on Location
                lstfile = Directory.GetFiles(parentDirectoryPath, "*" + Convert.ToString(strFileExtension).ToLower());

                micrositeListFileHTM = Directory.GetFiles(parentDirectoryPath, "*" + MicrositeExtentionHTM);
                micrositeListFileHTML = Directory.GetFiles(parentDirectoryPath, "*" + MicrositeExtentionHTML);

                #region file exist
                //If file exist on Location
                if (lstfile.Length != 0 || micrositeListFileHTM.Length != 0 || micrositeListFileHTML.Length != 0)
                {
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                "File/s exist on location: " + parentDirectoryPath, null, strPurchaseModuleName);
                    #region if (strProductType == strConference || strProductType == strCase || strProductType == strConfigCoursePack || strProductType == strForioType || strProductType == strConfigSimulation || strProductType == strTechnicalNote || strProductType == strExercise || strProductType == strBook || strProductType == strResearch)
                    //If ProductType is not Microsite or video 
                    if (strProductType == strConference || strProductType == strCase || strProductType == strConfigCoursePack || strProductType == strForioType || strProductType == strConfigSimulation || strProductType == strTechnicalNote || strProductType == strExercise || strProductType == strBook || strProductType == strResearch)
                    {
                        myDirSubFolder = new DirectoryInfo(strCopyZIPFile + @"\" + folderName);
                        //Create folder by slug name
                        myDirSubFolder.Create();

                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                    "Created sub folder: " + myDirSubFolder, null, strPurchaseModuleName);

                        // Use static Path methods to extract only the file name from the path. 
                        fileName = System.IO.Path.GetFileName(lstfile[0]);//JH-0046.pdf

                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                    "fileName: " + fileName, null, strPurchaseModuleName);
                        //File Extension other than JS(Video product)
                        if (Convert.ToString(strFileExtension) != appJSFileExtension &&
                            Convert.ToString(strProductVariant) != appPlaylist &&
                            strProductType != strForioType && strProductType != strMicrosite && strProductType != strConfigSimulation)
                        {
                            extentCount++; //to count total files found on physical location

                            System.IO.File.Copy(lstfile[0], myDirSubFolder.ToString() + "\\" + fileName, true);
                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                          "Copied file: " + fileName + " to folder: " + myDirSubFolder.ToString(), null, strPurchaseModuleName);
                        }

                        //Copying Supplement Files For Student except Teaching Note 
                        #region Supplement file (SF-Support)
                        if (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["userTypeStudent"]))
                        {

                            CopySupplementFiles(folderName, supportingFile, myDirSubFolder);
                            string strFileLocation = childFileLocation + strSlug;
                            string strVideoFileLocation = ConfigurationManager.AppSettings["SupportsFileLocation"];
                            InsertSupplementVideoDetails(customerId, intOrderId, intParentgroupedProductId,
                                        strFileLocation, strVideoFileLocation, strProductType);

                        }
                        #endregion

                        #region digital rigths for PDF
                        //This logic is for Digital Rights
                        if (Convert.ToString(strFileExtension).ToUpper() == ".PDF")
                        {
                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                        "Start adding DRM for file: " + fileName, null, strPurchaseModuleName);


                            if (Convert.ToString(strUniversity) != string.Empty)
                            {
                                universityName = Convert.ToString(strUniversity);
                            }

                            //Added by Dnyaneshwar patil Dated(22/2/2019)
                            //No Watermark for Producttype Conference
                            if (strProductType != strConference)
                            {

                                //Digital rights method
                                DigitalRights(myDirSubFolder.ToString() + "\\" + fileName, displayName, Convert.ToInt32(totalQuantity),
                                               createdDate.ToShortDateString(), intOrderId, universityName);

                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                            "End adding DRM for file: " + fileName, null, strPurchaseModuleName);
                            }

                        }
                        #endregion


                        //Process to get the teaching notes if user is FACULTY or Asministrator change on 4/18/2013
                        if (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["userTypeFaculty"]) ||
                           (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["userTypeNonDardenFaculty"])) ||
                           (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["Administrator"])))
                        {
                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess", "UserName:" +
                                        userName + ",UserTypeId:" + intUserType + ". Adding faculty support folder", null, strPurchaseModuleName);

                            #region Faculty support files(SF-FacultySupport)
                            // SF-FacultySupport directory.
                            myDirFaculty = new DirectoryInfo(ConfigurationManager.AppSettings["facultyFileLocation"] + folderName);
                            if (myDirFaculty.Exists)
                            {

                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess",
                                             "Start getting teaching notes. Path:  " + myDirFaculty.ToString(), null, strPurchaseModuleName);

                                string[] allSupportingFiles = System.IO.Directory.GetFiles(myDirFaculty.ToString());
                                for (int i = 0; i < allSupportingFiles.Length; i++)
                                {
                                    extentCount++;
                                    string strFile = Path.GetFullPath(allSupportingFiles[i]);
                                    string strFileNameExtension = Path.GetExtension(allSupportingFiles[i]).ToUpper();
                                    if (strFileNameExtension != ".MP4" && strFileNameExtension != ".JS")
                                    {
                                        string teachingNotes = "";
                                        //Use static Path methods to extract only the file name from the path.
                                        teachingNotes = System.IO.Path.GetFileName(strFile);

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess",
                                                 "Supporting file name: " + supportingFile, null, strPurchaseModuleName);

                                        System.IO.File.Copy(strFile, myDirSubFolder.ToString() + @"\" + teachingNotes, true);

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess",
                                            "File " + teachingNotes + " copied to location " + myDirSubFolder.ToString(), null, strPurchaseModuleName);
                                    }
                                }

                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                            "End getting teaching notes. Path:  " + myDirFaculty.ToString(), null, strPurchaseModuleName);

                            }
                            else
                            {
                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                         "Faculty support folder: " + strSlug + " does not exist on location: "
                                         + ConfigurationManager.AppSettings["facultyFileLocation"].ToString(), null, strPurchaseModuleName);
                            }
                            #endregion

                            #region Supplement file (SF-Support)
                            //To get all the Supplement files from SF-Support folder.
                            CopySupplementFiles(folderName, supportingFile, myDirSubFolder);
                            #endregion

                            #region Video Supplement for Faculty 
                            string strFileLocation = ConfigurationManager.AppSettings["facultyFileLocation"] + strSlug;
                            string strVideoFileLocation = ConfigurationManager.AppSettings["FacultySupportFileLocation"];
                            InsertSupplementVideoDetails(customerId, intOrderId, intParentgroupedProductId,
                                strFileLocation, strVideoFileLocation, strProductType);
                            #endregion

                            if (strFileExtension == appJSFileExtension || strProductType == strForioType || strProductType == strConfigSimulation)
                            {
                                extentCount++;
                            }
                        }


                    }
                    #endregion
                    #region else if Product type Microsite
                    else if (strProductType == strMicrosite)
                    {

                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                        "Purchased Product is a Microsite", null, strPurchaseModuleName);
                        // Create folder by Name_OrderDateYYMMDD_OrderNo in /Purchased Microsites
                        #region Creating folder ProductOrderNumber for Microsite
                        // It Create a folder for ProductOrderNumber with uname_purchasedate_ordernumber
                        strMicrositeCopyZIPFile = Convert.ToString(ConfigurationManager.AppSettings["MicrositeFilePath"]) + strzipName;
                        //strMicrositeCopyZIPFile =C:\inetpub\wwwroot\WMBranding\Purchased Microsites\VijayNagarsoge_191105_243562
                        if (!Directory.Exists(strMicrositeCopyZIPFile))
                        {
                            // If folder does not exist create folder
                            myMIcroDirCopyFiles = new DirectoryInfo(strMicrositeCopyZIPFile);
                            myMIcroDirCopyFiles.Create();
                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                        "Created folder on location: " + Convert.ToString(ConfigurationManager.AppSettings["MicrositeFilePath"])
                                         + " with name: " + strzipName, null, strPurchaseModuleName);
                        }
                        else
                        {
                            //Commented by Akash for Bug#3096
                            //If folder exists initialize object
                            //myMIcroDirCopyFiles = new DirectoryInfo(strMicrositeCopyZIPFile);
                            //DirectoryInfo[] dirSubInfo = myMIcroDirCopyFiles.GetDirectories();
                            //if (dirSubInfo.Count() > 0 && intOrderItemCount == 0)
                            //{

                            /* ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                     "Deleted folder on location: " + Convert.ToString(ConfigurationManager.AppSettings["MicrositeFilePath"])
                                      + " with name: " + strzipName, null, strPurchaseModuleName);
                             // delete existing folder and create new folder
                             myMIcroDirCopyFiles.Delete(true);

                             ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                    "Create folder on location: " + Convert.ToString(ConfigurationManager.AppSettings["MicrositeFilePath"])
                                     + " with name: " + strzipName, null, strPurchaseModuleName);

                             myMIcroDirCopyFiles = new DirectoryInfo(strMicrositeCopyZIPFile);
                             myMIcroDirCopyFiles.Create();*/

                            //}
                        }
                        #endregion

                        #region variable Declaration
                        // Declare Variables
                        String MicrositeProductPath = Convert.ToString(ConfigurationManager.AppSettings["parentFileLocation"]) +
                                                                    strSlug;
                        String[] fileContent;
                        String htmlFileName;
                        int micrositeProductId = intParentgroupedProductId;
                        String micrositeName = strSlug;
                        string strmicrositeProdSlugFolder = string.Empty;
                        //String includeFolderName = Convert.ToString(ConfigurationManager.AppSettings["IncludesFolder"]);
                        String includeFolderPath = strMicrositeCopyZIPFile + @"/" + micrositeName;
                        //includeFolderPath =C:\inetpub\wwwroot\WMBranding\Purchased Microsites\VijayNagarsoge_191105_243562 + @"/" + "micosite-the-history-of-adidas";
                        DirectoryInfo micrositeProdSlugFolder = null;
                        DirectoryInfo micrositeIncludesFolder = null;
                        DirectoryInfo micrositeChildProductFolder = null;
                        #endregion

                        #region create folder for productslug and includes folder


                        micrositeProdSlugFolder = new DirectoryInfo(strMicrositeCopyZIPFile + @"/" + micrositeName);
                        strmicrositeProdSlugFolder = Convert.ToString(micrositeProdSlugFolder);
                        micrositeIncludesFolder = new DirectoryInfo(includeFolderPath + @"/" + includeFolderName);
                        //It delete folder if same slug name folder exist in 'micrositeProdSlugFolder' folder.

                        if (Directory.Exists(strmicrositeProdSlugFolder))
                        {
                            //If folder exists initialize object
                            //micrositeProdSlugFolder = new DirectoryInfo(strmicrositeProdSlugFolder);
                            DirectoryInfo[] dirSubInfo = micrositeProdSlugFolder.GetDirectories();
                            if (dirSubInfo.Count() > 0)
                            {

                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                     "Deleted folder on location: " + Convert.ToString(ConfigurationManager.AppSettings["MicrositeFilePath"])
                                      + " with name: " + micrositeProdSlugFolder, null, strPurchaseModuleName);
                                // delete existing folder and create new folder
                                micrositeProdSlugFolder.Delete(true);
                            }
                            micrositeProdSlugFolder = new DirectoryInfo(strmicrositeProdSlugFolder);
                        }

                        //Create folder for microsite Purchased Microsites/Order/Slug
                        micrositeProdSlugFolder.Create();
                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                    "Created Slug Folder for Microsite" + "Microsite Name: " + micrositeName, null, strPurchaseModuleName);

                        //Create folder Includes Inside Purchased Microsites/Order/Slug/Includes                          
                        micrositeIncludesFolder.Create();
                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                    "Created Includes Folder for Microsite" + "Microsite Name: " + micrositeName, null, strPurchaseModuleName);
                        #endregion

                        // Copy .HTML to Purchased Microsites/Order/Slug/
                        if (Directory.Exists(MicrositeProductPath))
                        {
                            // Copy .html file from Slug folder to Purchased Microsites/Order/Slug
                            fileContent = Directory.GetFiles(parentDirectoryPath);
                            htmlFileName = System.IO.Path.GetFileName(fileContent[0]);
                            // check file with extention .html ot .htm
                            #region Copy .HTML file
                            if ((htmlFileName.Contains(msFileExtention1)) || (htmlFileName.Contains(msFileExtention2)))
                            {
                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                    ">HTML file found for Microsite" + "HTML File Name: " + htmlFileName, null, strPurchaseModuleName);

                                // Replace <base/> tag in HTML file
                                string word = "<base/>";
                                string BasePath = purchasedMicrositePath + strzipName + "/" + strSlug + "/" + includeFolderName + "/";
                                string replacement = "<base href=" + @"""" + BasePath + @"""/>";
                                string directory = MicrositeProductPath + @"\";
                                string destination = micrositeProdSlugFolder + @"\";
                                string saveFileName = htmlFileName;
                                string elementId = strMyElementId;
                                string replaceProdNumber = "";

                                // Use method to replace tag in .HTML file.
                                EditPaths(htmlFileName, word, replacement, elementId, replaceProdNumber, saveFileName, directory, destination, intUserType, strProductNumber);

                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                "HTML document edited", null, strPurchaseModuleName);
                                MSFileFoundCount++;
                                //System.IO.File.Copy(fileContent[0], micrositeProdSlugFolder.ToString() + "\\" + htmlFileName, true);
                            }
                            else
                            {
                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.singleOrderCompletion",
                                    "File : " + "Is not a html file.", null, strPurchaseModuleName);

                            }
                            #endregion

                            #region copy Including files for microsite
                            try
                            {
                                // Check if microsite Purchased Microsites/Order/Slug/Includes folder exists
                                if (micrositeIncludesFolder.Exists)
                                {
                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                    "Includes folder exist for Microsite on location: " + micrositeIncludesFolder, null, strPurchaseModuleName);

                                    string[] allMicrositeFiles = System.IO.Directory.GetFiles(MicrositeProductPath.ToString() + @"\" + includeFolderName);
                                    string includingFiles = "";
                                    // Copy the files and overwrite destination files if they already exist. 
                                    foreach (string strFile in allMicrositeFiles)
                                    {
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                            "Getting Microsite File inside Includes Folder" + "File Name: " + strFile, null, strPurchaseModuleName);
                                        // Use static Path methods to extract only the file name from the path.
                                        includingFiles = System.IO.Path.GetFileName(strFile);
                                        // copy file to destination directory
                                        System.IO.File.Copy(strFile, micrositeIncludesFolder.ToString() + @"\" + includingFiles, true);

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                           "Copied" + "File Name: " + strFile + micrositeIncludesFolder, null, strPurchaseModuleName);
                                    }
                                    if (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["userTypeFaculty"]) ||
                                    (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["userTypeNonDardenFaculty"])) ||
                                     (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["Administrator"])))
                                    {
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess", "UserName:" +
                                                    userName + ",UserTypeId:" + intUserType + ". Adding faculty support folder", null, strPurchaseModuleName);

                                        #region Faculty support files(SF-FacultySupport)
                                        // SF-FacultySupport directory.
                                        myDirFaculty = new DirectoryInfo(ConfigurationManager.AppSettings["facultyFileLocation"] + micrositeName);
                                        if (myDirFaculty.Exists)
                                        {
                                            DirectoryInfo strMicrositeMainSupportingFolder;
                                            strMicrositeMainSupportingFolder = new DirectoryInfo(micrositeIncludesFolder + @"/" + strProductNumber);
                                            strMicrositeMainSupportingFolder.Create();

                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess",
                                                         "Start getting teaching notes. Path:  " + myDirFaculty.ToString(), null, strPurchaseModuleName);

                                            string[] allSupportingFiles = System.IO.Directory.GetFiles(myDirFaculty.ToString());
                                            string teachingNotes = "";
                                            // Copy the files and overwrite destination files if they already exist. 
                                            foreach (string strsuportFile in allSupportingFiles)
                                            {
                                                //extentCount++; //to count total files found on physical location
                                                // Use static Path methods to extract only the file name from the path.
                                                teachingNotes = System.IO.Path.GetFileName(strsuportFile);
                                                System.IO.File.Copy(strsuportFile, strMicrositeMainSupportingFolder.ToString() + @"\" + teachingNotes, true);

                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                "File: " + strsuportFile + "copied to Location: " + micrositeIncludesFolder, null, strPurchaseModuleName);
                                            }
                                            string zipSupportFilePath = strMicrositeMainSupportingFolder + ".zip";
                                            String[] zipSupportFileContent = Directory.GetFiles(strMicrositeMainSupportingFolder.ToString());
                                            using (ZipFile zip = new ZipFile())
                                            {

                                                if (zipSupportFileContent.Length > 0)
                                                {
                                                    ////zip.Password = DecryptMaskedCard(Convert.ToString(dt.Rows[pendingOrderCount]["Password"]));
                                                    // Zip product folder in Purchased Microsites/Order/Slug/Includes
                                                    zip.AddDirectory(strMicrositeMainSupportingFolder.ToString());
                                                    zip.Save(zipSupportFilePath);

                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                    "Zipfile created  " + "to Location: " + zipSupportFilePath, null, strPurchaseModuleName);
                                                }
                                                else
                                                {
                                                    //Clear folder created for ptoduct
                                                    ClearFolder(strMicrositeMainSupportingFolder.ToString());
                                                    strMicrositeMainSupportingFolder.Delete();

                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                   "Directory cleared for product in microsite" + "on Location: " + strMicrositeMainSupportingFolder, null, strPurchaseModuleName);
                                                }
                                            }
                                            DirectoryInfo productSupportDirectory = new DirectoryInfo(strMicrositeMainSupportingFolder.ToString());
                                            if (productSupportDirectory.Exists)
                                            {
                                                // Clear existing folder for product number.
                                                ClearFolder(strMicrositeMainSupportingFolder.ToString());
                                                strMicrositeMainSupportingFolder.Delete();

                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess",
                                                          "Folder deleted from: " + strMicrositeMainSupportingFolder, null, strPurchaseModuleName);
                                            }
                                        }
                                        else
                                        {
                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                                     "Faculty support folder: " + strSlug + " does not exist on location: "
                                                     + ConfigurationManager.AppSettings["facultyFileLocation"].ToString(), null, strPurchaseModuleName);
                                        }
                                        #endregion
                                    }
                                    else
                                    {

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess",
                                                      "User is not Authorised to get Supporting content for Microsites.", null, strPurchaseModuleName);
                                    }
                                    #region Get Products related to Microsite
                                    try
                                    {
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                        "Get Products related to microsite", null, strPurchaseModuleName);

                                        // Get Microsite Product Details, Used stored procedure to get details of products included in Microsite
                                        DataTable dtMicrositeProductDetails = GetMicrositeProductDetails(micrositeProductId);
                                        //Count of totalorderitems
                                        int intTotalOrderItems = dtMicrositeProductDetails.Rows.Count;

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                        intTotalOrderItems + " are part of Microsite", null, strPurchaseModuleName);

                                        //Stringbuilder 
                                        sbJwplayer.Append(strstartJSPoint);
                                        //If record fetched for details regrding products included in microsite
                                        if (dtMicrositeProductDetails != null)
                                        {
                                            if (dtMicrositeProductDetails.Rows.Count > 0)
                                            {
                                                string strCommonJSdestination = string.Empty;
                                                string strcommonJSfilename = string.Empty;
                                                // For each product in Microsite 
                                                for (int orderDetailsCount2 = 0; orderDetailsCount2 < dtMicrositeProductDetails.Rows.Count; orderDetailsCount2++)
                                                {
                                                    string micrositechilfProductSlug = Convert.ToString(dtMicrositeProductDetails.Rows[orderDetailsCount2]["ProductSlug"]);
                                                    string micrositeChildPath = Convert.ToString(ConfigurationManager.AppSettings["parentFileLocation"]) +
                                                                                    micrositechilfProductSlug;


                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                    "Get details for product " + " which is part of Microsite", null, strPurchaseModuleName);
                                                    // Check if Product exists on Physical location on SF-Products directory.
                                                    if (Directory.Exists(micrositeChildPath))
                                                    {
                                                        // Variable declaration
                                                        String childProductNumber = Convert.ToString(dtMicrositeProductDetails.Rows[orderDetailsCount2]["ProductNumber"]);
                                                        String[] childFileContent;
                                                        String childFileName = "";
                                                        String strUniversityName;
                                                        String microSitePdfExtention = pdfFileExt;
                                                        String childProdtype = Convert.ToString(dtMicrositeProductDetails.Rows[orderDetailsCount2]["ProductType"]);

                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                        "Directory Exist for product:" + childFileName + " In SF-Products, which is part of Microsite", null, strPurchaseModuleName);

                                                        micrositeChildProductFolder = new DirectoryInfo(micrositeIncludesFolder + @"/" + childProductNumber);

                                                        #region Case included in Microsite
                                                        // Only case ie. PDF can be a part of microsite
                                                        // Check if product type is case 
                                                        if (childProdtype == strCase)
                                                        {
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                            "Product type is case so get PDF file from Location", null, strPurchaseModuleName);

                                                            // Create folder with product number in Purchased Microsites/Order/Slug/Includes directory
                                                            micrositeChildProductFolder.Create();

                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                            "Folder created by Product number in Includes folder. On path: " + micrositeChildProductFolder, null, strPurchaseModuleName);

                                                            // Copy .PDF file from Slug folder to Purchased Microsites/Order/Slug/Includes/ProductNo.
                                                            childFileContent = Directory.GetFiles(micrositeChildPath, "*" + Convert.ToString(pdfFileExt).ToLower());
                                                            childFileName = System.IO.Path.GetFileName(childFileContent[0]);

                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                           "Get files with .pdf extention for product from path: " + micrositeChildPath, null, strPurchaseModuleName);

                                                            //copy .PDF file to directory.
                                                            System.IO.File.Copy(childFileContent[0], micrositeChildProductFolder.ToString() + "\\" + childFileName, true);
                                                            #region digital rigths for PDF
                                                            //This logic is for Digital Rights
                                                            if (Convert.ToString(pdfFileExt).ToUpper() == ".PDF")
                                                            {
                                                                if (Convert.ToString(strUniversity) != string.Empty)
                                                                {
                                                                    strUniversityName = Convert.ToString(strUniversity);




                                                                    DigitalRights(micrositeChildProductFolder.ToString() + "\\" + childFileName, displayName, Convert.ToInt32(totalQuantity),
                                                                               createdDate.ToShortDateString(), intOrderId, strUniversityName);
                                                                    SubscriptionCount++;
                                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                   "Digital rights given to : " + childFileName, null, strPurchaseModuleName);

                                                                }
                                                                else
                                                                {
                                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                    "University name is null for : " + displayName, null, strPurchaseModuleName);

                                                                }
                                                            }
                                                            else
                                                            {
                                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                "File extention is not PDF" + displayName, null, strPurchaseModuleName);
                                                            }
                                                            #endregion
                                                        }
                                                        else
                                                        {
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                            "Product type is Video", null, strPurchaseModuleName);

                                                            // Product is a video product 
                                                            // Create directory by Purchased Microsites/Order/Slug/Includes/ProductNo
                                                            micrositeChildProductFolder.Create();

                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                            "Directory created by product number on location: " + micrositeChildProductFolder, null, strPurchaseModuleName);

                                                            childFileContent = Directory.GetFiles(micrositeChildPath);

                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                           "Files collected from physical location from path: " + micrositeChildPath, null, strPurchaseModuleName);

                                                            int chilfFileContentCount = 0;
                                                            foreach (string strFileList in childFileContent)
                                                            {
                                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                "Directory created by product number on location: " + micrositeChildProductFolder, null, strPurchaseModuleName);

                                                                childFileName = System.IO.Path.GetFileName(strFileList);
                                                                // Check if File is JavaScript for JW Player which plays Videos.
                                                                if (childFileName.Contains(Convert.ToString(appJSFileExtension).ToLower()))
                                                                {
                                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                    "File is a java script File name: " + childFileName, null, strPurchaseModuleName);

                                                                    childFileName = System.IO.Path.GetFileName(strFileList);
                                                                    string word = parentfileLoc + Convert.ToString(dtMicrositeProductDetails.Rows[orderDetailsCount2]["ProductSlug"]) + "/";
                                                                    string replacement = includeFolderName + @"/";
                                                                    string directory = micrositeChildPath + @"\";
                                                                    string destination = micrositeIncludesFolder + @"/";
                                                                    string saveFileName = childFileName;
                                                                    string replaceProdNumber = childProductNumber;
                                                                    string elementId = "myElement";
                                                                    strcommonJSfilename = saveFileName;
                                                                    strCommonJSdestination = destination;
                                                                    // Replace the source path in JS and save to Purchased Microsites/Order/Slug/Includes 
                                                                    EditPaths(childFileName, word, replacement, elementId, replaceProdNumber, saveFileName, directory, destination, intUserType, strProductNumber);

                                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                    "Java script edited file path replaced and saved succesfully to directory " + destination, null, strPurchaseModuleName);
                                                                }
                                                                else
                                                                {
                                                                    // File is not js and can be copied to Purchased Microsites/Order/Slug/Includes
                                                                    System.IO.File.Copy(childFileContent[chilfFileContentCount], micrositeIncludesFolder.ToString() + @"/" + childFileName, true);

                                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                    "File: " + childFileName + "related to video, copied to physical location ", null, strPurchaseModuleName);

                                                                }
                                                                chilfFileContentCount++;
                                                                SubscriptionCount++;
                                                            }
                                                        }
                                                        #endregion

                                                        #region Supporting files for cases if user is faculty or Administrator
                                                        // Supporting files for Products in Microsites to be zipped 
                                                        string micrositeFacultySupportFolder = Convert.ToString(ConfigurationManager.AppSettings["facultyFileLocation"]) +
                                                                                               Convert.ToString(dtMicrositeProductDetails.Rows[orderDetailsCount2]["ProductSlug"]);
                                                        string micrositeSupportFolder = Convert.ToString(ConfigurationManager.AppSettings["childFileLocation"]) +
                                                                                        Convert.ToString(dtMicrositeProductDetails.Rows[orderDetailsCount2]["ProductSlug"]);
                                                        string childProductName = Convert.ToString(dtMicrositeProductDetails.Rows[orderDetailsCount2]["ProductSlug"]);

                                                        //If user role is FACULTY, NON-Darden FACULTY or Administrator
                                                        if (Directory.Exists(micrositeFacultySupportFolder) &&
                                                            (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["userTypeFaculty"]) ||
                                                            (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["userTypeNonDardenFaculty"])) ||
                                                            (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["Administrator"]))))
                                                        {
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                    "SF-Faculty Support folder exist for product " + childFileName + "related to video, copied to physical location ", null, strPurchaseModuleName);

                                                            string[] childFacultySupportingContent = System.IO.Directory.GetFiles(micrositeFacultySupportFolder.ToString());
                                                            string childFacultysupportIncludingFiles = "";

                                                            // Copy the files and overwrite destination files if they already exist. 
                                                            foreach (string strFile in childFacultySupportingContent)
                                                            {
                                                                // Use static Path methods to extract only the file name from the path.
                                                                childFacultysupportIncludingFiles = System.IO.Path.GetFileName(strFile);
                                                                System.IO.File.Copy(strFile, micrositeChildProductFolder.ToString() + @"/" + childFacultysupportIncludingFiles, true);

                                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                    "Supporting files copied for Product: " + childFileName + "related to video, copied to physical location ", null, strPurchaseModuleName);
                                                            }
                                                        }
                                                        #endregion

                                                        #region Supporting file for product in SF-Support folder.
                                                        //If Product folder exist in SF-Faculty Support
                                                        if (Directory.Exists(micrositeSupportFolder))
                                                        {
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                            "Supporting files copied for Product: " + childFileName + "related to video, copied to physical location from SF-Support ", null, strPurchaseModuleName);

                                                            string[] childsupportingContent = System.IO.Directory.GetFiles(micrositeSupportFolder.ToString());
                                                            string childSupportIncludingFiles = "";
                                                            foreach (string strSupportFile in childsupportingContent)
                                                            {
                                                                // Use static Path methods to extract only the file name from the path.
                                                                childSupportIncludingFiles = System.IO.Path.GetFileName(strSupportFile);
                                                                System.IO.File.Copy(strSupportFile, micrositeChildProductFolder.ToString() + @"/" + childSupportIncludingFiles, true);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                                                                  "Product Does not have supporting files in SF-Support folder" + "Product Name: " + childProductName, null, strPurchaseModuleName);
                                                        }
                                                        #endregion

                                                        string zipFilePath = micrositeChildProductFolder + ".zip";
                                                        String[] zipFileContent = Directory.GetFiles(micrositeChildProductFolder.ToString());
                                                        using (ZipFile zip = new ZipFile())
                                                        {
                                                            try
                                                            {
                                                                if (zipFileContent.Length > 0)
                                                                {
                                                                    ////zip.Password = DecryptMaskedCard(Convert.ToString(dt.Rows[pendingOrderCount]["Password"]));
                                                                    // Zip product folder in Purchased Microsites/Order/Slug/Includes
                                                                    zip.AddDirectory(micrositeChildProductFolder.ToString());
                                                                    zip.Save(zipFilePath);

                                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                    "Zip file saved to location: " + zipFilePath + "related to video ", null, strPurchaseModuleName);
                                                                }
                                                                else
                                                                {
                                                                    //Clear folder created by Product Number for product in Microsite
                                                                    ClearFolder(micrositeChildProductFolder.ToString());
                                                                    micrositeChildProductFolder.Delete();

                                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                    "Directory deletd from path " + micrositeChildProductFolder + "related to video ", null, strPurchaseModuleName);
                                                                }
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.singleOrderCompletion",
                                                                            "Error: " + ex.Message, null, strPurchaseModuleName);
                                                            }
                                                        }
                                                        DirectoryInfo productDirectory = new DirectoryInfo(micrositeChildProductFolder.ToString());
                                                        if (productDirectory.Exists)
                                                        {
                                                            // Clear existing folder for product number.
                                                            ClearFolder(micrositeChildProductFolder.ToString());
                                                            micrositeChildProductFolder.Delete();

                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess",
                                                                      "Folder deleted from: " + micrositeChildProductFolder, null, strPurchaseModuleName);
                                                        }


                                                    }
                                                    else
                                                    {
                                                        fileNotFound++;

                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess", "Parent folder "
                                                                    + strSlug.ToString() + " does not exist on physical location: " + parentFileLocation, null, strPurchaseModuleName);

                                                        Console.WriteLine("Product Included in Microsite " + micrositechilfProductSlug.ToString() + " does not exist on physical location: " + micrositeChildPath);
                                                        //Send email to administrator
                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess",
                                                                   "Call method sendMailToAdministratorFileNotFound, Parameter: strProduct =" + strProduct + ",parentDirectoryPath=" +
                                                                        parentDirectoryPath + ", intOrderId =" + intOrderId + ", strSlug" + micrositechilfProductSlug, null, strPurchaseModuleName);

                                                        //Send email to administrator
                                                        sendMailToAdministratorFileNotFound(Convert.ToString(micrositechilfProductSlug), micrositeChildPath, intOrderId, Convert.ToString(micrositechilfProductSlug));

                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess",
                                                                      "Method sendMailToAdministratorFileNotFound excuted successfully.", null, strPurchaseModuleName);
                                                    }
                                                }

                                                #region common JS
                                                // String builer to build the javascript for single or multiple Video products.
                                                using (StreamWriter writer = new StreamWriter(strCommonJSdestination + strProductNumber + (Convert.ToString(appJSFileExtension).ToLower()), true))
                                                {

                                                    sbJwplayer.Append(strendJSPoint);
                                                    writer.Write(sbJwplayer.ToString());
                                                    sbJwplayer.Clear();
                                                    writer.Close();
                                                }
                                                #endregion

                                            }

                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.singleOrderCompletion",
                                                    "Error: " + "Cannot process Products included in Microsite" + ex.Message, null, strPurchaseModuleName);
                                    }
                                    #endregion
                                }
                                else
                                {
                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                        "Includes Directory does not exist on physical location: " + micrositeIncludesFolder, null, strPurchaseModuleName);
                                }

                            }
                            catch (Exception ex)
                            {
                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.singleOrderCompletion",
                                            "Error: " + "Copy Including files for microsites in Microsite/Includes Folder" + ex.Message, null, strPurchaseModuleName);
                            }
                            SubscriptionCount++;
                            #endregion
                        }

                    }
                    #endregion

                    #region else if Product type playlist
                    else if (strProductType == appPlaylist || strProductType == strVideoType)
                    {
                        SubscriptionCount++;
                        #region Faculty support files(SF-FacultySupport)
                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                    "Checking if Faculty support folder: " + myDirSupport + " exists", null, strPurchaseModuleName);

                        // SF-FacultySupport directory.
                        myDirFaculty = new DirectoryInfo(ConfigurationManager.AppSettings["facultyFileLocation"] + folderName);
                        if (myDirFaculty.Exists)
                        {

                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess",
                                         "Start getting teaching notes. Path:  " + myDirFaculty.ToString(), null, strPurchaseModuleName);
                            if (myDirSubFolder == null)
                            {
                                myDirSubFolder = new DirectoryInfo(strCopyZIPFile + @"\" + folderName);

                                myDirSubFolder.Create();

                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                            "Created sub folder: " + myDirSubFolder, null, strPurchaseModuleName);
                            }
                            myDirSubFolder = new DirectoryInfo(strCopyZIPFile + @"\" + folderName);

                            myDirSubFolder.Create();
                            string[] allSupportingFiles = System.IO.Directory.GetFiles(myDirFaculty.ToString());
                            if (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["userTypeFaculty"]) ||
                               (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["userTypeNonDardenFaculty"])) ||
                               (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["Administrator"])))
                            {
                                string teachingNotes = "";
                                // Copy the files and overwrite destination files if they already exist. 
                                foreach (string strFile in allSupportingFiles)
                                {
                                    extentCount++; //to count total files found on physical location

                                    // Use static Path methods to extract only the file name from the path.
                                    teachingNotes = System.IO.Path.GetFileName(strFile);
                                    System.IO.File.Copy(strFile, myDirSubFolder.ToString() + @"\" + teachingNotes, true);

                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                                "Teaching notes file name: " + teachingNotes.ToString(), null, strPurchaseModuleName);

                                }

                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                            "End getting teaching notes. Path:  " + myDirFaculty.ToString(), null, strPurchaseModuleName);
                            }

                            //Copying Supplement Files For Student except Teaching Note 
                            #region Supplement file (SF-Support)
                            if (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["userTypeStudent"]))
                            {
                                CopySupplementFiles(folderName, supportingFile, myDirSubFolder);
                                string strFileLocation = childFileLocation + strSlug;
                                string strVideoFileLocation = ConfigurationManager.AppSettings["SupportsFileLocation"];
                                InsertSupplementVideoDetails(customerId, intOrderId, intParentgroupedProductId,
                                    strFileLocation, strVideoFileLocation, strProductType);

                            }
                            #endregion

                        }
                        else
                        {
                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                     "Faculty support folder: " + strSlug + " does not exist on location: "
                                     + ConfigurationManager.AppSettings["facultyFileLocation"].ToString(), null, strPurchaseModuleName);
                        }
                        #endregion


                    }
                    #endregion

                }
                #endregion
                else
                #region file not found
                {
                    //IF FOLDER NOT FOUND ON PHYSICAL LOCATION
                    fileNotFound++;
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: VideoDetails",
                                "File(s) DOES NOT exist on location: " + parentDirectoryPath, null, strPurchaseModuleName);

                    Console.WriteLine("File does not exist with extension " + extent + " into parent folder " + strSlug
                              + " on physical location: " + parentFileLocation);

                    //Send email to administrator
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: VideoDetails",
                               "Call method sendMailToAdministratorFileNotFound, Parameter: strProduct =" + strProduct + ",parentDirectoryPath=" +
                                    parentDirectoryPath + ", orderId =" + intOrderId + ", strSlug" + strSlug, null, strPurchaseModuleName);

                    sendMailToAdministratorFileNotFound(strProduct, parentDirectoryPath, intOrderId, strSlug);

                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: VideoDetails",
                               "Method sendMailToAdministratorFileNotFound excuted successfully.", null, strPurchaseModuleName);

                }
                #endregion
            }
            #endregion

            #region folder does not exist
            else
            {
                //IF FOLDER NOT FOUND ON PHYSICAL LOCATION
                fileNotFound++;
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess", "Parent folder "
                            + strSlug.ToString() + " does not exist on physical location: " + parentFileLocation, null, strPurchaseModuleName);
                Console.WriteLine("Parent folder " + strSlug.ToString() + " does not exist on physical location: " + parentFileLocation);

                //Send email to administrator
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess",
                           "Call method sendMailToAdministratorFileNotFound, Parameter: strProduct =" + strProduct + ",parentDirectoryPath=" +
                                parentDirectoryPath + ", orderId =" + intOrderId + ", strSlug" + strSlug, null, strPurchaseModuleName);

                //Send email to administrator
                sendMailToAdministratorFileNotFound(Convert.ToString(strProduct), parentDirectoryPath, intOrderId, Convert.ToString(strSlug));

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess",
                              "Method sendMailToAdministratorFileNotFound excuted successfully.", null, strPurchaseModuleName);
            }
            #endregion

            //Netra
            countHashtable.Add("extentCount", extentCount);
            countHashtable.Add("SubscriptionCount", SubscriptionCount);
            countHashtable.Add("fileNotFound", fileNotFound);

            return countHashtable;
        }
        catch (Exception ex)
        {

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.ProductProcess", ex.Message, ex, strPurchaseModuleName);
            return countHashtable;

        }

    }
    #endregion



    #endregion

    #region Insert


    /// <summary>
    /// 
    /// </summary>
    /// <param name="sessionId"></param>
    /// <param name="ForioSession"></param>
    /// <param name="simulationId"></param>
    /// <param name="createdUrl"></param>
    /// <param name="orderGuid"></param>
    /// <param name="EmailAddress"></param>
    /// <param name="firstname"></param>
    /// <param name="lastname"></param>
    /// <param name="role"></param>
    /// <returns></returns>
    public int Insertsimulation(string sessionId, string ForioSession, string simulationId, string createdUrl
                                        , string orderGuid, string EmailAddress, string firstname, string lastname, string userid, string role)
    {
        int _records = 0;
        try
        {
            SqlCommand myCommandSave = new SqlCommand();
            //SqlCommand myCommand = new SqlCommand("pr_UVA_GetProductIdByUrlRecord", myConnection);
            myCommandSave.CommandType = CommandType.StoredProcedure;
            myCommandSave.CommandText = "pr_uva_InsertSimulationDetails";

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.Insertsimulation",
                "Stored Procedure: pr_uva_InsertSimulationDetails ", null, strPurchaseModuleName);

            myCommandSave.Parameters.AddWithValue("@sessionId", SqlDbType.VarChar).Value = sessionId;
            myCommandSave.Parameters.AddWithValue("@ForioSession", SqlDbType.VarChar).Value = ForioSession;
            myCommandSave.Parameters.AddWithValue("@simulationId", SqlDbType.VarChar).Value = simulationId;
            myCommandSave.Parameters.AddWithValue("@createdUrl", SqlDbType.VarChar).Value = createdUrl;
            myCommandSave.Parameters.AddWithValue("@OrderProductVariantGuid", SqlDbType.VarChar).Value = orderGuid;
            myCommandSave.Parameters.AddWithValue("@EmailAddress", SqlDbType.VarChar).Value = EmailAddress;
            myCommandSave.Parameters.AddWithValue("@firstname", SqlDbType.VarChar).Value = firstname;
            myCommandSave.Parameters.AddWithValue("@lastname", SqlDbType.VarChar).Value = lastname;
            myCommandSave.Parameters.AddWithValue("@UserId", SqlDbType.NVarChar).Value = userid;
            myCommandSave.Parameters.AddWithValue("@RoleName", SqlDbType.VarChar).Value = role;


            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.Insertsimulation",
                  "@sessionId = " + sessionId + " ,@ForioSession = " + ForioSession + ", @simulationId = " + simulationId + ",@createdUrl =" + createdUrl
                  + ",@OrderProductVariantGuid =" + orderGuid + ",@EmailAddress=" + EmailAddress + ",@firstname=" + firstname + ",@lastname=" + lastname + ",@RoleName=" + role
                  , null, strPurchaseModuleName);

            myCommandSave.Connection = thisConnection;
            thisConnection.Open();
            _records = myCommandSave.ExecuteNonQuery();

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.Insertsimulationv",
                       "Record inserted into [wm_ShoppingCartItem] Successfully. Id = " + _records, null, strPurchaseModuleName);
            thisConnection.Close();
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.Insertsimulation",
                        "Information saved in table (wm_simulation & wm_forio_logins), simulation Id: " + simulationId + ",Email Id: "
                        + EmailAddress + ", Role: " + role, null, strPurchaseModuleName);
            return _records;
            //return i;
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.Insertsimulation",
                        "Error while inserting Simulation details into [wm_Simulation] table. Error: "
                        + ex.Message, ex, strPurchaseModuleName);
            Console.WriteLine(ex.Message);
            thisConnection.Close();
            return _records;
        }

    }


    /// <summary>
    /// 
    /// </summary>
    /// <param name="CustomerId"></param>
    /// <param name="OrderId"></param>
    /// <param name="orderGuid"></param>
    /// <param name="Product"></param>
    /// <returns></returns>
    public int InsertVideoSubscription(Int32 CustomerId, Int32 OrderId, Int32 strProductId, String strInsertVideoURL, String strGuid)
    {

        int intRecords = 0;
        try
        {
            SqlCommand myCommandSave = new SqlCommand();
            myCommandSave.CommandType = CommandType.StoredProcedure;
            myCommandSave.CommandText = "pr_uva_InsertSubscriptionDetails";
            myCommandSave.Parameters.AddWithValue("@OrderGuid", SqlDbType.VarChar).Value = strGuid;
            myCommandSave.Parameters.AddWithValue("@CustomerId", SqlDbType.VarChar).Value = CustomerId;
            myCommandSave.Parameters.AddWithValue("@OrderId", SqlDbType.VarChar).Value = OrderId;
            myCommandSave.Parameters.AddWithValue("@SubscriptionURL", SqlDbType.VarChar).Value = strInsertVideoURL;
            myCommandSave.Parameters.AddWithValue("@ProductId", SqlDbType.VarChar).Value = strProductId;
            myCommandSave.Connection = thisConnection;
            thisConnection.Open();
            intRecords = (int)myCommandSave.ExecuteNonQuery();

            //change log just pass mention the return value
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                        "Information saved in table (wm_SubscriptionDetails), for OrderId: " + OrderId + ", CustomerID: "
                        + CustomerId + ", No of days Subscripbed:  ", null, strPurchaseModuleName);
            thisConnection.Close();
            return intRecords;
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.InsertDownload",
                         ex.Message, ex, strPurchaseModuleName);
            return 0;
        }
    }

    public int InsertDownload(string DownloadUrl, string FileName)
    {
        int downloadId = 0;
        string Extension = ".zip";
        Guid g;
        g = Guid.NewGuid();
        try
        {
            string fileContentType = GetContentType(FileName + Extension);

            thisConnection.Open();
            //'" + null + "'  - CAST(" + fileContentType + " AS nvarchar(MAX))'
            string insertDownload = "INSERT INTO Download(DownloadGuid,UseDownloadUrl,DownloadUrl,ContentType,[Filename],Extension, IsNew)OUTPUT INSERTED.ID " +
                "values('" + g + "','" + true + "','" + DownloadUrl + "','" + fileContentType + "','" + FileName + "','" + Extension + "','" + true + "')";
            SqlCommand cmdInsertDownload = new SqlCommand(insertDownload, thisConnection);
            //SqlCommand cmd = new SqlCommand(strquery, con);
            cmdInsertDownload.Parameters.AddWithValue("Guid", g);
            cmdInsertDownload.Parameters.AddWithValue("UseDownloadUrl", true);
            cmdInsertDownload.Parameters.AddWithValue("DownloadUrl", DownloadUrl);//Convert.ToString(myUrl));
            //cmdInsertDownload.Parameters.Add("@binaryValue", SqlDbType.VarBinary, 8000).Value = downloadBinary;
            //cmdInsertDownload.Parameters.Add("DownloadBinary", SqlDbType.VarBinary).Value = dwnloadBinary;
            cmdInsertDownload.Parameters.AddWithValue("ContentType", fileContentType);
            cmdInsertDownload.Parameters.AddWithValue("Filename", FileName);
            cmdInsertDownload.Parameters.AddWithValue("Extension", Extension);
            cmdInsertDownload.Parameters.AddWithValue("IsNew", true);

            Console.WriteLine("Add Download Detail.");

            //cmd.ExecuteNonQuery();
            downloadId = (Int32)cmdInsertDownload.ExecuteScalar();
            Console.WriteLine("Add Download detail completed.");

            //downloadId = Convert.ToInt32(cmdInsertDownload.ExecuteScalar());
            cmdInsertDownload.Dispose();
            //}
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.InsertDownload",
                        "Inserted url and other details to DOWNLOAD table successfully.", null, strPurchaseModuleName);
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.InsertDownload",
                        "Error while inserting url and other details to DOWNLOAD table. Error: " + ex.Message, ex, strPurchaseModuleName);
            Console.WriteLine(ex.Message);

        }
        finally
        {
            ////ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.InsertDownload", 
            //"In Finally of try catch closed the connection", null, _getConnection.strPurchaseModuleName);
            thisConnection.Close();
        }
        return downloadId;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="intOrderId"></param>
    /// <returns></returns>
    public int InsertCompletedOrders(int intOrderId)
    {
        int _records = 0;
        try
        {
            string strCompletedBy = strCompletedByFlagValue;
            SqlCommand myCommandSave = new SqlCommand();
            //SqlCommand myCommand = new SqlCommand("pr_UVA_GetProductIdByUrlRecord", myConnection);
            myCommandSave.CommandType = CommandType.StoredProcedure;
            myCommandSave.CommandText = "pr_uva_InsertCompletedOrders";
            myCommandSave.Parameters.AddWithValue("@OrderId", intOrderId);
            myCommandSave.Parameters.AddWithValue("@Date", DateTime.UtcNow);
            myCommandSave.Parameters.AddWithValue("@CompletedBy", strCompletedBy);

            myCommandSave.Connection = thisConnection;
            thisConnection.Open();
            _records = (int)myCommandSave.ExecuteNonQuery();
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.InsertCompletedOrders",
                       "Record inserted into [wm_CompletedOrders] Successfully. Id = " + _records, null, strPurchaseModuleName);
            thisConnection.Close();
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.InsertCompletedOrders",
                        "Error while inserting order details into [wm_CompletedOrders] table. Error: "
                        + ex.Message, ex, strPurchaseModuleName);
        }
        finally
        {
            thisConnection.Close();
        }

        return _records;
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="CustomerId"></param>
    /// <param name="SessionId"></param>
    /// <returns></returns>
    public int InsertSimulationToLibrary(int CustomerId, string SessionId)
    {
        int _records = 0;
        try
        {
            SqlCommand myCommandSave = new SqlCommand();
            //SqlCommand myCommand = new SqlCommand("pr_UVA_GetProductIdByUrlRecord", myConnection);
            myCommandSave.CommandType = CommandType.StoredProcedure;
            myCommandSave.CommandText = "pr_uva_InsertSimulationToLibrary";
            myCommandSave.Parameters.AddWithValue("@CustomerId", CustomerId);
            myCommandSave.Parameters.AddWithValue("@SessionId", SessionId);
            myCommandSave.Connection = thisConnection;
            thisConnection.Open();
            _records = (int)myCommandSave.ExecuteNonQuery();
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.InsertCompletedOrders",
                       "Record inserted into [wm_ShoppingCartItem] Successfully. Id = " + _records, null, strPurchaseModuleName);
            thisConnection.Close();
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.InsertCompletedOrders",
                        "Error while inserting order details into [wm_ShoppingCartItem] table. Error: "
                        + ex.Message, ex, strPurchaseModuleName);
        }
        finally
        {
            thisConnection.Close();
        }

        return _records;
    }

    #endregion

    #region Update

    private void UpdateOrderStatusWithOnlyExtension(int intOrderId)
    {
        try
        {
            SqlCommand thisCommand = thisConnection.CreateCommand();
            int value = Convert.ToInt32(ConfigurationManager.AppSettings["Pending"]);
            thisCommand.CommandText = "UPDATE [Order] SET OrderStatusId =" + Convert.ToInt32(ConfigurationManager.AppSettings["Complete"]) +
                ", PaymentStatusId =" + Convert.ToInt32(ConfigurationManager.AppSettings["Complete"]) +
                ", PaidDateUtc ='" + DateTime.UtcNow +
                "' WHERE Id =" + intOrderId;
            thisConnection.Open();
            thisCommand.ExecuteNonQuery();
            thisConnection.Close();
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.UpdateOrderStatusWithOnlyExtension",
                        "Updating orderStatusId for" + intOrderId + " to " + Convert.ToInt32(ConfigurationManager.AppSettings["Complete"]),
                        null, strPurchaseModuleName);
        }
        catch (Exception ex)

        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.UpdateOrderStatusWithOnlyExtension",
                        "Error while updating for" + intOrderId + " to " + Convert.ToInt32(ConfigurationManager.AppSettings["Complete"]) +
                        " Error: " + ex.Message, null, strPurchaseModuleName);
        }
    }

    private void UpdateOrderStatusWithHardCopy(int intOrderId)
    {
        try
        {
            SqlCommand thisCommand = thisConnection.CreateCommand();
            int value = Convert.ToInt32(ConfigurationManager.AppSettings["Pending"]);
            thisCommand.CommandText = "UPDATE [Order] SET OrderStatusId =" + Convert.ToInt32(ConfigurationManager.AppSettings["Processing"]) +
                ", PaymentStatusId =" + Convert.ToInt32(ConfigurationManager.AppSettings["Complete"]) +
                ", PaidDateUtc ='" + DateTime.UtcNow +
                "' WHERE Id =" + intOrderId;
            thisConnection.Open();
            thisCommand.ExecuteNonQuery();
            thisConnection.Close();
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.UpdateOrderStatusWithHardCopy",
                        "Updating orderStatusId for" + intOrderId + " to " + Convert.ToInt32(ConfigurationManager.AppSettings["Complete"]),
                        null, strPurchaseModuleName);
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.UpdateOrderStatusWithHardCopy",
                        "Error while updating for" + intOrderId + " to " + Convert.ToInt32(ConfigurationManager.AppSettings["Complete"]) +
                        " Error: " + ex.Message, null, strPurchaseModuleName);
        }
    }

    public int UpdateOrderReceiptSend(int OrderId)
    {
        int _records = 0;
        try
        {
            SqlCommand myCommandSave = new SqlCommand();
            myCommandSave.CommandType = CommandType.StoredProcedure;
            myCommandSave.CommandText = "pr_uva_UpdateOrderReceiptSend";
            myCommandSave.Parameters.AddWithValue("@OrderId", OrderId);
            myCommandSave.Connection = thisConnection;
            thisConnection.Open();
            _records = (int)myCommandSave.ExecuteNonQuery();
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.UpdateOrderReceiptSend",
                       "Record inserted into [wm_CompletedOrders] Successfully. Id = " + _records, null, strPurchaseModuleName);
            thisConnection.Close();
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.UpdateOrderReceiptSend",
                        "Error while inserting order details into [wm_CompletedOrders] table. Error: "
                        + ex.Message, ex, strPurchaseModuleName);
        }
        finally
        {
            thisConnection.Close();
        }

        return _records;
    }


    #endregion

    #region Collect Data(Get)

    private DataTable GetOrderDetailsCP(int OrderId, int intCPId)
    {
        DataTable dt = new DataTable();
        try
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.GetOrderDetailsCP", "Stored Procedure: pr_uva_GetOrderDetails_CP ,Parameter: OrderId ="
                + OrderId + ",CPId=" + intCPId, null, strPurchaseModuleName);
            SqlCommand cmd = new SqlCommand("pr_uva_GetOrderDetails_CP", thisConnection);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@OrderId", OrderId);
            cmd.Parameters.AddWithValue("@CPId", intCPId);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt);
            return dt;
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.GetOrderDetailsCP", ex.Message, ex, strPurchaseModuleName);
            return dt;
        }

    }

    private DataTable GetMicrositeProductDetails(int micrositeProductId)
    {
        DataTable dt = new DataTable();
        try
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.pr_uva_GetMicrositeProductDetails", "Stored Procedure: pr_uva_GetMicrositeProductDetails ,Parameter: OrderId ="
                , null, strPurchaseModuleName);
            SqlCommand cmd = new SqlCommand("pr_uva_GetMicrositeProductDetails", thisConnection);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@MicrositeProductId", micrositeProductId);
            SqlDataAdapter da = new SqlDataAdapter(cmd);
            da.Fill(dt);
            return dt;
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.GetMicrositeProductDetails", ex.Message, ex, strPurchaseModuleName);
            return dt;
        }

    }

    private DataTable GetInCompleteOrders()
    {
        DataTable dt = new DataTable();
        try
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.GetInCompleteOrders",
                      "Stored Procedure:pr_uva_GetOrderData", null, strPurchaseModuleName);
            SqlCommand cmdd = new SqlCommand("pr_uva_GetOrderData", thisConnection);
            cmdd.CommandType = CommandType.StoredProcedure;

            SqlDataAdapter da = new SqlDataAdapter(cmdd);
            da.Fill(dt);

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.GetInCompleteOrders",
                        "Get order data successfully", null, strPurchaseModuleName);
            return dt;
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.GetInCompleteOrders",
                        "Error while fetching records using sp 'pr_UVA_GetOrderData' Error:" + ex.Message,
                        null, strPurchaseModuleName);
            return dt;
        }
    }

    /// <summary>
    /// Get Customer Id And Payment Method from order table
    /// </summary>
    /// <param name="OrderId"></param>
    /// <returns></returns>
    private DataTable GetCustomerIdAndPaymentMethod(int OrderId)
    {
        DataTable dt = new DataTable();
        try
        {
            SqlCommand cmdd = new SqlCommand("pr_uva_GetCustomerIdAndPaymentMethodFromOrder", thisConnection);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                        "GetConnection.GetCustomerIdAndPaymentMethod",
                        "Stored Procedure : pr_uva_GetCustomerIdAndPaymentMethodFromOrder , Parameter : OrderId ="
                        + OrderId, null, strPurchaseModuleName);

            cmdd.CommandType = CommandType.StoredProcedure;
            cmdd.Parameters.AddWithValue("@OrderId", OrderId);
            SqlDataAdapter da = new SqlDataAdapter(cmdd);
            da.Fill(dt);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success,
                        "GetConnection.GetCustomerIdAndPaymentMethod",
                        "Get Order Customer Id And Payment Method successfully",
                         null, strPurchaseModuleName);
            return dt;
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error,
                        "GetConnection.GetCustomerIdAndPaymentMethod",
                        "Stored Procedure : pr_uva_GetCustomerIdAndPaymentMethodFromOrder: " + ex.Message,
                        ex, strPurchaseModuleName);
            return dt;
        }

    }



    private DataTable GetOrderDetails(int OrderId)
    {
        DataTable dt = new DataTable();
        try
        {
            SqlCommand cmdd = new SqlCommand("pr_uva_GetOrderDetails", thisConnection);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.GetOrderDetails",
                     "Executing Stored Procedure : pr_uva_GetOrderDetails , Parameter - OrderId : " + OrderId, null, strPurchaseModuleName);

            cmdd.CommandType = CommandType.StoredProcedure;
            cmdd.Parameters.AddWithValue("@OrderId", OrderId);
            SqlDataAdapter da = new SqlDataAdapter(cmdd);
            da.Fill(dt);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.GetOrderDetails",
                      "Get Order Details successfully",
                      null, strPurchaseModuleName);
            return dt;
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.GetOrderDetails",
                        "Stored Procedure : pr_uva_GetOrderDetails: " + ex.Message, ex, strPurchaseModuleName);
            return dt;
        }

    }

    private DataTable GetOrderProductData(int OrderId)
    {
        DataTable dt = new DataTable();
        try
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.GetOrderProductData",
             "Executing Stored Procedure : pr_uva_GetOrderProductData,Parameter - OrderId : " + OrderId, null, strPurchaseModuleName);
            SqlCommand cmdd = new SqlCommand("pr_uva_GetOrderProductData", thisConnection);
            cmdd.CommandType = CommandType.StoredProcedure;
            cmdd.Parameters.AddWithValue("@OrderId", OrderId);
            SqlDataAdapter da = new SqlDataAdapter(cmdd);
            da.Fill(dt);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.GetOrderProductData",
                   "Successfully fetched all records from ORDER table.", null, strPurchaseModuleName);
            return dt;
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.GetOrderProductData",
                        ex.Message, ex, strPurchaseModuleName);
            return dt;
        }


    }

    private DataTable GetCustomerAddress(int AddressId)
    {
        DataTable dt = new DataTable();
        try
        {
            SqlCommand cmdd = new SqlCommand("pr_uva_GetCustomerAddress", thisConnection);
            cmdd.CommandType = CommandType.StoredProcedure;
            cmdd.Parameters.AddWithValue("@AddressId", AddressId);
            SqlDataAdapter da = new SqlDataAdapter(cmdd);
            da.Fill(dt);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.GetCustomerAddress",
                  "Successfully fetched all records from ADDRESS table.", null, strPurchaseModuleName);
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.GetCustomerAddress",
                        ex.Message, ex, strPurchaseModuleName);
        }

        return dt;
    }

    private DataTable GetProductDetails(int OrderId)
    {
        DataTable dt = new DataTable();
        try
        {
            SqlCommand cmdd = new SqlCommand("pr_uva_GetProductDetails", thisConnection);
            cmdd.CommandType = CommandType.StoredProcedure;
            cmdd.Parameters.AddWithValue("@OrderId", OrderId);
            SqlDataAdapter da = new SqlDataAdapter(cmdd);
            da.Fill(dt);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                  "GetConnection.GetProductDetails", "Successfully fetched all records from PRODUCT table.",
                  null, strPurchaseModuleName);
        }
        catch (Exception ex)
        {

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.GetProductDetails",
                        ex.Message, ex, strPurchaseModuleName);
        }

        return dt;
    }

    public DataTable GetOrderPlacedDetailforOrderId(int intOrderId)
    {
        DataTable dt = new DataTable();
        try
        {
            SqlCommand cmdd = new SqlCommand("pr_uva_GetOrderPlacedDetailForOrderId", thisConnection);
            cmdd.CommandType = CommandType.StoredProcedure;
            cmdd.Parameters.AddWithValue("@OrderId", intOrderId);
            SqlDataAdapter da = new SqlDataAdapter(cmdd);
            da.Fill(dt);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.GetOrderPlacedDetailforOrderId",
                      "Successfully get OrderId : " + intOrderId + " Detail for sending Order Receipt mail",
                      null, strPurchaseModuleName);
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.GetOrderPlacedDetailforOrderId",
                        "Error while fetching records using sp 'pr_uva_GetOrderPlacedDetailForOrderId'. Error : " + ex.Message,
                        null, strPurchaseModuleName);
        }
        return dt;
    }

    public DataTable GetOrderPlacedDetail()
    {
        DataTable dt = new DataTable();
        try
        {
            SqlCommand cmdd = new SqlCommand("pr_uva_GetOrderPlacedDetail", thisConnection);
            cmdd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter da = new SqlDataAdapter(cmdd);
            da.Fill(dt);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.GetOrderPlacedDetail",
                      "Get Order placed info successfully",
                      null, strPurchaseModuleName);
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.GetOrderPlacedDetail",
                        "Error while fetching records using sp 'pr_UVA_GetOrderDetails' Error:" + ex.Message,
                        null, strPurchaseModuleName);
        }
        return dt;
    }


    private DataTable GetCustomerEmail(int CustomerId)
    {
        DataTable dt = new DataTable();
        try
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.GetCustomerInfo",
                "Stored Proccedure : pr_uva_GetCustomerEmail ,Parameter:CustomerId=" + CustomerId, null, strPurchaseModuleName);

            SqlCommand cmdd = new SqlCommand("pr_uva_GetCustomerEmail", thisConnection);
            cmdd.CommandType = CommandType.StoredProcedure;
            cmdd.Parameters.AddWithValue("@CustomerId", CustomerId);
            SqlDataAdapter da = new SqlDataAdapter(cmdd);
            da.Fill(dt);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.GetCustomerInfo",
                   "Successfully fetched all records from CUSTOMER table.", null, strPurchaseModuleName);

            return dt;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.GetCustomerInfo",
                        ex.Message, ex, strPurchaseModuleName);
            return dt;
        }


    }

    private DataTable GetUserType(int CustomerId)
    {
        DataTable dt = new DataTable();
        try
        {
            SqlCommand cmdd = new SqlCommand("pr_uva_GetUserType", thisConnection);
            cmdd.CommandType = CommandType.StoredProcedure;
            cmdd.Parameters.AddWithValue("@CustomerId", CustomerId);
            SqlDataAdapter da = new SqlDataAdapter(cmdd);
            da.Fill(dt);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.GetUserType",
                       "Successfully fetched all records from CUSTOMER table like university name,user type.", null, strPurchaseModuleName);
            return dt;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.GetUserType",
                        "Error while fetching university name. " + ex.Message, ex, strPurchaseModuleName);
            return dt;
        }


    }

    public string GetContentType(string fileName)
    {
        string contentType = "application/octetstream";
        string ext = System.IO.Path.GetExtension(fileName).ToLower();

        Microsoft.Win32.RegistryKey registryKey = Microsoft.Win32.Registry.ClassesRoot.OpenSubKey(ext);
        if (registryKey != null && registryKey.GetValue("Content Type") != null)
            contentType = registryKey.GetValue("Content Type").ToString();

        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.GetContentType",
                    "Getting purchased product's type", null, strPurchaseModuleName);

        return contentType;
    }

    public static string GetconnectionString
    {
        get
        {
            return ConfigurationManager.ConnectionStrings["ConnectionString"].ConnectionString.ToString();
        }
    }

    public int GetProductAttributeId(int intProductVariantAttributeId)
    {
        int intReturnValueId = 0;
        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.GetProductAttributeId",
                      "Parameter :intProductVariantAttributeId =" + intProductVariantAttributeId, null, strPurchaseModuleName);
        try
        {

            SqlCommand cmdd = new SqlCommand("pr_uva_GetProductAttributeId", thisConnection);
            cmdd.CommandType = CommandType.StoredProcedure;
            cmdd.Parameters.AddWithValue("@ProductVariantAttributeId", intProductVariantAttributeId);
            thisConnection.Open();
            intReturnValueId = Convert.ToInt32(cmdd.ExecuteScalar());

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.GetProductAttributeId",
                       "get intReturnValueId =" + intReturnValueId, null, strPurchaseModuleName);
            thisConnection.Close();
            return intReturnValueId;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.GetProductAttributeId",
                          ex.Message, ex, strPurchaseModuleName);
            thisConnection.Close();
            return intReturnValueId;
        }
    }

    //To get simulation platform type (i.e. Simulate or Epicenter) for respective product.
    public string GetSimulationPlatformType(int ProductId)
    {
        string returnSimuationType = String.Empty;
        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.GetSimulationPlatformType",
                      "Checking whether the simulation platform is 'Simulate' or 'Epicenter' using input parameter :ProductId =" + ProductId, null, strPurchaseModuleName);
        try
        {
            SqlCommand cmd = new SqlCommand("pr_uva_GetSimulationType", thisConnection);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@productId", ProductId);
            thisConnection.Open();
            returnSimuationType = Convert.ToString(cmd.ExecuteScalar());

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.GetSimulationPlatformType",
                       "Simuation platform type =" + returnSimuationType, null, strPurchaseModuleName);
            thisConnection.Close();
            return returnSimuationType;
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.GetSimulationPlatformType",
                          ex.Message, ex, strPurchaseModuleName);
            return returnSimuationType;
        }

    }

    public bool AddMemberAPI(int orderDetailsCount, int orderId, string strOAuthToken, int customerId, string strUserId, string SessionId,
                                string strUserRole, string simulationUrl, string UserType, string strGroupId)
    {
        bool ReturnStatus = true;

        DataTable dtOrderDetails = GetOrderDetails(orderId);

        EpicenterAPI api = new EpicenterAPI(strEpicenterPublicKey, strEpicenterSecretKey);

        #region add member to group API

        try
        {
            dynamic addmemberresponse = api.AddMemberToGroup(strOAuthToken, strUserId, strGroupId, strUserRole);

            try
            {
                strResponseId = addmemberresponse["id"].ToString();

                if (!string.IsNullOrEmpty(strResponseId))
                {
                    strStatusCode = "201";
                }
                else
                {
                    strStatusCode = "";
                }

            }
            catch (Exception ex)
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.CallEpicenterAPI",
                "The key name as id is not found in API response hense API call is not successful ", ex, strPurchaseModuleName);
                strStatusCode = "";
            }

            if (strStatusCode == "201")
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.CallEpicenterAPI",
                    "User : " + strUserId + " successfully added to group: " + strGroupId + "for OrderId:" + orderId,
                     null, strPurchaseModuleName);

                Insertsimulation(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString(),
                                     strGroupId, dtOrderDetails.Rows[orderDetailsCount]["ParentGroupedProductId"].ToString(),
                                     simulationUrl, dtOrderDetails.Rows[orderDetailsCount]["OrderItemGuid"].ToString(),
                                     dtOrderDetails.Rows[orderDetailsCount]["email"].ToString(),
                                     dtOrderDetails.Rows[orderDetailsCount]["FirstName"].ToString(),
                                     dtOrderDetails.Rows[orderDetailsCount]["LastName"].ToString(),
                                     strUserId,
                                     UserType);

                sessions.Add(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString());
                if (strIsSimulationSession == "true")
                {
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.CallEpicenterAPI",
                                "Purchased only simulation session, System sending email for simulation link to faculty for OrderId:" + orderId,
                                null, strPurchaseModuleName);

                    //Send storefront link to factuly for automatic purchase of License.
                    simulationUrl = ConfigurationManager.AppSettings["StorefrontSite"] + "GroupId=" +
                                    strGroupId;


                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.CallEpicenterAPI",
                                "SImulation URL :" + simulationUrl + "for OrderId:" + orderId,
                                null, strPurchaseModuleName);

                    SendSimulationMailToFaculty(strGroupId,
                                        dtOrderDetails.Rows[orderDetailsCount]["FirstName"].ToString() + " " +
                                        dtOrderDetails.Rows[orderDetailsCount]["LastName"].ToString(),
                                        dtOrderDetails.Rows[orderDetailsCount]["email"].ToString(),
                                        simulationUrl, dtOrderDetails.Rows[orderDetailsCount]["Name"].ToString(),
                                        dtOrderDetails.Rows[orderDetailsCount]["Id"].ToString(),
                                        dtOrderDetails.Rows[orderDetailsCount]["AttributeDescription"].ToString());
                }
                forioCount++;
                this.InsertSimulationToLibrary((Convert.ToInt32(customerId)), SessionId);

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.CallEpicenterAPI",
                           "Simulation product added to library", null, strPurchaseModuleName);

            }
            else
            {
                try
                {
                    strMessage = addmemberresponse["message"];
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning, "GetConnection.CallEpicenterAPI",
                                        "(In Addmember API) error: " + strMessage, null, strPurchaseModuleName);
                }
                catch (Exception ex)
                {
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.CallEpicenterAPI",
                     ex.Message, ex, strPurchaseModuleName);

                }

                fileNotFound++;
                ReturnStatus = false;
            }
        }
        catch (Exception ex)
        {
            fileNotFound++;
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.CallEpicenterAPI",
            ex.Message, ex, strPurchaseModuleName);
            ReturnStatus = false;
        }
        #endregion

        return ReturnStatus;
    }

    public bool AddUserAPI(int orderDetailsCount, int orderId, string strOAuthToken, string strFirstName, string strLastName,
                               string strEmailId, int customerId, string strForioName, string strEpicenterAccount, string strUniversity,
                               string simulationUrl, string SessionId, string strProjectFullName, string strGroupId, string strUserRole, string UserType)
    {
        bool ReturnStatus = true;

        DataTable dtOrderDetails = GetOrderDetails(orderId);

        EpicenterAPI api = new EpicenterAPI(strEpicenterPublicKey, strEpicenterSecretKey);

        #region for user creation

        try
        {
            strPassword = strFirstName + "@" + strLastName + orderId;
            dynamic userresponse = api.CreateUser(strOAuthToken, strEmailId, strEpicenterAccount, strPassword, strFirstName, strLastName);

            try
            {
                strResponseUserId = userresponse["id"].ToString();

                if (!string.IsNullOrEmpty(strResponseUserId))
                {
                    strStatusCode = "201";
                }
                else
                {
                    strStatusCode = "";
                }

            }
            catch (Exception ex)
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.CallEpicenterAPI",
                "The key name as (User) Id  is not found in API response hense API call is not successful ", ex, strPurchaseModuleName);
                strStatusCode = "";
            }

            if (strStatusCode == "201")
            {
                var strUserId = userresponse["id"];
                SendEpicenterAuthenticationMailToUser(strEmailId, strLoginEpicenterUrl, strForioName,
                                                      strEmailId, strPassword);


                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.CallEpicenterAPI",
                            "New user created in forio.com, user name: "
                             + strEmailId + ", User id: " + strUserId
                             + ", account: darden for OrderId:" + orderId,
                             null, strPurchaseModuleName);

                #region for password recovery API

                try
                {
                    simulationUrl = "http://forio.com/app/darden/marketing";

                    // simulationUrl = Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["SimulationPath"]);
                    dynamic passrecoveryresponse = api.PasswordRecoveryAPI(strEmailId, strEpicenterAccount, strProjectFullName, simulationUrl);

                    try
                    {
                        strResponseMessagepassrecovery = passrecoveryresponse["message"].ToString();

                        if (!string.IsNullOrEmpty(strResponseMessagepassrecovery))
                        {
                            strStatusCode = "200";
                        }
                        else
                        {
                            strStatusCode = "";
                        }

                    }
                    catch (Exception ex)
                    {
                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.CallEpicenterAPI",
                        "The key name as Message is not found in API response hense API call is not successful ", ex, strPurchaseModuleName);
                        strStatusCode = "";
                    }

                    if (strStatusCode == "200")
                    {
                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.CallEpicenterAPI",
                                "Notification sent to user name: "
                                 + strEmailId + " for account: " + strEpicenterAccount
                                 + ", redirect to link " + simulationUrl + "for OrderId:" + orderId,
                                 null, strPurchaseModuleName);

                        AddMemberAPI(orderDetailsCount, orderId, strOAuthToken, customerId, strUserId, SessionId, strUserRole, simulationUrl, UserType, strGroupId);

                    }
                    else
                    {
                        try
                        {
                            strMessage = passrecoveryresponse["message"];

                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning, "GetConnection.CallEpicenterAPI",
                                                    "error: " + strMessage, null, strPurchaseModuleName);
                        }
                        catch (Exception ex)
                        {

                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.CallEpicenterAPI",
                            ex.Message, ex, strPurchaseModuleName);
                        }

                        fileNotFound++;
                        ReturnStatus = false;
                    }
                }
                catch (Exception ex)
                {
                    fileNotFound++;
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.CallEpicenterAPI",
                    ex.Message, ex, strPurchaseModuleName);
                    ReturnStatus = false;
                }
                #endregion
            }
            else
            {
                strMessage = userresponse["message"];

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning, "GetConnection.CallEpicenterAPI",
                                            "error: " + strMessage, null, strPurchaseModuleName);

                dynamic responseEnduser = api.SearchEndUser(strOAuthToken, strEpicenterAccount, strEmailId);

                try
                {
                    strSearchResponseUserId = (responseEnduser[0])["id"];

                    if (!string.IsNullOrEmpty(strSearchResponseUserId))
                    {
                        strStatusCode = "200";
                    }
                    else
                    {
                        strStatusCode = "";
                    }

                }
                catch (Exception ex)
                {
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.CallEpicenterAPI",
                    "The key name as Id is not found in API response hense API call is not successful ", ex, strPurchaseModuleName);
                    strStatusCode = "";
                }

                if (strStatusCode == "200")
                {
                    var strSearchedUserId = (responseEnduser[0])["id"];

                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.CallEpicenterAPI",
                                        "Successfully found user id : " + strSearchedUserId + " of user name:  " + strEmailId + "for OrderId:" + orderId
                                        , null, strPurchaseModuleName);

                    AddMemberAPI(orderDetailsCount, orderId, strOAuthToken, customerId, strSearchedUserId, SessionId, strUserRole, simulationUrl, UserType, strGroupId);
                }
                else
                {
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning, "GetConnection.CallEpicenterAPI",
                                            "An user : " + strEmailId + " is not found in search user API", null, strPurchaseModuleName);
                    fileNotFound++;
                    ReturnStatus = false;
                }
            }

        }
        catch (Exception ex)
        {
            fileNotFound++;
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.CallEpicenterAPI",
                ex.Message, ex, strPurchaseModuleName);
            ReturnStatus = false;
        }
        #endregion

        return ReturnStatus;
    }





    public bool CallEpicenterAPI(int orderDetailsCount, int orderId, string strOAuthToken, string strFirstName, string strLastName,
                                 string strEmailId, int customerId, string strForioName, string strEpicenterAccount, string strUniversity,
                                 string simulationUrl, string SessionId, string UserType)
    {
        bool ReturnStatus = true;

        DataTable dtOrderDetails = GetOrderDetails(orderId);
        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.CallEpicenterAPI",
                                                                "Purchase Epicenter simulation session/simulation session + User license for OrderId:" + orderId, null, strPurchaseModuleName);

        string strClassDescription = dtOrderDetails.Rows[orderDetailsCount]["AttributesXml"].ToString();

        System.Xml.XmlTextReader reader = new System.Xml.XmlTextReader(new System.IO.StringReader(strClassDescription));
        reader.Read();
        System.Data.DataSet ds = new System.Data.DataSet();
        ds.ReadXml(reader, System.Data.XmlReadMode.Auto);
        description = ds.Tables["ProductVariantAttributeValue"].Rows[0]["Value"].ToString();

        //strGroupName = "DBP-" + orderId + "-" + strUniversity;
        Guid uuid = Guid.NewGuid();
        string uuidString = uuid.ToString();
     
        strGroupName = "DBP-" + orderId + "-" + uuidString;
        strGroupName = strGroupName.ToLower();
        strForioName = dtOrderDetails.Rows[orderDetailsCount]["Name"].ToString();
        strProjectFormatId = strForioName.ToLower() + "-" + description;

        // strUserRole = "facilitator";
        //ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
        //            "get the class description from attribute xml for epicenter platform", null, strPurchaseModuleName);

        DataRow[] dr = dtOrderDetails.Select("sessionid='" +
                       dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString() + "'");

        EpicenterAPI api = new EpicenterAPI(strEpicenterPublicKey, strEpicenterSecretKey);
        #region for project id creation
        try
        {

            //ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.CallEpicenterAPI",
            //                "Project created in forio.com, ProjectId:" + strProjectId, null, strPurchaseModuleName);

            #region for Group creation API

            if (!sessions.Contains(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString()))
            {
                if (dr.Length == 1)
                {
                    maxUserCount = "0";
                    strIsSimulationSession = "true";

                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                                "GetConnection.getConnection", "Purchase simulation session Only,"
                                + " maxUsers=" + maxUserCount + "for OrderId:" + orderId, null, strPurchaseModuleName);

                    postData = "name=" + strGroupName + "account=" + strEpicenterAccount + "project=" + strProjectId;
                }
                else
                {
                    DataRow[] dr1 = dtOrderDetails.Select("sessionid='" + dtOrderDetails.Rows[orderDetailsCount]
                                    ["sessionid"].ToString() + "' and ProductVariant='" +
                                    ConfigurationManager.AppSettings["UserLicense"] + "'");
                    maxUserCount = dr1[0]["Quantity"].ToString();
                    strIsSimulationSession = "false";

                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.CallEpicenterAPI",
                                "Purchase user license:" + " maxusercount=" + maxUserCount + "for OrderId:" + orderId, null,
                                strPurchaseModuleName);

                }

                try
                {
                    dynamic groupresponse = api.CreateGroup(strOAuthToken, strGroupName, strEpicenterAccount, strProjectId, maxUserCount);

                    try
                    {
                        strGroupId = groupresponse["groupId"].ToString();

                        if (!string.IsNullOrEmpty(strGroupId))
                        {
                            strStatusCode = "201";
                        }
                        else
                        {
                            strStatusCode = "";
                        }

                    }
                    catch (Exception ex)
                    {
                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.CallEpicenterAPI",
                        "The key name as Group Id  is not found in API response hense API call is not successful ", ex, strPurchaseModuleName);
                        strStatusCode = "";
                    }

                    if (strStatusCode == "201")
                    {
                        strGroupId = groupresponse["groupId"];
                        strGroupname = groupresponse["name"];
                        strProjectName = groupresponse["project"];

                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.CallEpicenterAPI",
                                        "Group created in forio.com, group name: "
                                         + strGroupname + ", group id: " + strGroupId
                                         + ", account: darden, project name: " + strProjectName + "for OrderId:" + orderId,
                                         null, strPurchaseModuleName);

                        AddUserAPI(orderDetailsCount, orderId, strOAuthToken, strFirstName, strLastName, strEmailId, customerId, strForioName,
                                   strEpicenterAccount, strUniversity, simulationUrl, SessionId, strProjectFullName, strGroupId, strFacultyUserRole, UserType);
                    }
                    else
                    {
                        strMessage = groupresponse["message"];

                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning, "GetConnection.CallEpicenterAPI",
                                    "In group API error: " + strMessage, null, strPurchaseModuleName);
                        fileNotFound++;
                        ReturnStatus = false;
                    }
                }
                catch (Exception ex)
                {
                    fileNotFound++;
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.CallEpicenterAPI",
                                ex.Message, ex, strPurchaseModuleName);
                    ReturnStatus = false;
                }
            }
            #endregion

        }
        catch (Exception ex)
        {
            fileNotFound++;
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.CallEpicenterAPI",
                        ex.Message, ex, strPurchaseModuleName);


            ReturnStatus = false;
        }
        #endregion


        return ReturnStatus;
    }


    public bool CheckOAuthAPICallOrNot(int orderDetailsCount, int orderId, string strOAuthToken, string strFirstName, string strLastName,
                                  string strEmailId, int customerId, string strForioName, string strEpicenterAccount, string strUniversity,
                                  string simulationUrl, string SessionId, string UserType)
    {
        bool ReturnStatus = false;
        EpicenterAPI api = new EpicenterAPI(strEpicenterPublicKey, strEpicenterSecretKey);

        dynamic tokenresponse = api.createoauthtoken();
        try
        {
            strRefreshToken = tokenresponse["refresh_token"].ToString();

            if (!string.IsNullOrEmpty(strRefreshToken))
            {
                strStatusCode = "200";
            }
            else
            {
                strStatusCode = "";
            }

        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.CallEpicenterAPI",
                        "The key name as refresh_token  is not found in API response hense API call is not successful ", ex, strPurchaseModuleName);
            strStatusCode = "";
        }

        if (strStatusCode == "200")
        {
            strOAuthToken = tokenresponse["access_token"];

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.getConnection",
                    "Generate epicenter OAuth token :" + strOAuthToken + "for OrderId:" + orderId, null, strPurchaseModuleName);

            bool Status = CallEpicenterAPI(orderDetailsCount, orderId, strOAuthToken, strFirstName, strLastName, strEmailId, customerId,
                                            strForioName, strEpicenterAccount, strUniversity, simulationUrl, SessionId, strFacultyUser);
            if (Status == true)
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.getConnection",
                    "Successfully call all API for OrderId:" + orderId, null, strPurchaseModuleName);
                ReturnStatus = true;
            }
            else
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.getConnection",
                    "Failed to call API for OrderId:" + orderId, null, strPurchaseModuleName);
                ReturnStatus = false;

            }
        }
        else
        {
            CountOauthToken++;
            if (CountOauthToken < 3)
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning,
                                    "GetConnection.getConnection", " OAuth API is not called at iteration : " + CountOauthToken + "for OrderId:" + orderId,
                                    null, strPurchaseModuleName);
                //ApiStatus = CallOAuthAPI();
                //CheckAPICallOrNot(ApiStatus);
                // api.createoauthtoken();
                ReturnStatus = CheckOAuthAPICallOrNot(orderDetailsCount, orderId, strOAuthToken, strFirstName, strLastName, strEmailId, customerId,
                                            strForioName, strEpicenterAccount, strUniversity, simulationUrl, SessionId, strFacultyUser);
            }
            else
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning,
                                    "GetConnection.getConnection", " Process End because OAuth API is not called at iteration : " + CountOauthToken + "for OrderId:" + orderId,
                                    null, strPurchaseModuleName);
                ReturnStatus = false;
            }

        }

        return ReturnStatus;
    }







    #region Main

    public string singleOrderCompletion(int intOrderId, string strCompletedByFlag)
    {
        //new variable for rtoc webservice completed by flag
        strCompletedByFlagValue = strCompletedByFlag;
        string strfilenotfound = string.Empty;

        ArrayList sessions = new ArrayList();
        DataTable dtUserType = new DataTable();
        DataTable dtTarget = new DataTable();
        int hardCopyOrCopyrightPermissions = 0;
        int forioCount = 0;
        int intUserType = 0;
        DirectoryInfo myDirCopyFiles = null;
        DirectoryInfo myDirSupport = null;
        DirectoryInfo myDirFaculty = null;
        DirectoryInfo myDirSubFolder = null;
        string strFileExtension = string.Empty;
        string strCopyZIPFile = string.Empty;
        string strZipPath = string.Empty;
        string userName = string.Empty;
        string totalQuantity = string.Empty;
        string parentDirectoryPath = string.Empty;
        string childDirectoryPath = string.Empty;
        string uName = string.Empty;
        string d = string.Empty;
        //Product Order Number for zipFolder                   
        string strzipName = string.Empty;
        string strSlug = string.Empty;
        string strProduct = string.Empty;
        string strUniversity = string.Empty;
        string SessionId = string.Empty;
        string description = string.Empty;
        string organization = string.Empty;
        string allowUpload = "false";
        string maxUserCount = "0";
        string singleSignOnOnly = "false";
        string simulationUrl = string.Empty;
        string strProductType = string.Empty;
        //Microsite variable declaration
        string strMicrositeCopyZIPFile = string.Empty;
        DirectoryInfo myMIcroDirCopyFiles = null;
        //new variables 
        DataTable dtOrder = new DataTable();

        try
        {
            //Get Customer Id And Payment Method from Order Id
            dtOrder = GetCustomerIdAndPaymentMethod(intOrderId);

            Hashtable OrderHashTable = new Hashtable();
            #region If Order Count >0
            //If order count greater than 0 the proceed.
            if (dtOrder.Rows.Count > 0)
            {
                int extentCount = 0;
                int SubscriptionCount = 0;
                int fileNotFound = 0;

                OrderHashTable.Add("extentCount", extentCount);
                OrderHashTable.Add("SubscriptionCount", SubscriptionCount);
                OrderHashTable.Add("fileNotFound", fileNotFound);
                int customerId = Convert.ToInt32(dtOrder.Rows[0]["CustomerId"]);

                // Get User Type
                dtUserType = GetUserType(Convert.ToInt32(customerId));
                if (dtUserType.Rows.Count > 0)
                {
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                                "GetConnection.singleOrderCompletion",
                                "Processing OrderId: " + intOrderId, null, strPurchaseModuleName);

                    // Get Order Details
                    DataTable dtOrderDetails = GetOrderDetails(intOrderId);

                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                               "Order Items Count is : " + dtOrderDetails.Rows.Count + "", null, strPurchaseModuleName);

                    //If Order details present, then proceed.
                    if (dtOrderDetails.Rows.Count > 0)
                    {
                        try
                        {
                            #region for each Order details
                            int intOrderItemCount = 0;

                            Console.WriteLine("Processing order");
                            Console.WriteLine("Order Id: " + intOrderId);

                            //Processing individual order
                            for (int orderDetailsCount = 0; orderDetailsCount < dtOrderDetails.Rows.Count; orderDetailsCount++)
                            {
                                Console.WriteLine("Product Name:" + dtOrderDetails.Rows[orderDetailsCount]["Product"]);
                                Console.WriteLine("Product Type:" + dtOrderDetails.Rows[orderDetailsCount]["ProductType"]);

                                strProductType = Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["ProductType"]);
                                #region LocalVariables
                                //strProductNumber =M-0871
                                string strProductNumber = Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["ProductNumber"]);

                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                              "Get Next Order Item", null, strPurchaseModuleName);

                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                              "Product : " + dtOrderDetails.Rows[orderDetailsCount]["Product"] + " , ProductType : " + strProductType + " ProductNumber : " + strProductNumber + ", Product Variant : " + dtOrderDetails.Rows[orderDetailsCount]["ProductVariant"] + " , FileExtension : " + dtOrderDetails.Rows[orderDetailsCount]["FileExtension"] + "", null, strPurchaseModuleName);

                                //userName =Vijay Nagarsoge
                                userName = System.Text.RegularExpressions.Regex.Replace(
                                                        Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["FirstName"] + " " + dtOrderDetails.Rows[orderDetailsCount]["LastName"]),
                                                        @"[*?:\|<>\\/]", "").Replace("\"", "");


                                DateTime createdDate = Convert.ToDateTime(dtOrderDetails.Rows[orderDetailsCount]["CreatedOnUtc"]);
                                totalQuantity = Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["Quantity"]);//1

                                //modifying zip file name with format: uname_purchasedate_ordernumber
                                uName = userName.Replace(" ", "");//vijaynagarsoge
                                d = createdDate.ToString("yyMMdd");//190916
                                //Product Order Number for zipFolder
                                //Regex pattern = new Regex("[;,\t\r ]|[\n]{2}");                                   
                                strzipName = uName + '_' + d + '_' + intOrderId;//strzipName =VijayNagarsoge_190916_241510
                                strZipPath = Path.Combine(compressedFilePath + strzipName);//C:\inetpub\wwwroot\WMBranding\Purchased Products\VijayNagarsoge_190916_241505

                                #endregion
                                #region Delete Purchased microsite inside order placed folder for ProductOrderNumber with uname_purchasedate_ordernumber
                                //Used to delete folder if exist before order placed in Purchased Microsite folder.
                                strMicrositeCopyZIPFile = Convert.ToString(ConfigurationManager.AppSettings["MicrositeFilePath"]) + strzipName;
                                //strMicrositeCopyZIPFile= C:\inetpub\wwwroot\WMBranding\Purchased Microsites\VijayNagarsoge_190916_241505
                                if (Directory.Exists(strMicrositeCopyZIPFile))
                                {
                                    myMIcroDirCopyFiles = new DirectoryInfo(strMicrositeCopyZIPFile);
                                    DirectoryInfo[] dirSubInfo = myMIcroDirCopyFiles.GetDirectories();
                                    if (dirSubInfo.Count() > 0 && intOrderItemCount == 0)
                                    {
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                 "Deleted folder on location: " + Convert.ToString(ConfigurationManager.AppSettings["MicrositeFilePath"])
                                                  + " with name: " + strzipName, null, strPurchaseModuleName);
                                        // delete existing folder and create new folder
                                        myMIcroDirCopyFiles.Delete(true);
                                    }
                                }
                                #endregion
                                #region Creating folder for ProductOrderNumber
                                // It Creates a folder for ProductOrderNumber with uname_purchasedate_ordernumber

                                // strCopyZIPFile =C:\inetpub\wwwroot\WMBranding\Purchased Products\VijayNagarsoge_190916_241505
                                strCopyZIPFile = Convert.ToString(ConfigurationManager.AppSettings["compressedFilePath"]) + strzipName;

                                if (!Directory.Exists(strCopyZIPFile))
                                {
                                    myDirCopyFiles = new DirectoryInfo(strCopyZIPFile);
                                    myDirCopyFiles.Create();

                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                "Created folder on location: " + Convert.ToString(ConfigurationManager.AppSettings["compressedFilePath"])
                                                 + " with name: " + strzipName, null, strPurchaseModuleName);
                                }
                                else
                                {
                                    //If folder exists initialize object
                                    myDirCopyFiles = new DirectoryInfo(strCopyZIPFile);
                                    DirectoryInfo[] dirSubInfo = myDirCopyFiles.GetDirectories();
                                    if (dirSubInfo.Count() > 0 && intOrderItemCount == 0)
                                    {
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                "Deleted folder on location: " + Convert.ToString(ConfigurationManager.AppSettings["compressedFilePath"])
                                                 + " with name: " + strzipName, null, strPurchaseModuleName);

                                        myDirCopyFiles.Delete(true);

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                               "Create folder on location: " + Convert.ToString(ConfigurationManager.AppSettings["compressedFilePath"])
                                                + " with name: " + strzipName, null, strPurchaseModuleName);

                                        myDirCopyFiles = new DirectoryInfo(strCopyZIPFile);
                                        myDirCopyFiles.Create();

                                    }
                                }
                                #endregion
                                //parentDirectoryPath =D:\USER_DATA\Search\StoreFront\SF-Products\test-submission-5
                                parentDirectoryPath = parentFileLocation + dtOrderDetails.Rows[orderDetailsCount]["Slug"].ToString();
                                childDirectoryPath = childFileLocation + dtOrderDetails.Rows[orderDetailsCount]["Slug"].ToString();
                                // Create If condition for Course Pack changes made by Amey And Akash.
                                // variables used for coursepack function 
                                #region Local Variable Declaration
                                // TODO: Move variable declaration on Top 
                                strCoursePackSlugName = dtOrderDetails.Rows[orderDetailsCount]["Slug"].ToString();
                                strSlug = dtOrderDetails.Rows[orderDetailsCount]["Slug"].ToString();
                                strProductVariant = Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["ProductVariant"]);
                                strFileExtension = Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["FileExtension"]);
                                strProduct = Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["Product"]);
                                strUniversity = Convert.ToString(dtUserType.Rows[0]["University"]);
                                intUserType = Convert.ToInt32(dtUserType.Rows[0]["UserType"]);

                                SessionId = dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString();
                                description = HttpUtility.UrlEncode(dtOrderDetails.Rows[orderDetailsCount]["FirstName"].ToString()
                                                    + " " + dtOrderDetails.Rows[orderDetailsCount]["LastName"].ToString());

                                displayName = dtOrderDetails.Rows[orderDetailsCount]["PdfFirstName"].ToString()
                                                    + " " + dtOrderDetails.Rows[orderDetailsCount]["PdfLastName"].ToString();
                                organization = HttpUtility.UrlEncode(ConfigurationManager.AppSettings["Organization"].ToString());

                                //Forio Epicenter declare
                                strEpicenterAccount = HttpUtility.UrlEncode(ConfigurationManager.AppSettings["Account"].ToString());
                                strFirstName = dtOrderDetails.Rows[orderDetailsCount]["FirstName"].ToString();
                                strLastName = dtOrderDetails.Rows[orderDetailsCount]["LastName"].ToString();
                                strEmailId = dtOrderDetails.Rows[orderDetailsCount]["email"].ToString();

                                int intParentgroupedProductId = Convert.ToInt32(dtOrderDetails.Rows[orderDetailsCount]["ParentGroupedProductId"].ToString());

                                #endregion

                                #region Course pack product

                                // Check if purchased product is Course Pack.
                                if (string.IsNullOrEmpty(strFileExtension) && strProductVariant == strConfigCoursePack)
                                {
                                    int intCPId = Convert.ToInt32(dtOrderDetails.Rows[orderDetailsCount]["PRDID"].ToString());

                                    //Coding for coursepack 
                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                "Purchased product is course pack" + "Course Pack Id: " + intCPId, null, strPurchaseModuleName);

                                    //Insert Cource Pack Supplement Video Details In wm_subscription table
                                    #region
                                    if (strProductType == strConfigCoursePack)
                                    {

                                        InsertCoursePackSupplementVideoDetails(intOrderId, intCPId, customerId);
                                    }
                                    #endregion

                                    DataTable dtCoursePackOrderDetails = GetOrderDetailsCP(intOrderId, intCPId);

                                    string strCPDirectory = strCopyZIPFile + "\\" + strSlug;

                                    //If data table contains records
                                    if (dtCoursePackOrderDetails != null)
                                    {
                                        // If data table row count > 0
                                        if (dtCoursePackOrderDetails.Rows.Count > 0)
                                        {
                                            if (!Directory.Exists(strCPDirectory))
                                            {
                                                myDirSubFolder = new DirectoryInfo(strCPDirectory);
                                                myDirSubFolder.Create();
                                            }

                                            strCopyZIPFile = strCopyZIPFile + "\\" + strSlug;

                                            #region OrderDetailsCount
                                            for (int orderDetailsCount1 = 0; orderDetailsCount1 < dtCoursePackOrderDetails.Rows.Count; orderDetailsCount1++)
                                            {
                                                SessionId = dtCoursePackOrderDetails.Rows[orderDetailsCount1]["sessionid"].ToString();
                                                strSlug = dtCoursePackOrderDetails.Rows[orderDetailsCount1]["Slug"].ToString();
                                                string strCPproductType = dtCoursePackOrderDetails.Rows[orderDetailsCount1]["ProductType"].ToString();


                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                            "Inside CP for loop strSlug:" + strSlug, null, strPurchaseModuleName);

                                                strFileExtension = Convert.ToString(dtCoursePackOrderDetails.Rows[orderDetailsCount1]["FileExtension"]);
                                                strProduct = Convert.ToString(dtCoursePackOrderDetails.Rows[orderDetailsCount1]["Product"]);
                                                parentDirectoryPath = parentFileLocation + dtCoursePackOrderDetails.Rows[orderDetailsCount1]["Slug"].ToString();

                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                          "Inside CP for loop parentDirectoryPath:" + parentDirectoryPath, null, strPurchaseModuleName);

                                                childDirectoryPath = childFileLocation + dtCoursePackOrderDetails.Rows[orderDetailsCount1]["Slug"].ToString();
                                                intParentgroupedProductId = Convert.ToInt32(dtCoursePackOrderDetails.Rows[orderDetailsCount1]["ParentgroupedProductId"].ToString());
                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                          "Inside CP for loop childDirectoryPath:" + childDirectoryPath, null, strPurchaseModuleName);

                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                         "Inside CP for loop strCopyZIPFile:" + strCopyZIPFile, null, strPurchaseModuleName);
                                                //strCopyZIPFile=C:\inetpub\wwwroot\WMBranding\Purchased Products\VijayNagarsoge_190916_241510\cpmsit
                                                if (!Directory.Exists(strCopyZIPFile))
                                                {
                                                    myDirCopyFiles = new DirectoryInfo(strCopyZIPFile);
                                                    myDirCopyFiles.Create();

                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                "Created folder on location: " + strCopyZIPFile, null, strPurchaseModuleName);
                                                }
                                                //Check if Variant is Simulation(Course Pack License)
                                                #region Simulation in Course Pack
                                                if (Convert.ToString(dtCoursePackOrderDetails.Rows[orderDetailsCount1]["Name"])
                                                        == ConfigurationManager.AppSettings["CoursePackLicense"])// ="Course Pack License"
                                                {
                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                               "Processing Simulation product in Course Pack (Course Pack Licence) ", null, strPurchaseModuleName);

                                                    SessionId = dtCoursePackOrderDetails.Rows[orderDetailsCount1]["sessionid"].ToString();
                                                    description = "";
                                                    //HttpUtility.UrlEncode(dtOrderDetails.Rows[orderDetailsCount]["FirstName"].ToString() + " " 
                                                    //+ dtOrderDetails.Rows[orderDetailsCount]["LastName"].ToString());
                                                    organization = HttpUtility.UrlEncode(ConfigurationManager.AppSettings["Organization"].ToString());
                                                    allowUpload = "false";
                                                    maxUserCount = "0";
                                                    singleSignOnOnly = "false";
                                                    simulationUrl = "";

                                                    if (dtCoursePackOrderDetails.Rows[orderDetailsCount1]["SimulationPath"].ToString().Contains("forio.com"))
                                                    {
                                                        try
                                                        {
                                                            // TODO: Add comments and check for if statement for simulation URL
                                                            //simulationUrl = dtOrderDetails.Rows[orderDetailsCount]["SimulationPath"].ToString().Split(',')
                                                            //[0].Substring(Convert.ToInt16(ConfigurationManager.AppSettings["simurllen"])).ToString();
                                                            //simulationUrl = simulationUrl.Substring(0, simulationUrl.Length - 1);
                                                            string[] urlContaintant = dtCoursePackOrderDetails.Rows[orderDetailsCount1]["SimulationPath"].ToString()
                                                                                      .Split(',')[0].Split('/');
                                                            if (urlContaintant.Length >= 2)
                                                            {
                                                                if (urlContaintant[urlContaintant.Length - 1] == "")
                                                                {
                                                                    simulationUrl = urlContaintant[urlContaintant.Length - 3] + "/"
                                                                                    + urlContaintant[urlContaintant.Length - 2];
                                                                }
                                                                else
                                                                {
                                                                    simulationUrl = urlContaintant[urlContaintant.Length - 2] + "/"
                                                                                    + urlContaintant[urlContaintant.Length - 1];
                                                                }
                                                            }
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            fileNotFound++;
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.singleOrderCompletion",
                                                                        ex.Message, ex, strPurchaseModuleName);
                                                        }
                                                    }
                                                    else
                                                    {
                                                        simulationUrl = Convert.ToString(dtCoursePackOrderDetails.Rows[orderDetailsCount1]["SimulationPath"]);
                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                    "Simulation Path: " + simulationUrl, null, strPurchaseModuleName);
                                                    }
                                                    // TODO: else part for If condition and should be logged in as a error
                                                    if (Convert.ToString(dtCoursePackOrderDetails.Rows[orderDetailsCount1]["SimulationPath"]) != null ||
                                                           Convert.ToString(dtCoursePackOrderDetails.Rows[orderDetailsCount1]["Name"]) == strConfigCoursePackLicense)
                                                    {
                                                        try
                                                        {
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                        "Student Purchased a License", null, strPurchaseModuleName);
                                                            API api = new API(simulationUrl, ConfigurationManager.AppSettings["userName"],
                                                                              ConfigurationManager.AppSettings["password"]);
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                        "Loged-in in forio.com for creating license", null, strPurchaseModuleName);
                                                            // Code for create login for Student.

                                                            string LicenseXml = dtCoursePackOrderDetails.Rows[orderDetailsCount1]["AttributesXml"].ToString();
                                                            System.Xml.XmlTextReader reader = new System.Xml.XmlTextReader(new System.IO.StringReader(LicenseXml));
                                                            reader.Read();
                                                            System.Data.DataSet ds = new System.Data.DataSet();
                                                            ds.ReadXml(reader, System.Data.XmlReadMode.Auto);
                                                            //string simulationId = ds.Tables["ProductVariantAttributeValue"].Rows[Convert.ToInt16(
                                                            //ConfigurationManager.AppSettings["rownumSession"])]["Value"].ToString();
                                                            //string simulationId = ds.Tables["ProductVariantAttributeValue"].Rows[0]["Value"].ToString();
                                                            string simulationId = dtCoursePackOrderDetails.Rows[orderDetailsCount1]["SimulationId"].ToString();
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                        "Get the session Id from attributxml field", null, strPurchaseModuleName);
                                                            api.method = 0;
                                                            dynamic reponse2 = api.assignUserToSession(simulationId,
                                                                                   dtCoursePackOrderDetails.Rows[orderDetailsCount1]["email"].ToString(),
                                                                                   dtCoursePackOrderDetails.Rows[orderDetailsCount1]["FirstName"].ToString(),
                                                                                   dtCoursePackOrderDetails.Rows[orderDetailsCount1]["LastName"].ToString());
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                        "License created in forio.com, Session Id: " + sessions, null, strPurchaseModuleName);

                                                            //To Amey Check the value 
                                                            var statusCode = reponse2["status_code"];
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                        "Status code: " + statusCode, null, strPurchaseModuleName);
                                                            if (statusCode == "201")
                                                            {

                                                                Console.WriteLine("License Created in forio.com under Session:" + simulationId);
                                                                dynamic reponse3 = api.SendLogins(dtCoursePackOrderDetails.Rows[orderDetailsCount1]["email"].ToString());
                                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                            "User name and password information sent to user, email Id: " +
                                                                            dtCoursePackOrderDetails.Rows[orderDetailsCount1]["email"].ToString(), null, strPurchaseModuleName);
                                                                //To Amey Check the value 
                                                                statusCode = reponse3["status_code"];

                                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                      "Status code: " + statusCode, null, strPurchaseModuleName);

                                                                if (statusCode != "201")
                                                                {
                                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                                "Password information not send, email Id: " + dtOrderDetails.Rows[orderDetailsCount]
                                                                                 ["email"].ToString(), null, strPurchaseModuleName);
                                                                    string message = "Mail not send to Faculty (Error): " + reponse2["message"];
                                                                    fileNotFound++;
                                                                    Console.WriteLine("Login details sent to email address:" + dtCoursePackOrderDetails.Rows[orderDetailsCount1]
                                                                                    ["email"].ToString());

                                                                }
                                                                Insertsimulation(simulationId, "",
                                                                                     dtCoursePackOrderDetails.Rows[orderDetailsCount1]["ParentGroupedProductId"].ToString(),
                                                                                     simulationUrl, dtCoursePackOrderDetails.Rows[orderDetailsCount1]["OrderItemGuid"].ToString(),
                                                                                     dtCoursePackOrderDetails.Rows[orderDetailsCount1]["email"].ToString(),
                                                                                     dtCoursePackOrderDetails.Rows[orderDetailsCount1]["FirstName"].ToString(),
                                                                                     dtCoursePackOrderDetails.Rows[orderDetailsCount1]["LastName"].ToString(), "", "Non-Faculty");//Move to config
                                                                sessions.Add(dtCoursePackOrderDetails.Rows[orderDetailsCount1]["sessionid"].ToString());
                                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                            "Session Id added into arraylist", null, strPurchaseModuleName);
                                                                forioCount++;
                                                            }
                                                            else
                                                            {
                                                                string message = reponse2["message"];
                                                                fileNotFound++;
                                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning, "GetConnection.singleOrderCompletion",
                                                                            message, null, strPurchaseModuleName);
                                                            }
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            fileNotFound++;
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.singleOrderCompletion",
                                                                        ex.Message, ex, strPurchaseModuleName);

                                                        }
                                                    }

                                                }
                                                #endregion Simulation in Course Pack

                                                #region Microsite Inside Course Pack
                                                else if (strCPproductType == strMicrosite)
                                                {

                                                    strProductType = strMicrosite;
                                                    strProductNumber = dtCoursePackOrderDetails.Rows[orderDetailsCount1]["ProductNumber"].ToString();

                                                    Hashtable ProductHashtable = ProductProcess(fileNotFound, totalQuantity, userName, createdDate, parentFileLocation,
                                                               childFileLocation, myDirCopyFiles, myDirSupport, myDirFaculty, myDirSubFolder,
                                                               strSlug, strFileExtension, strProduct, strUniversity, intUserType, parentDirectoryPath,
                                                               strCopyZIPFile, extentCount, intOrderId, customerId, strProductVariant, SubscriptionCount, strProductType, displayName,
                                                               strzipName, intOrderItemCount, intParentgroupedProductId, strProductNumber);


                                                    OrderHashTable["extentCount"] = (int)OrderHashTable["extentCount"] + (int)ProductHashtable["extentCount"];
                                                    OrderHashTable["SubscriptionCount"] = (int)OrderHashTable["SubscriptionCount"] + (int)ProductHashtable["SubscriptionCount"];
                                                    //   OrderHashTable["transcriptCount"] = (int)OrderHashTable["transcriptCount"] + (int)ProductHashtable["transcriptCount"];
                                                    OrderHashTable["fileNotFound"] = (int)OrderHashTable["fileNotFound"] + (int)ProductHashtable["fileNotFound"];


                                                }
                                                #endregion

                                                #region PDF, EPUB, Video Product
                                                // Method "ProductProcess" for .PDF, .EPUB, Video products.
                                                else
                                                {
                                                    strProductType = dtCoursePackOrderDetails.Rows[orderDetailsCount1]["ProductType"].ToString();
                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                                           "ProductProcess method", null, strPurchaseModuleName);

                                                    Hashtable ProductHashTable = ProductProcess(fileNotFound, totalQuantity, userName, createdDate, parentFileLocation, childFileLocation,
                                                                            myDirCopyFiles, myDirSupport, myDirFaculty, myDirSubFolder, strSlug, strFileExtension, strProduct,
                                                                            strUniversity, intUserType, parentDirectoryPath, strCopyZIPFile, extentCount, intOrderId, customerId, strProductVariant
                                                                            , SubscriptionCount, strProductType, displayName, strzipName, intOrderItemCount, intParentgroupedProductId, strProductNumber);

                                                    OrderHashTable["extentCount"] = (int)OrderHashTable["extentCount"] + (int)ProductHashTable["extentCount"];
                                                    OrderHashTable["SubscriptionCount"] = (int)OrderHashTable["SubscriptionCount"] + (int)ProductHashTable["SubscriptionCount"];
                                                    //  OrderHashTable["transcriptCount"] = (int)OrderHashTable["transcriptCount"] + (int)ProductHashTable["transcriptCount"];
                                                    OrderHashTable["fileNotFound"] = (int)OrderHashTable["fileNotFound"] + (int)ProductHashTable["fileNotFound"];
                                                }
                                                #endregion PDF, EPUB, Video Product
                                            }
                                            #endregion
                                        }
                                    }
                                }

                                #endregion

                                #region Hard Copy and .PDF, .EPUB, Video Product
                                // This condition check for file extention not empty & product variant is Copyright Permissions 
                                // for hardCopy Or CopyrightPermissions , .epub, .pdf & supporting files.
                                else if (strFileExtension != string.Empty && strProductType != strMicrosite ||                  //file extention not empty and product not a microsite 
                                         strProductVariant == strConfigCopyRightPermission && strProductType != strMicrosite || // Coptright permission and not a microsite
                                         strProductVariant == strConfigMasterHardCopy && strProductType != strMicrosite ||      // Masterhardcopy and not a microsite
                                         strProductVariant == strConfigStudentHardCopy && strProductType != strMicrosite || //StudentHardCopy and not a microsite
                                          strProductVariant == strProductVariantBook && strProductType != strMicrosite)    //Product Variant Book and not a microsite    
                                {
                                    if (strProductVariant == strConfigCopyRightPermission ||
                                        strProductVariant == strConfigMasterHardCopy ||
                                        strProductVariant == strConfigStudentHardCopy ||
                                       strProductVariant == strProductVariantBook)
                                    {
                                        hardCopyOrCopyrightPermissions++;
                                    }
                                    else
                                    {
                                        //changes made by akash
                                        // Used a Product Process function insted of here code
                                        #region Used a ProductProcess function instead of code done here

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                    "Searching for folder: " + dtOrderDetails.Rows[orderDetailsCount]["Slug"].ToString() +
                                                    " exists on location: " + parentFileLocation, null, strPurchaseModuleName);
                                        Hashtable ProductHashtable = ProductProcess(fileNotFound, totalQuantity, userName, createdDate, parentFileLocation,
                                                                childFileLocation, myDirCopyFiles, myDirSupport, myDirFaculty, myDirSubFolder,
                                                                strSlug, strFileExtension, strProduct, strUniversity, intUserType, parentDirectoryPath,
                                                                strCopyZIPFile, extentCount, intOrderId, customerId, strProductVariant, SubscriptionCount, strProductType, displayName
                                                                , strzipName, intOrderItemCount, intParentgroupedProductId, strProductNumber);


                                        OrderHashTable["extentCount"] = (int)OrderHashTable["extentCount"] + (int)ProductHashtable["extentCount"];
                                        OrderHashTable["SubscriptionCount"] = (int)OrderHashTable["SubscriptionCount"] + (int)ProductHashtable["SubscriptionCount"];
                                        //   OrderHashTable["transcriptCount"] = (int)OrderHashTable["transcriptCount"] + (int)ProductHashtable["transcriptCount"];
                                        OrderHashTable["fileNotFound"] = (int)OrderHashTable["fileNotFound"] + (int)ProductHashtable["fileNotFound"];

                                        #endregion
                                        //up to here changes made
                                    }
                                }// End of electronic delivery
                                #endregion

                                #region Microsite
                                else if (strProductType == strMicrosite)
                                {
                                    //ProductProcess method, It's a common method used to process microsite and other product type except Simulation
                                    Hashtable ProductHashtable = ProductProcess(fileNotFound, totalQuantity, userName, createdDate, parentFileLocation,
                                                                childFileLocation, myDirCopyFiles, myDirSupport, myDirFaculty, myDirSubFolder,
                                                                strSlug, strFileExtension, strProduct, strUniversity, intUserType, parentDirectoryPath,
                                                                strCopyZIPFile, extentCount, intOrderId, customerId, strProductVariant, SubscriptionCount, strProductType, displayName,
                                                                strzipName, intOrderItemCount, intParentgroupedProductId, strProductNumber);


                                    OrderHashTable["extentCount"] = (int)OrderHashTable["extentCount"] + (int)ProductHashtable["extentCount"];
                                    OrderHashTable["SubscriptionCount"] = (int)OrderHashTable["SubscriptionCount"] + (int)ProductHashtable["SubscriptionCount"];
                                    OrderHashTable["fileNotFound"] = (int)OrderHashTable["fileNotFound"] + (int)ProductHashtable["fileNotFound"];
                                }
                                #endregion

                                #region Simulation Product
                                else
                                {
                                    #region Finding simulationUrl

                                    SessionId = dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString();
                                    description = "";
                                    organization = HttpUtility.UrlEncode(ConfigurationManager.AppSettings["Organization"].ToString());
                                    allowUpload = "false";
                                    maxUserCount = "0";
                                    singleSignOnOnly = "false";
                                    simulationUrl = "";
                                    if (dtOrderDetails.Rows[orderDetailsCount]["SimulationPath"].ToString().Contains("forio.com"))
                                    {
                                        try
                                        {
                                            string[] urlContaintant = dtOrderDetails.Rows[orderDetailsCount]["SimulationPath"].ToString()
                                                                      .Split(',')[0].Split('/');
                                            if (urlContaintant.Length >= 2)
                                            {
                                                if (urlContaintant[urlContaintant.Length - 1] == "")
                                                {
                                                    simulationUrl = urlContaintant[urlContaintant.Length - 3] + "/"
                                                                    + urlContaintant[urlContaintant.Length - 2];
                                                }
                                                else
                                                {
                                                    simulationUrl = urlContaintant[urlContaintant.Length - 2] + "/"
                                                                    + urlContaintant[urlContaintant.Length - 1];
                                                }
                                            }
                                        }
                                        catch (Exception ex)
                                        {
                                            fileNotFound++;
                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.getConnection",
                                                        ex.Message, ex, strPurchaseModuleName);
                                        }
                                    }
                                    else
                                    {
                                        simulationUrl = Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["SimulationPath"]);
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                    "Simulation Path: " + simulationUrl, null, strPurchaseModuleName);
                                    }

                                    #endregion


                                    strSimulationPlatform = GetSimulationPlatformType(intParentgroupedProductId);

                                    if (strSimulationPlatform.ToUpper() == strSimulationPlatformName.ToUpper())  //Simulate
                                    {

                                        #region  ProductVariant == StudentLicense

                                        if (Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["ProductVariant"]) ==
                                                    Convert.ToString(ConfigurationManager.AppSettings["StudentLicense"]))
                                        {
                                            try
                                            {
                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                            "Student Purchased a License", null, strPurchaseModuleName);
                                                API api = new API(simulationUrl, ConfigurationManager.AppSettings["userName"],
                                                                  ConfigurationManager.AppSettings["password"]);
                                                //Forio Login
                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                            "Loged-in in forio.com for creating license for OrderId: " + intOrderId, null, strPurchaseModuleName);
                                                // Code for create login for Student.

                                                string LicenseXml = dtOrderDetails.Rows[orderDetailsCount]["AttributesXml"].ToString();
                                                System.Xml.XmlTextReader reader = new System.Xml.XmlTextReader(new System.IO.StringReader(LicenseXml));
                                                reader.Read();
                                                System.Data.DataSet ds = new System.Data.DataSet();
                                                string simulationId = string.Empty;
                                                ds.ReadXml(reader, System.Data.XmlReadMode.Auto);
                                                //  string simulationId = ds.Tables["ProductVariantAttributeValue"].Rows[Convert.ToInt16(
                                                //   ConfigurationManager.AppSettings["rownumSession"])]["Value"].ToString();
                                                for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                                                {
                                                    if (intProductAttributeId == GetProductAttributeId(Convert.ToInt32(ds.Tables["ProductVariantAttribute"].Rows[i]["ID"].ToString())))
                                                    {
                                                        simulationId = ds.Tables["ProductVariantAttributeValue"].Rows[i]["Value"].ToString();
                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                    "Get the session Id from attributxml field as simulationId" + simulationId + "for OrderId:" + intOrderId, null, strPurchaseModuleName);
                                                    }
                                                }
                                                api.method = 0;
                                                dynamic reponse2 = api.assignUserToSession(simulationId,
                                                                       dtOrderDetails.Rows[orderDetailsCount]["email"].ToString(),
                                                                       dtOrderDetails.Rows[orderDetailsCount]["FirstName"].ToString(),
                                                                       dtOrderDetails.Rows[orderDetailsCount]["LastName"].ToString());
                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.getConnection",
                                                            "License created in forio.com, Session Id: " + sessions + "for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                var statusCode = reponse2["status_code"];

                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.getConnection",
                                                            "Status code: " + statusCode + "for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                //Statuscode 201 means Success
                                                if (statusCode == "201")
                                                {
                                                    Console.WriteLine("License Created in forio.com under Session:" + simulationId);
                                                    dynamic reponse3 = api.SendLogins(dtOrderDetails.Rows[orderDetailsCount]["email"].ToString());

                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.getConnection",
                                                                "User name and password information sent to user, email Id: " +
                                                                dtOrderDetails.Rows[orderDetailsCount]["email"].ToString() + "for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                    if (statusCode != "201")
                                                    {
                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                    "Password information not send, email Id: " + dtOrderDetails.Rows[orderDetailsCount]
                                                                     ["email"].ToString() + "for OrderId:" + intOrderId, null, strPurchaseModuleName);
                                                        string message = "Mail not send to Faculty (Error): " + reponse2["message"];
                                                        fileNotFound++;

                                                        Console.WriteLine("Login details sent to email address:" + dtOrderDetails.Rows[orderDetailsCount]
                                                                ["email"].ToString());
                                                    }

                                                    this.Insertsimulation(simulationId, "",
                                                                         dtOrderDetails.Rows[orderDetailsCount]["ParentGroupedProductId"].ToString(),
                                                                         simulationUrl, dtOrderDetails.Rows[orderDetailsCount]["OrderItemGuid"].ToString(),
                                                                         dtOrderDetails.Rows[orderDetailsCount]["email"].ToString(),
                                                                         dtOrderDetails.Rows[orderDetailsCount]["FirstName"].ToString(),
                                                                         dtOrderDetails.Rows[orderDetailsCount]["LastName"].ToString(), "", "Non-Faculty");

                                                    sessions.Add(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString());

                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                "Session Id added into array list", null, strPurchaseModuleName);

                                                    forioCount++;
                                                }
                                                else
                                                {
                                                    string message = reponse2["message"];
                                                    fileNotFound++;
                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning, "GetConnection.getConnection",
                                                                message, null, strPurchaseModuleName);
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                fileNotFound++;
                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.getConnection",
                                                            ex.Message, ex, strPurchaseModuleName);

                                            }
                                        }
                                        #endregion

                                        #region userType=Faculty,NonDardenFaculty,Administrator and ProductVariant= SimulationSession
                                        else if (Convert.ToInt32(dtUserType.Rows[0]["UserType"]) ==
                                                        Convert.ToInt32(ConfigurationManager.AppSettings["userTypeFaculty"]) ||
                                                        (Convert.ToInt32(dtUserType.Rows[0]["UserType"]) ==
                                                         Convert.ToInt32(ConfigurationManager.AppSettings["userTypeNonDardenFaculty"])) ||
                                                        (Convert.ToInt32(dtUserType.Rows[0]["UserType"]) ==
                                                         Convert.ToInt32(ConfigurationManager.AppSettings["Administrator"])))
                                        {
                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                        "User type is faculty for OrderId: " + intOrderId, null, strPurchaseModuleName);

                                            if (Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["ProductVariant"]) ==
                                                        Convert.ToString(ConfigurationManager.AppSettings["SimulationSession"]))
                                            {

                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                            "Purchase only simulation session or simulation session+ User license for OrderId: " + intOrderId, null, strPurchaseModuleName);

                                                string ClassDescription = dtOrderDetails.Rows[orderDetailsCount]["AttributesXml"].ToString();
                                                System.Xml.XmlTextReader reader = new System.Xml.XmlTextReader(new System.IO.StringReader(ClassDescription));
                                                reader.Read();
                                                System.Data.DataSet ds = new System.Data.DataSet();
                                                ds.ReadXml(reader, System.Data.XmlReadMode.Auto);
                                                description = ds.Tables["ProductVariantAttributeValue"].Rows[0]["Value"].ToString();

                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.getConnection",
                                                            "Get the class description from attribute xml description as :" + description + " for OrderId: " + intOrderId, null, strPurchaseModuleName);

                                                DataRow[] dr = dtOrderDetails.Select("sessionid='" +
                                                               dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString() + "'");
                                                API api = new API(simulationUrl, ConfigurationManager.AppSettings["userName"],
                                                              ConfigurationManager.AppSettings["password"]);

                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                            "Logged in forio.com using API for OrderId: " + intOrderId, null, strPurchaseModuleName);

                                                if (!sessions.Contains(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString()))
                                                {

                                                    if (dr.Length == 1)
                                                    {
                                                        //Code for create session. 
                                                        allowUpload = "false";
                                                        maxUserCount = "0";
                                                        //singleSignOnOnly = "false";
                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                                                                    "GetConnection.getConnection", "Purchase simulation session Only, allowupload="
                                                                    + allowUpload + ", maxusercound=" + maxUserCount + " for OrderId: " + intOrderId, null, strPurchaseModuleName);

                                                        postData = "description=" + description + "&organization=" + organization + "&allowUpload="
                                                                    + allowUpload + "&singleSignOnOnly=" + singleSignOnOnly + "&defaultExpirationDate="
                                                                    + System.DateTime.Now.AddMonths(6).ToString("MM/dd/yyyy");
                                                    }
                                                    else
                                                    {
                                                        DataRow[] dr1 = dtOrderDetails.Select("sessionid='" + dtOrderDetails.Rows[orderDetailsCount]
                                                                        ["sessionid"].ToString() + "' and ProductVariant='" +
                                                                        ConfigurationManager.AppSettings["UserLicense"] + "'");

                                                        allowUpload = "true";
                                                        maxUserCount = dr1[0]["Quantity"].ToString();
                                                        singleSignOnOnly = "false";
                                                        api.method = 1;
                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                    "Purchase user license allowupload= " + allowUpload + ", maxusercount= " + maxUserCount + " for OrderId: " + intOrderId, null,
                                                                    strPurchaseModuleName);

                                                        postData = "description=" + description + "&organization=" + organization +
                                                                   "&allowUpload=" + allowUpload + "&singleSignOnOnly=" + singleSignOnOnly +
                                                                   "&maxUserCount=" + maxUserCount + "&defaultExpirationDate=" +
                                                                    System.DateTime.Now.AddMonths(6).ToString("MM/dd/yyyy");
                                                    }
                                                    try
                                                    {
                                                        var response = api.createSession(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString(), postData);
                                                        var statusCode = response["status_code"];
                                                        dynamic reponse2 = null;
                                                        //var sessionResponse=
                                                        if (statusCode == "201")
                                                        {
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                        "Session created in forio.com, SessionId:" + dtOrderDetails.Rows[orderDetailsCount]
                                                                        ["sessionid"].ToString() + "for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                            Console.WriteLine("Session Created in forio.com, Session:" + dtOrderDetails.Rows[orderDetailsCount]
                                                                    ["sessionid"].ToString() + ", maxUserCount=" + maxUserCount);

                                                            var userId = response["group"]["id"];
                                                            reponse2 = api.assignAdminToSession(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString(),
                                                                       dtOrderDetails.Rows[orderDetailsCount]["email"].ToString(),
                                                                       dtOrderDetails.Rows[orderDetailsCount]["FirstName"].ToString(),
                                                                       dtOrderDetails.Rows[orderDetailsCount]["LastName"].ToString());

                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                        "Administrator assigned to session: emailId: " + dtOrderDetails.Rows[orderDetailsCount]
                                                                        ["email"].ToString() + ", sessionId: " + SessionId + "for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                            statusCode = reponse2["status_code"];
                                                            if (statusCode == "201")
                                                            {
                                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                            "Faculty assigned to session for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                                dynamic reponse3 = api.SendLogins(dtOrderDetails.Rows[orderDetailsCount]["email"].ToString());

                                                                Console.WriteLine("Login details sent to email address:" + dtOrderDetails.Rows[orderDetailsCount]
                                                                        ["email"].ToString());

                                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                            "Forio Login details send to email:" + dtOrderDetails.Rows[orderDetailsCount]
                                                                            ["email"].ToString() + "for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                                sessions.Add(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString());
                                                                Insertsimulation(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString(),
                                                                                     userId, dtOrderDetails.Rows[orderDetailsCount]["ParentGroupedProductId"].ToString(),
                                                                                     simulationUrl, dtOrderDetails.Rows[orderDetailsCount]["OrderItemGuid"].ToString(),
                                                                                     dtOrderDetails.Rows[orderDetailsCount]["email"].ToString(),
                                                                                     dtOrderDetails.Rows[orderDetailsCount]["FirstName"].ToString(),
                                                                                     dtOrderDetails.Rows[orderDetailsCount]["LastName"].ToString(), "",
                                                                                     "Faculty");

                                                                if (allowUpload == "false")
                                                                {
                                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                                "Purchased only simulation session, System sending email for simulation link to faculty for OrderId:" + intOrderId,
                                                                                null, strPurchaseModuleName);

                                                                    //Send storefront link to factuly for automatic purchase of License.
                                                                    simulationUrl = ConfigurationManager.AppSettings["StorefrontSite"] + "session=" +
                                                                                    dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString();


                                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                                "simulation URL ::" + simulationUrl + "for OrderId:" + intOrderId,
                                                                                null, strPurchaseModuleName);

                                                                    SendSimulationMailToFaculty(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString(),
                                                                                        dtOrderDetails.Rows[orderDetailsCount]["FirstName"].ToString() + " " +
                                                                                        dtOrderDetails.Rows[orderDetailsCount]["LastName"].ToString(),
                                                                                        dtOrderDetails.Rows[orderDetailsCount]["email"].ToString(),
                                                                                        simulationUrl, dtOrderDetails.Rows[orderDetailsCount]["Name"].ToString(),
                                                                                        dtOrderDetails.Rows[orderDetailsCount]["Id"].ToString(),
                                                                                        dtOrderDetails.Rows[orderDetailsCount]["AttributeDescription"].ToString());
                                                                }
                                                                forioCount++;
                                                                this.InsertSimulationToLibrary((Convert.ToInt32(customerId)), SessionId);

                                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                           "Simulation product added to library for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                            }
                                                            else
                                                            {
                                                                string message = reponse2["message"];
                                                                fileNotFound++;
                                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                            message, null, strPurchaseModuleName);
                                                            }
                                                        }
                                                        else
                                                        {
                                                            string message = reponse2["message"];
                                                            fileNotFound++;
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                        message, null, strPurchaseModuleName);
                                                        }
                                                    }
                                                    catch (Exception ex)
                                                    {

                                                        fileNotFound++;
                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.getConnection",
                                                                    ex.Message, ex, strPurchaseModuleName);

                                                    }

                                                }
                                            }
                                            else
                                            {
                                                try
                                                {
                                                    if (Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["ProductVariant"]) ==
                                                                Convert.ToString(ConfigurationManager.AppSettings["UserLicense"]))
                                                    {
                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                    "Faculty purchase user license for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                        DataRow[] dr1 = dtOrderDetails.Select("sessionid='" + dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString()
                                                                        + "' and ProductVariant='" + ConfigurationManager.AppSettings["SimulationSession"] + "'");
                                                        API api = new API(simulationUrl, ConfigurationManager.AppSettings["userName"], ConfigurationManager.AppSettings["password"]);

                                                        string ClassDescription = dr1[0]["AttributesXml"].ToString(); //dtOrderDetails.Rows[orderDetailsCount]["AttributesXml"].ToString();
                                                        System.Xml.XmlTextReader reader = new System.Xml.XmlTextReader(new System.IO.StringReader(ClassDescription));
                                                        reader.Read();
                                                        System.Data.DataSet ds = new System.Data.DataSet();
                                                        ds.ReadXml(reader, System.Data.XmlReadMode.Auto);
                                                        description = ds.Tables["ProductVariantAttributeValue"].Rows[0]["Value"].ToString();

                                                        if (!sessions.Contains(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString()))
                                                        {
                                                            maxUserCount = dtOrderDetails.Rows[orderDetailsCount]["quantity"].ToString();
                                                            //singleSignOnOnly = "false";
                                                            allowUpload = "true";
                                                            api.method = 1;
                                                            dynamic reponse2 = null;
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection", "Maxusercount:"
                                                                        + maxUserCount + ", allowupload: " + allowUpload + "for OrderId:" + intOrderId, null, strPurchaseModuleName);
                                                            try
                                                            {
                                                                var response = api.createSession(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString(),
                                                                               "description=" + description + "&organization=" + organization + "&allowUpload=" +
                                                                               allowUpload + "&singleSignOnOnly=" + singleSignOnOnly + "&maxUserCount=" + maxUserCount
                                                                               + "&defaultExpirationDate=" + System.DateTime.Now.AddMonths(6).ToString("MM/dd/yyyy"));

                                                                var statusCode = response["status_code"];
                                                                if (statusCode == "201")
                                                                {
                                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                                "Session created in forio.com for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                                    Console.WriteLine("Session Created in forio.com, Session:" + dtOrderDetails.Rows[orderDetailsCount]
                                                                            ["sessionid"].ToString() + ", maxUserCount=" + maxUserCount);

                                                                    var userId = response["group"]["id"];
                                                                    reponse2 = api.assignAdminToSession(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString(),
                                                                                    dtOrderDetails.Rows[orderDetailsCount]["email"].ToString(),
                                                                                    dtOrderDetails.Rows[orderDetailsCount]["FirstName"].ToString(),
                                                                                    dtOrderDetails.Rows[orderDetailsCount]["LastName"].ToString());
                                                                    statusCode = reponse2["status_code"];
                                                                    if (statusCode == "201")
                                                                    {
                                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                                   "Admin/Faculty assign to session in forio.com for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                                        dynamic reponse3 = api.SendLogins(dtOrderDetails.Rows[orderDetailsCount]["email"].ToString());

                                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                                    "Login details send to user", null, strPurchaseModuleName);

                                                                        Console.WriteLine("Login details sent to email address:" +
                                                                                           dtOrderDetails.Rows[orderDetailsCount]["email"].ToString());

                                                                        this.Insertsimulation(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString(),
                                                                                            userId, dtOrderDetails.Rows[orderDetailsCount]["ParentGroupedProductId"].ToString(),
                                                                                            simulationUrl, dtOrderDetails.Rows[orderDetailsCount]["OrderItemGuid"].ToString(),
                                                                                            dtOrderDetails.Rows[orderDetailsCount]["email"].ToString(),
                                                                                            dtOrderDetails.Rows[orderDetailsCount]["FirstName"].ToString(),
                                                                                            dtOrderDetails.Rows[orderDetailsCount]["LastName"].ToString(), "", "Faculty");
                                                                        sessions.Add(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString());
                                                                        forioCount++;

                                                                    }
                                                                    else
                                                                    {
                                                                        string message = reponse2["message"];
                                                                        fileNotFound++;
                                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                                    message, null, strPurchaseModuleName);

                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    string message = reponse2["message"];
                                                                    fileNotFound++;
                                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                                message, null, strPurchaseModuleName);
                                                                }
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                fileNotFound++;
                                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.getConnection", ex.Message,
                                                                            ex, strPurchaseModuleName);

                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        hardCopyOrCopyrightPermissions++;
                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                    "Purchased product is: " + Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["ProductVariant"]) +
                                                                    ". Incrementing count of non electric download found to: " + hardCopyOrCopyrightPermissions,
                                                                    null, strPurchaseModuleName);

                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    SendOrderStuckEmail(Convert.ToInt32(intOrderId), strName, OrderHashTable);

                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.getConnection",
                                                                     ex.Message, null, strPurchaseModuleName);

                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                    "Purchased product is: " + Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["ProductVariant"]) +
                                                                    "Problem processing order containing user license. Order no :" + intOrderId,
                                                                    null, strPurchaseModuleName);
                                                    //Email to Admin
                                                    break;
                                                }
                                            }
                                        }
                                        else
                                        {
                                            hardCopyOrCopyrightPermissions++;
                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                        "Purchased product is: " + Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["ProductVariant"]) + ". Incrementing count of non electric download found to: " + hardCopyOrCopyrightPermissions, null, strPurchaseModuleName);

                                        }

                                        #endregion

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                     "Searching for folder: " + dtOrderDetails.Rows[orderDetailsCount]["Slug"].ToString() +
                                                     " exists on location: " + parentFileLocation, null, strPurchaseModuleName);

                                        Hashtable ProductHashtable = ProductProcess(fileNotFound, totalQuantity, userName, createdDate, parentFileLocation,
                                                                childFileLocation, myDirCopyFiles, myDirSupport, myDirFaculty, myDirSubFolder,
                                                                strSlug, strFileExtension, strProduct, strUniversity, intUserType, parentDirectoryPath,
                                                                strCopyZIPFile, extentCount, intOrderId, customerId, strProductVariant, SubscriptionCount, strProductType, displayName
                                                                , strzipName, intOrderItemCount, intParentgroupedProductId, strProductNumber);


                                        OrderHashTable["extentCount"] = (int)OrderHashTable["extentCount"] + (int)ProductHashtable["extentCount"];
                                        OrderHashTable["SubscriptionCount"] = (int)OrderHashTable["SubscriptionCount"] + (int)ProductHashtable["SubscriptionCount"];

                                        OrderHashTable["fileNotFound"] = (int)OrderHashTable["fileNotFound"] + (int)ProductHashtable["fileNotFound"];
                                    }
                                    else if (strSimulationPlatform.ToUpper() == strEpicenterPlatformName.ToUpper())    //Epicenter
                                    {
                                        #region  ProductVariant == SimulationSession (Facuty user) For epicenter type

                                        if (Convert.ToInt32(dtUserType.Rows[0]["UserType"]) ==
                                            Convert.ToInt32(ConfigurationManager.AppSettings["userTypeFaculty"]) ||
                                           (Convert.ToInt32(dtUserType.Rows[0]["UserType"]) ==
                                             Convert.ToInt32(ConfigurationManager.AppSettings["userTypeNonDardenFaculty"])) ||
                                           (Convert.ToInt32(dtUserType.Rows[0]["UserType"]) ==
                                             Convert.ToInt32(ConfigurationManager.AppSettings["Administrator"])))
                                        {

                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                        "User type is faculty for epicenter platform for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                            #region ProductVariant = SimulationSession
                                            if (Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["ProductVariant"]) ==
                                                        Convert.ToString(ConfigurationManager.AppSettings["SimulationSession"]))
                                            {

                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                            "Purchase Epicenter simulation session/simulation session + User license for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                string ClassDescription = dtOrderDetails.Rows[orderDetailsCount]["AttributesXml"].ToString();

                                                System.Xml.XmlTextReader reader = new System.Xml.XmlTextReader(new System.IO.StringReader(ClassDescription));
                                                reader.Read();
                                                System.Data.DataSet ds = new System.Data.DataSet();
                                                ds.ReadXml(reader, System.Data.XmlReadMode.Auto);
                                                description = ds.Tables["ProductVariantAttributeValue"].Rows[0]["Value"].ToString();

                                                strGroupName = "DBP-" + intOrderId + "-" + strUniversity;
                                                strForioName = dtOrderDetails.Rows[orderDetailsCount]["Name"].ToString();
                                                strProjectFormatId = strForioName.ToLower() + "-" + description;

                                                DataRow[] dr = dtOrderDetails.Select("sessionid='" +
                                                               dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString() + "'");

                                                bool status = CheckOAuthAPICallOrNot(orderDetailsCount, intOrderId, strOAuthToken, strFirstName, strLastName, strEmailId, customerId,
                                                                                   strForioName, strEpicenterAccount, strUniversity, simulationUrl, SessionId, strFacultyUser);
                                                if (status == true)
                                                {
                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.getConnection",
                                                    "Successfully called OAuth API for OrderId:" + intOrderId, null, strPurchaseModuleName);
                                                }
                                                else
                                                {
                                                    fileNotFound++;
                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning, "GetConnection.getConnection",
                                                    "After trying 3 times, call to OAuth API failed for OrderId:" + intOrderId, null, strPurchaseModuleName);
                                                    goto End;
                                                }
                                                #endregion
                                            }
                                            else
                                            {
                                                #region ProductVariant = UserLicense

                                                try
                                                {
                                                    if (Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["ProductVariant"]) ==
                                                               Convert.ToString(ConfigurationManager.AppSettings["UserLicense"]))
                                                    {
                                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                    "Faculty purchase user license for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                        sessions.Add(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString());

                                                        if (!sessions.Contains(dtOrderDetails.Rows[orderDetailsCount]["sessionid"].ToString()))
                                                        {
                                                            bool Status = CallEpicenterAPI(orderDetailsCount, intOrderId, strOAuthToken, strFirstName, strLastName, strEmailId,
                                                                                           customerId, strForioName, strEpicenterAccount, strUniversity, simulationUrl,
                                                                                           SessionId, strFacultyUser);
                                                        }
                                                        else
                                                        {
                                                            hardCopyOrCopyrightPermissions++;
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                        "Purchased product is: " + Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["ProductVariant"]) +
                                                                        ". Incrementing count of non electric download found to: " + hardCopyOrCopyrightPermissions,
                                                                        null, strPurchaseModuleName);
                                                        }
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    SendOrderStuckEmail(Convert.ToInt32(intOrderId), strName, OrderHashTable);
                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.getConnection",
                                                                     ex.Message, null, strPurchaseModuleName);

                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                    "Purchased product is: " + Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["ProductVariant"]) +
                                                                    "Problem processing order containing user license. Order no :" + intOrderId,
                                                                    null, strPurchaseModuleName);
                                                    break;
                                                }
                                                #endregion
                                            }
                                        }

                                        #region ProductVariant = StudentLicense
                                        else if (Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["ProductVariant"]) ==
                                                    Convert.ToString(ConfigurationManager.AppSettings["StudentLicense"]))
                                        {
                                            try
                                            {
                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                           "Student Purchased a License for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                EpicenterAPI api = new EpicenterAPI(strEpicenterPublicKey, strEpicenterSecretKey);

                                                dynamic tokenresponse = api.createoauthtoken();
                                                strOAuthToken = tokenresponse["access_token"];
                                                //strStatusCode = tokenresponse["statuscode"].ToString();
                                                try
                                                {
                                                    strRefreshToken = tokenresponse["refresh_token"].ToString();

                                                    if (!string.IsNullOrEmpty(strRefreshToken))
                                                    {
                                                        strStatusCode = "200";
                                                    }
                                                    else
                                                    {
                                                        strStatusCode = "";
                                                    }

                                                }
                                                catch (Exception ex)
                                                {
                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.CallEpicenterAPI",
                                                                "The key name as refresh_token  is not found in API response hense API call is not successful ", ex, strPurchaseModuleName);
                                                    strStatusCode = "";
                                                }
                                                if (strStatusCode == "200")
                                                {
                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.getConnection",
                                                           "Generate epicenter OAuth Token :" + strOAuthToken + "for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                    string LicenseXml = dtOrderDetails.Rows[orderDetailsCount]["AttributesXml"].ToString();
                                                    System.Xml.XmlTextReader reader = new System.Xml.XmlTextReader(new System.IO.StringReader(LicenseXml));
                                                    reader.Read();
                                                    System.Data.DataSet ds = new System.Data.DataSet();
                                                    //string simulationId = string.Empty;
                                                    ds.ReadXml(reader, System.Data.XmlReadMode.Auto);
                                                    for (int i = 0; i < ds.Tables[0].Rows.Count; i++)
                                                    {
                                                        if (intProductAttributeId == GetProductAttributeId(Convert.ToInt32(ds.Tables["ProductVariantAttribute"].Rows[i]["ID"].ToString())))
                                                        {
                                                            strGroupId = ds.Tables["ProductVariantAttributeValue"].Rows[i]["Value"].ToString();
                                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                                        "Get the group id from attributexml field :" + strGroupId + "for OrderId:" + intOrderId, null, strPurchaseModuleName);

                                                            bool Status = AddUserAPI(orderDetailsCount, intOrderId, strOAuthToken, strFirstName, strLastName,
                                                                                     strEmailId, customerId, strForioName, strEpicenterAccount, strUniversity,
                                                                                     simulationUrl, SessionId, strProjectFullName, strGroupId, strStudentUserRole, strStudentUser);
                                                            if (Status == true)
                                                            {
                                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.getConnection",
                                                                "Successfully called student license API for OrderId:" + intOrderId, null, strPurchaseModuleName);
                                                                forioCount++;
                                                            }
                                                            else
                                                            {
                                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning, "GetConnection.getConnection",
                                                                "Failed to call student license API for OrderId:" + intOrderId, null, strPurchaseModuleName);
                                                                goto End;
                                                            }
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    strMessage = tokenresponse["message"];

                                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.getConnection",
                                                                            "(Oauth token)error is : " + strMessage, null, strPurchaseModuleName);
                                                    fileNotFound++;
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                fileNotFound++;
                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.getConnection",
                                                            ex.Message, ex, strPurchaseModuleName);
                                            }
                                        }
                                        else
                                        {
                                            hardCopyOrCopyrightPermissions++;
                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                                        "Purchased product is: " + Convert.ToString(dtOrderDetails.Rows[orderDetailsCount]["ProductVariant"])
                                                         + ". Incrementing count of non electric download found to: "
                                                         + hardCopyOrCopyrightPermissions, null, strPurchaseModuleName);
                                        }
                                        #endregion
                                        //  }

                                        #endregion


                                        Hashtable ProductHashtable = ProductProcess(fileNotFound, totalQuantity, userName, createdDate, parentFileLocation,
                                                                childFileLocation, myDirCopyFiles, myDirSupport, myDirFaculty, myDirSubFolder,
                                                                strSlug, strFileExtension, strProduct, strUniversity, intUserType, parentDirectoryPath,
                                                                strCopyZIPFile, extentCount, intOrderId, customerId, strProductVariant, SubscriptionCount, strProductType, displayName,
                                                                strzipName, intOrderItemCount, intParentgroupedProductId, strProductNumber);


                                        OrderHashTable["extentCount"] = (int)OrderHashTable["extentCount"] + (int)ProductHashtable["extentCount"];
                                        OrderHashTable["SubscriptionCount"] = (int)OrderHashTable["SubscriptionCount"] + (int)ProductHashtable["SubscriptionCount"];
                                        OrderHashTable["fileNotFound"] = (int)OrderHashTable["fileNotFound"] + (int)ProductHashtable["fileNotFound"];
                                    }

                                    //  }

                                }
                                #endregion
                                //
                                intOrderItemCount++;

                            End:
                                Console.WriteLine("Take next order.....");

                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                                            "Sale Completion Process take next order.", null, strPurchaseModuleName);


                            }//Individual order
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.singleOrderCompletion",
                                                       ex.Message, ex, strPurchaseModuleName);
                        }
                        #region file found
                        // if (fileNotFound == 0)
                        try
                        {
                            if ((int)OrderHashTable["fileNotFound"] == 0)
                            {
                                DataRow[] resultCopy = dtOrderDetails.Select("Id=" + intOrderId);
                                dtTarget = dtOrderDetails.Copy().Select("Id=" + intOrderId).CopyToDataTable();


                                // if (extentCount >= 1 && hardCopyOrCopyrightPermissions >= 1 && forioCount >= 1)
                                #region if Extent, Video, Hardcopy, Forio
                                if ((int)OrderHashTable["extentCount"] >= 1 && (int)OrderHashTable["SubscriptionCount"] >= 1 && hardCopyOrCopyrightPermissions >= 1 && forioCount >= 1)
                                {
                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                "Zipping folder:" + myDirCopyFiles.ToString(), null, strPurchaseModuleName);

                                    string zipFilePath = myDirCopyFiles + ".zip";
                                    using (ZipFile zip = new ZipFile())
                                    {
                                        try
                                        {

                                            ////zip.Password = DecryptMaskedCard(Convert.ToString(dt.Rows[pendingOrderCount]["Password"]));
                                            zip.AddDirectory(myDirCopyFiles.ToString());
                                            zip.Save(zipFilePath);
                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.singleOrderCompletion",
                                             "Successfully included all products in zip.", null, strPurchaseModuleName);
                                        }
                                        catch (Exception ex)
                                        {
                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.singleOrderCompletion",
                                                        "Error: " + ex.Message, null, strPurchaseModuleName);
                                        }

                                    }

                                    ClearFolder(myDirCopyFiles.ToString());
                                    myDirCopyFiles.Delete();

                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                "Temporary created folder: " + myDirCopyFiles.ToString() + " deleted successfully.",
                                                null, strPurchaseModuleName);

                                    string strCopyrightPermissions = "'" + Convert.ToString(ConfigurationManager.AppSettings["CopyrightPermissions"]) + "'";
                                    DataRow[] result = dtTarget.Select("ProductVariant=" + strCopyrightPermissions + "");

                                    if (result.Length == 0)
                                    {
                                        if (SendMailToCustomerProductsPurchased(strZipPath, Convert.ToInt32(intOrderId), strFileExtension, strName, OrderHashTable))
                                        {
                                            if (Convert.ToString(dtOrder.Rows[0]["PaymentMethodSystemName"]) ==
                                                        Convert.ToString(ConfigurationManager.AppSettings["PaymentMethodSystemName"]))
                                            {
                                                UpdateOrderStatusWithHardCopy(Convert.ToInt32(intOrderId));
                                            }
                                            this.InsertCompletedOrders(intOrderId);
                                        }
                                    }
                                    else
                                    {
                                        if (SendMailToCustomerProductsPurchased(strZipPath, Convert.ToInt32(intOrderId), strFileExtension, strName, OrderHashTable))
                                        {
                                            if (Convert.ToString(dtOrder.Rows[0]["PaymentMethodSystemName"]) ==
                                                        Convert.ToString(ConfigurationManager.AppSettings["PaymentMethodSystemName"]))
                                            {
                                                UpdateOrderStatusWithOnlyExtension(Convert.ToInt32(intOrderId));
                                            }
                                            this.InsertCompletedOrders(intOrderId);
                                        }
                                    }
                                }
                                #endregion
                                else
                                {
                                    #region Extent, Video OR Extent && Forio
                                    if ((hardCopyOrCopyrightPermissions == 0 && (int)OrderHashTable["extentCount"] >= 1 && (int)OrderHashTable["SubscriptionCount"] >= 1 && forioCount == 0)
                                        || (hardCopyOrCopyrightPermissions == 0 && (int)OrderHashTable["extentCount"] >= 1 && (int)OrderHashTable["SubscriptionCount"] == 0 && forioCount >= 1))
                                    {
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                    "Zipping folder:" + myDirCopyFiles.ToString(), null, strPurchaseModuleName);

                                        string zipFilePath = myDirCopyFiles + ".zip";
                                        using (ZipFile zip = new ZipFile())
                                        {
                                            try
                                            {
                                                ////zip.Password = DecryptMaskedCard(Convert.ToString(dt.Rows[pendingOrderCount]["Password"]));
                                                zip.AddDirectory(myDirCopyFiles.ToString());
                                                zip.Save(zipFilePath);
                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.singleOrderCompletion",
                                                 "Successfully included all products in zip.", null, strPurchaseModuleName);

                                            }
                                            catch (Exception ex)
                                            {
                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.singleOrderCompletion",
                                                            "Error: " + ex.Message, null, strPurchaseModuleName);
                                            }

                                        }

                                        ClearFolder(myDirCopyFiles.ToString());
                                        myDirCopyFiles.Delete();

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                    "Temporary created folder: " + myDirCopyFiles.ToString() + " deleted successfully.",
                                                    null, strPurchaseModuleName);

                                        if (SendMailToCustomerProductsPurchased(strZipPath, Convert.ToInt32(intOrderId), strFileExtension, strName, OrderHashTable))
                                        {
                                            if (Convert.ToString(dtOrder.Rows[0]["PaymentMethodSystemName"]) ==
                                                        Convert.ToString(ConfigurationManager.AppSettings["PaymentMethodSystemName"]))
                                            {
                                                UpdateOrderStatusWithOnlyExtension(Convert.ToInt32(intOrderId));
                                            }
                                            this.InsertCompletedOrders(intOrderId);
                                        }


                                    }
                                    #endregion

                                    #region Hardcopy, ExtentCount, Video
                                    if (hardCopyOrCopyrightPermissions >= 1 && (int)OrderHashTable["extentCount"] >= 1 && (int)OrderHashTable["SubscriptionCount"] >= 1 && forioCount == 0)
                                    {
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                    "Zipping folder:" + myDirCopyFiles.ToString(), null, strPurchaseModuleName);

                                        string zipFilePath = myDirCopyFiles + ".zip";
                                        using (ZipFile zip = new ZipFile())
                                        {
                                            try
                                            {
                                                ////zip.Password = DecryptMaskedCard(Convert.ToString(dt.Rows[pendingOrderCount]["Password"]));
                                                zip.AddDirectory(myDirCopyFiles.ToString());
                                                zip.Save(zipFilePath);
                                            }
                                            catch (Exception ex)
                                            {
                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.singleOrderCompletion",
                                                            "Error: " + ex.Message, null, strPurchaseModuleName);
                                            }

                                        }
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.singleOrderCompletion",
                                                    "Successfully included all products in zip.", null, strPurchaseModuleName);

                                        ClearFolder(myDirCopyFiles.ToString());
                                        myDirCopyFiles.Delete();

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                    "Temporary created folder: " + myDirCopyFiles.ToString() + " deleted successfully.",
                                                    null, strPurchaseModuleName);

                                        string strCopyrightPermissions = "'" + Convert.ToString(ConfigurationManager.AppSettings["CopyrightPermissions"]) + "'";
                                        DataRow[] result = dtTarget.Select("ProductVariant=" + strCopyrightPermissions + "");

                                        if (result.Length == 0)
                                        {
                                            if (SendMailToCustomerProductsPurchased(strZipPath, Convert.ToInt32(intOrderId), strFileExtension, strName, OrderHashTable))
                                            {
                                                if (Convert.ToString(dtOrder.Rows[0]["PaymentMethodSystemName"]) ==
                                                            Convert.ToString(ConfigurationManager.AppSettings["PaymentMethodSystemName"]))
                                                {
                                                    UpdateOrderStatusWithHardCopy(Convert.ToInt32(intOrderId));
                                                }
                                                this.InsertCompletedOrders(intOrderId);
                                            }
                                        }
                                        else
                                        {
                                            if (SendMailToCustomerProductsPurchased(strZipPath, Convert.ToInt32(intOrderId), strFileExtension, strName, OrderHashTable))
                                            {
                                                if (Convert.ToString(dtOrder.Rows[0]["PaymentMethodSystemName"]) ==
                                                            Convert.ToString(ConfigurationManager.AppSettings["PaymentMethodSystemName"]))
                                                {
                                                    UpdateOrderStatusWithOnlyExtension(Convert.ToInt32(intOrderId));
                                                }
                                                this.InsertCompletedOrders(intOrderId);
                                            }
                                        }
                                    }
                                    #endregion

                                    #region Extent, Video, Forio
                                    if (hardCopyOrCopyrightPermissions == 0 && (int)OrderHashTable["extentCount"] >= 1 && (int)OrderHashTable["SubscriptionCount"] >= 1 && forioCount >= 1)
                                    {

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                    "Zipping folder:" + myDirCopyFiles.ToString(), null, strPurchaseModuleName);
                                        string zipFilePath = myDirCopyFiles + ".zip";
                                        using (ZipFile zip = new ZipFile())
                                        {
                                            ////zip.Password = DecryptMaskedCard(Convert.ToString(dt.Rows[pendingOrderCount]["Password"]));
                                            zip.AddDirectory(myDirCopyFiles.ToString());
                                            //zip.Password("");
                                            zip.Save(zipFilePath);
                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.singleOrderCompletion",
                                                   "Successfully included all products in zip.", null, strPurchaseModuleName);

                                        }

                                        ClearFolder(myDirCopyFiles.ToString());
                                        myDirCopyFiles.Delete();

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                    "Temporary created folder: " + myDirCopyFiles.ToString() + " deleted successfully.", null, strPurchaseModuleName);

                                        if (SendMailToCustomerProductsPurchased(strZipPath, Convert.ToInt32(intOrderId), strFileExtension, strName, OrderHashTable))
                                        {
                                            if (Convert.ToString(dtOrder.Rows[0]["PaymentMethodSystemName"]) ==
                                                        Convert.ToString(ConfigurationManager.AppSettings["PaymentMethodSystemName"]))
                                            {
                                                UpdateOrderStatusWithOnlyExtension(Convert.ToInt32(intOrderId));
                                            }
                                            this.InsertCompletedOrders(intOrderId);
                                        }
                                    }
                                    #endregion

                                    #region Forio OR Video OR Video && Forio
                                    if ((hardCopyOrCopyrightPermissions == 0 && (int)OrderHashTable["extentCount"] == 0 && (int)OrderHashTable["SubscriptionCount"] == 0 && forioCount >= 1)
                                        || (hardCopyOrCopyrightPermissions == 0 && (int)OrderHashTable["extentCount"] == 0 && (int)OrderHashTable["SubscriptionCount"] >= 1 && forioCount == 0
                                        || (hardCopyOrCopyrightPermissions == 0 && (int)OrderHashTable["extentCount"] == 0 && (int)OrderHashTable["SubscriptionCount"] >= 1 && forioCount >= 1)))
                                    {
                                        int fileCount = Directory.GetFiles(myDirCopyFiles.FullName).Length;
                                        if (fileCount > 0)
                                        {
                                            string zipFilePath = myDirCopyFiles + ".zip";
                                            using (ZipFile zip = new ZipFile())
                                            {
                                                ////zip.Password = DecryptMaskedCard(Convert.ToString(dt.Rows[pendingOrderCount]["Password"]));
                                                zip.AddDirectory(myDirCopyFiles.ToString());
                                                //zip.Password("");
                                                zip.Save(zipFilePath);
                                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.singleOrderCompletion",
                                                       "Successfully included all products in zip.", null, strPurchaseModuleName);

                                            }
                                        }

                                        if (SendMailToCustomerProductsPurchased(strZipPath, Convert.ToInt32(intOrderId), strFileExtension, strName, OrderHashTable))
                                        {
                                            if (Convert.ToString(dtOrder.Rows[0]["PaymentMethodSystemName"]) ==
                                                        Convert.ToString(ConfigurationManager.AppSettings["PaymentMethodSystemName"]))
                                            {
                                                UpdateOrderStatusWithOnlyExtension(Convert.ToInt32(intOrderId));
                                            }
                                            this.InsertCompletedOrders(intOrderId);
                                        }
                                        //Condition != changes to == empty
                                        //if (Convert.ToString(myDirCopyFiles) != "")
                                        //{
                                        //    ClearFolder(myDirCopyFiles.ToString());
                                        //    myDirCopyFiles.Delete();
                                        //}

                                    }
                                    #endregion

                                    #region Hardcopy, Video, Forio OR Harcopy && Video OR Hardcopy OR Hardcopy && Forio
                                    if ((hardCopyOrCopyrightPermissions >= 1 && (int)OrderHashTable["extentCount"] == 0 && (int)OrderHashTable["SubscriptionCount"] >= 1 && forioCount >= 1)
                                        || (hardCopyOrCopyrightPermissions >= 1 && (int)OrderHashTable["extentCount"] == 0 && (int)OrderHashTable["SubscriptionCount"] >= 1 && forioCount == 0)
                                        || (hardCopyOrCopyrightPermissions >= 1 && (int)OrderHashTable["extentCount"] == 0 && (int)OrderHashTable["SubscriptionCount"] == 0 && forioCount >= 1))
                                    {
                                        string strCopyrightPermissions = "'" + Convert.ToString(ConfigurationManager.AppSettings["CopyrightPermissions"]) + "'";
                                        DataRow[] result = dtTarget.Select("ProductVariant=" + strCopyrightPermissions + "");

                                        if (result.Length == 0)
                                        {
                                            if (SendMailToCustomerProductsPurchased(strZipPath, Convert.ToInt32(intOrderId), strFileExtension, strName, OrderHashTable))
                                            {
                                                if (Convert.ToString(dtOrder.Rows[0]["PaymentMethodSystemName"]) ==
                                                            Convert.ToString(ConfigurationManager.AppSettings["PaymentMethodSystemName"]))
                                                {
                                                    UpdateOrderStatusWithHardCopy(Convert.ToInt32(intOrderId));
                                                }
                                                this.InsertCompletedOrders(intOrderId);
                                            }
                                        }
                                        else
                                        {
                                            if (SendMailToCustomerProductsPurchased(strZipPath, Convert.ToInt32(intOrderId), strFileExtension, strName, OrderHashTable))
                                            {
                                                if (Convert.ToString(dtOrder.Rows[0]["PaymentMethodSystemName"]) ==
                                                            Convert.ToString(ConfigurationManager.AppSettings["PaymentMethodSystemName"]))
                                                {
                                                    UpdateOrderStatusWithOnlyExtension(Convert.ToInt32(intOrderId));
                                                }
                                                this.InsertCompletedOrders(intOrderId);
                                            }
                                        }
                                    }
                                    #endregion

                                    #region Extent
                                    if (hardCopyOrCopyrightPermissions == 0 && (int)OrderHashTable["extentCount"] >= 1 && (int)OrderHashTable["SubscriptionCount"] == 0 && forioCount == 0)
                                    {
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                 "Zipping folder:" + myDirCopyFiles.ToString(), null, strPurchaseModuleName);
                                        string zipFilePath = myDirCopyFiles + ".zip";
                                        using (ZipFile zip = new ZipFile())
                                        {
                                            ////zip.Password = DecryptMaskedCard(Convert.ToString(dt.Rows[pendingOrderCount]["Password"]));
                                            zip.AddDirectory(myDirCopyFiles.ToString());
                                            //zipFilePath= C:\inetpub\wwwroot\WMBranding\Purchased Products\VijayNagarsoge_190925_242521.zip
                                            zip.Save(zipFilePath);
                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.singleOrderCompletion",
                                                   "Successfully included all products in zip.", null, strPurchaseModuleName);

                                        }

                                        ClearFolder(myDirCopyFiles.ToString());
                                        myDirCopyFiles.Delete();

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                    "Temporary created folder: " + myDirCopyFiles.ToString() + " deleted successfully.", null, strPurchaseModuleName);


                                        if (SendMailToCustomerProductsPurchased(strZipPath, Convert.ToInt32(intOrderId), strFileExtension, strName, OrderHashTable))
                                        {
                                            if (Convert.ToString(dtOrder.Rows[0]["PaymentMethodSystemName"]) ==
                                                        Convert.ToString(ConfigurationManager.AppSettings["PaymentMethodSystemName"]))
                                            {
                                                UpdateOrderStatusWithOnlyExtension(Convert.ToInt32(intOrderId));
                                            }
                                            this.InsertCompletedOrders(intOrderId);
                                        }


                                    }
                                    #endregion

                                    #region Hardcopy
                                    if ((hardCopyOrCopyrightPermissions >= 1 && (int)OrderHashTable["extentCount"] == 0 && (int)OrderHashTable["SubscriptionCount"] == 0 && forioCount == 0))
                                    {
                                        if (SendMailToCustomerProductsPurchased(strZipPath, Convert.ToInt32(intOrderId), strFileExtension, strName, OrderHashTable))
                                        {
                                            if (Convert.ToString(dtOrder.Rows[0]["PaymentMethodSystemName"]) ==
                                                        Convert.ToString(ConfigurationManager.AppSettings["PaymentMethodSystemName"]))
                                            {
                                                UpdateOrderStatusWithHardCopy(Convert.ToInt32(intOrderId));
                                            }
                                            this.InsertCompletedOrders(intOrderId);

                                            if (Convert.ToString(myDirCopyFiles) != "")
                                            {
                                                DirectoryInfo directory = new DirectoryInfo(myDirCopyFiles.ToString());
                                                if (directory.Exists)
                                                {
                                                    ClearFolder(myDirCopyFiles.ToString());
                                                    myDirCopyFiles.Delete();
                                                }
                                            }
                                        }
                                    }
                                    #endregion

                                    #region Hardcopy, Extent OR Hardcopy && Extent && Forio
                                    if ((hardCopyOrCopyrightPermissions >= 1 && (int)OrderHashTable["extentCount"] >= 1 && (int)OrderHashTable["SubscriptionCount"] == 0 && forioCount == 0)
                                        || (hardCopyOrCopyrightPermissions >= 1 && (int)OrderHashTable["extentCount"] >= 1 && (int)OrderHashTable["SubscriptionCount"] == 0 && forioCount >= 1))
                                    {
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                 "Zipping folder:" + myDirCopyFiles.ToString(), null, strPurchaseModuleName);
                                        string zipFilePath = myDirCopyFiles + ".zip";
                                        using (ZipFile zip = new ZipFile())
                                        {
                                            ////zip.Password = DecryptMaskedCard(Convert.ToString(dt.Rows[pendingOrderCount]["Password"]));
                                            zip.AddDirectory(myDirCopyFiles.ToString());
                                            //zip.Password("");
                                            zip.Save(zipFilePath);
                                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.singleOrderCompletion",
                                                   "Successfully included all products in zip.", null, strPurchaseModuleName);

                                        }

                                        ClearFolder(myDirCopyFiles.ToString());
                                        myDirCopyFiles.Delete();

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                    "Temporary created folder: " + myDirCopyFiles.ToString() + " deleted successfully.", null, strPurchaseModuleName);


                                        string strCopyrightPermissions = "'" + Convert.ToString(ConfigurationManager.AppSettings["CopyrightPermissions"]) + "'";
                                        DataRow[] result = dtTarget.Select("ProductVariant=" + strCopyrightPermissions + "");

                                        if (result.Length == 0)
                                        {
                                            if (SendMailToCustomerProductsPurchased(strZipPath, Convert.ToInt32(intOrderId), strFileExtension, strName, OrderHashTable))
                                            {
                                                if (Convert.ToString(dtOrder.Rows[0]["PaymentMethodSystemName"]) ==
                                                            Convert.ToString(ConfigurationManager.AppSettings["PaymentMethodSystemName"]))
                                                {
                                                    UpdateOrderStatusWithHardCopy(Convert.ToInt32(intOrderId));
                                                }
                                                this.InsertCompletedOrders(intOrderId);
                                            }
                                        }
                                        else
                                        {
                                            if (SendMailToCustomerProductsPurchased(strZipPath, Convert.ToInt32(intOrderId), strFileExtension, strName, OrderHashTable))
                                            {
                                                if (Convert.ToString(dtOrder.Rows[0]["PaymentMethodSystemName"]) ==
                                                            Convert.ToString(ConfigurationManager.AppSettings["PaymentMethodSystemName"]))
                                                {
                                                    UpdateOrderStatusWithOnlyExtension(Convert.ToInt32(intOrderId));
                                                }
                                                this.InsertCompletedOrders(intOrderId);
                                            }
                                        }

                                    }
                                    #endregion


                                }
                            }

                            #endregion
                            else
                            #region file not found
                            {
                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                   "File not found fileNotFound =" + fileNotFound, null, strPurchaseModuleName);

                                if (!string.IsNullOrEmpty(Convert.ToString(myDirCopyFiles)))
                                {
                                    DirectoryInfo directory = new DirectoryInfo(myDirCopyFiles.ToString());
                                    if (directory.Exists)
                                    {
                                        ClearFolder(myDirCopyFiles.ToString());
                                        myDirCopyFiles.Delete();
                                    }
                                }
                            }
                            #endregion
                        }
                        catch (Exception ex)
                        {
                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.singleOrderCompletion",
                                                       ex.Message, ex, strPurchaseModuleName);
                        }
                    }
                    else
                    {
                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                               "No order details found for Order Id : " + intOrderId + "",
                               null, strPurchaseModuleName);

                    }

                    strfilenotfound = "FileNotFound Count is : " + (int)OrderHashTable["fileNotFound"];
                    OrderHashTable.Clear();
                }
                else
                {
                    Console.WriteLine("CustomerID:" + Convert.ToInt32(customerId) + "  is not found in wm_UVAUserInformation table.");

                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning, "GetConnection.singleOrderCompletion",
                                "CustomerID:" + Convert.ToInt32(customerId) + " is not found in wm_UVAUserInformation table.",
                                null, strPurchaseModuleName);
                }

                //    }//for end
            }
            #endregion
            else
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Warning,
                        "GetConnection.singleOrderCompletion",
                        "No record found for order Id : " + intOrderId + " in order table",
                         null, strPurchaseModuleName);
            }

        }
        catch (SqlException e)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.singleOrderCompletion",
                        "In singleOrderCompletion() methods Catch block. Error: " + e.Message, e, strPurchaseModuleName);
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.singleOrderCompletion",
                        "In singleOrderCompletion() methods Catch block. Error: " + ex.Message, ex, strPurchaseModuleName);
        }

        return strfilenotfound;
    }

    #endregion

    #endregion

    #region Utilities



    public string ToCurrencyString(decimal value)
    {
        //ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.cs, 
        //public string ToCurrencyString()", "Getting $ symbol from ToCurrencyString() method.", null, strPurchaseModuleName);
        //if (value == 0)
        //    return String.Empty;
        // return value.ToString("C");

        return value.ToString("C", CultureInfo.CreateSpecificCulture(ConfigurationManager.AppSettings["Currency"]));


    }

    public void EditPaths(string fileName, string word, string replacement, string elementId, string replaceProdNumber, string saveFileName, string directory, string destination, int intUserType, string strProductNumber)
    {
        //Variable Declaration
        string strJSReference = "<" + strScriptTagStart + @"""" + strJSScript + @"""" + strSrc + @"""" + includeFolderName + "/" + strProductNumber + appJSFileExtension + @"""" + "><" + strScriptTagEnd + ">";
        StreamReader reader = new StreamReader(directory + fileName);
        string input = reader.ReadToEnd();

        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                    "File read Complete" + "File Name: " + fileName, null, strPurchaseModuleName);
        string fileJS = directory + fileName;


        // If file extention is .HTML or .HTM
        if ((fileName.Contains(msFileExtention1)) || (fileName.Contains(msFileExtention2)))
        {
            using (StreamWriter writer = new StreamWriter(destination + saveFileName, true))
            {
                //String preOutput = input.Replace(word, replacement);
                String appendOut = input + strJSReference;
                HtmlDocument doc = new HtmlDocument();
                doc.LoadHtml(appendOut);
                try
                {
                    //If user type is not FACULTY or NON-DardenFACULTY or Administrator
                    if (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["userTypeFaculty"]) ||
                           (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["userTypeNonDardenFaculty"])) ||
                           (Convert.ToInt32(intUserType) == Convert.ToInt32(ConfigurationManager.AppSettings["Administrator"])))
                    {
                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                    "User Is Admin or Faculty or Nondarden Faculty" + "User type: " + intUserType, null, strPurchaseModuleName);
                    }
                    else
                    {
                        var nodes = doc.DocumentNode.SelectNodes("//div[@id='OnlyForFaculty']");
                        foreach (HtmlNode node in nodes)
                        {
                            //Remove <div id="OnlyForFaculty"></div>
                            node.Remove();
                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                                   "User Is Not Athorised to get supporting document,Div tag removed" + "User type: " + intUserType, null, strPurchaseModuleName);
                        }

                    }
                }
                catch (Exception ex)
                {

                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.singleOrderCompletion",
                        "In getConnection() methods Catch block.HTML does not contain Div tag. Error: " + ex.Message, ex, strPurchaseModuleName);
                }
                //Create Html document
                String output = doc.DocumentNode.InnerHtml;
                writer.Write(output);
                writer.Close();
            }
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                   "Words replaced in File" + " File Name: " + fileName, null, strPurchaseModuleName);
        }
        else
        {
            //Append Java script for multiple video's
            int startIndex = input.IndexOf("jwplayer");
            int intstartIndex = input.IndexOf("function");
            int intLastindex = input.LastIndexOf('}');

            //old path with new path and element with Product Number
            string output = input.Remove(intLastindex, 1).Insert(startIndex, Environment.NewLine).Replace(elementId, replaceProdNumber).Replace(word, replacement);
            output = output.Remove(intstartIndex, strstartJSPoint.Count());//.Replace(strstartJSPoint, strBlank);
            sbJwplayer.Append(output);
        }
    }

    private string DecryptMaskedCard(string cipherText)
    {
        try
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.DecryptMaskedCard",
                        "Decrypting value: " + cipherText + " by using HARDCODE encryption private Key.", null, strPurchaseModuleName);
            string encryptionPrivateKey = "273ece6f97dd844d";//273ece6f97dd844d
            var tDESalg = new TripleDESCryptoServiceProvider();
            tDESalg.Key = new ASCIIEncoding().GetBytes(encryptionPrivateKey.Substring(0, 16));
            tDESalg.IV = new ASCIIEncoding().GetBytes(encryptionPrivateKey.Substring(8, 8));

            byte[] buffer = Convert.FromBase64String(cipherText);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.DecryptMaskedCard",
                        "Decrypted value: " + cipherText + " by using HARDCODE encryption private Key.", null, strPurchaseModuleName);
            return DecryptTextFromMemory(buffer, tDESalg.Key, tDESalg.IV);
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.DecryptMaskedCard",
                        "Error: " + ex.Message, null, strPurchaseModuleName);
            return string.Empty;
        }
    }

    private string DecryptTextFromMemory(byte[] data, byte[] key, byte[] iv)
    {
        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.cs,private string DecryptTextFromMemory()",
                    "Decrypt Text From Memory.", null, strPurchaseModuleName);
        using (var ms = new MemoryStream(data))
        {
            using (var cs = new CryptoStream(ms, new TripleDESCryptoServiceProvider().CreateDecryptor(key, iv), CryptoStreamMode.Read))
            {
                var sr = new StreamReader(cs, new UnicodeEncoding());

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.DecryptTextFromMemory",
                   sr.ReadLine(), null, strPurchaseModuleName);

                return sr.ReadLine();
            }
        }
    }

    public string Alignment(int rowIndex)
    {

        string alignment = string.Empty;

        //ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.Alignment",
        //            "alignment =" + alignment, null, strPurchaseModuleName);

        if (rowIndex.Equals(0))
            alignment = "left";
        else
            alignment = "right";

        return alignment;
    }

    private void ClearFolder(string FolderName)
    {
        DirectoryInfo dir = new DirectoryInfo(FolderName);
        foreach (FileInfo fi in dir.GetFiles())
        {
            fi.IsReadOnly = false;
            fi.Delete();
        }
        foreach (DirectoryInfo di in dir.GetDirectories())
        {
            ClearFolder(di.FullName);
            di.Delete();
        }
        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ClearFolder",
                    "Deleting folder which is created temporary for making proper zip structure.", null, strPurchaseModuleName);
    }

    //public void DigitalRights(string fileNameWithPath, string userName, int totalCopies, string purchasedDate, int intOrderId, string instituteName)
    //{
    //    try
    //    {
    //        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.DigitalRights",
    //                    "Start adding Digital rights to " + fileNameWithPath, null, strPurchaseModuleName);

    //        string oldFile = fileNameWithPath;
    //        string newFile = fileNameWithPath + "1";
    //        int PageCount = 0;
    //        //oldFile=C:\inetpub\wwwroot\WMBranding\Purchased Products\VijayNagarsoge_190916_241506\test-submission-5\JH-0046.pdf
    //        if (IsValidPdf(oldFile))
    //        {
    //            // open the reader
    //            PdfReader reader = new PdfReader(oldFile);
    //            Rectangle size = reader.GetPageSizeWithRotation(1);
    //            Document document = new Document(size);

    //            BaseFont bfTimes = BaseFont.CreateFont(BaseFont.TIMES_ROMAN, BaseFont.CP1252, false);
    //            Font times = new Font(bfTimes, 9, Font.NORMAL, BaseColor.BLACK);
    //            Font whiteText = new Font(bfTimes, 9, Font.NORMAL, BaseColor.WHITE);
    //            string text = string.Empty;
    //            //if (totalCopies > 1)

    //            ////text = " DBPOrderNumber: " + intOrderId;
    //            //Added bu Harish Patil On 05-29-2013 for changing the hidden text.
    //            text = ConfigurationManager.AppSettings["HiddenText"] + intOrderId;

    //            //else
    //            //    text = userName + " has purchased " + totalCopies + " copy on " + purchasedDate;
    //            PdfPCell cell = new PdfPCell();
    //            cell.HorizontalAlignment = Element.ALIGN_CENTER;
    //            cell.VerticalAlignment = Element.ALIGN_MIDDLE;

    //            string txt = string.Empty;

    //            if (string.IsNullOrEmpty(instituteName))
    //            {
    //                txt = Convert.ToString(ConfigurationManager.AppSettings["DRMText1"]) + userName + ".";
    //            }
    //            else
    //            {
    //                txt = Convert.ToString(ConfigurationManager.AppSettings["DRMText1"]) + userName +
    //                      Convert.ToString(ConfigurationManager.AppSettings["DRMText2"]) + instituteName + ".";
    //            }

    //            string txtLine = Convert.ToString(ConfigurationManager.AppSettings["DRMText3"]);

    //            // int yPos = Element.ALIGN_MIDDLE;

    //            using (FileStream fs = new FileStream(newFile, FileMode.Create, FileAccess.Write, FileShare.None))
    //            {
    //                using (PdfStamper stamper = new PdfStamper(reader, fs))
    //                {
    //                    PageCount = reader.NumberOfPages;
    //                    for (int i = 1; i <= PageCount; i++)
    //                    {
    //                        ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_CENTER, new
    //                            Phrase(String.Format(text, i, PageCount), whiteText), 150, 750, 0);
    //                        ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_CENTER,
    //                            new Phrase(String.Format(txtLine, i, PageCount), times), 30, 400, 90);
    //                        ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_BOTTOM,
    //                            new Phrase(String.Format(txt, i, PageCount), times), 20, 245, 90);
    //                        //ColumnText.ShowTextAligned(stamper.GetOverContent(i), 
    //                        //Element.ALIGN_CENTER, new Phrase(String.Format(text, i, PageCount), times), 150, 40, 0);
    //                        //ColumnText.ShowTextAligned(stamper.GetOverContent(i), Element.ALIGN_CENTER,
    //                        //    new Phrase(String.Format("Page {0} of {1}", i, PageCount), times), 320, 40, 0);

    //                    }
    //                }
    //            }
    //            File.Delete(oldFile);
    //            File.Copy(newFile, oldFile);
    //            File.Delete(newFile);
    //            PdfReader newReader = new PdfReader(oldFile);
    //            FileStream fstream = new FileStream(newFile, FileMode.Create, FileAccess.Write, FileShare.None);
    //            PdfStamper pdfstamper = new PdfStamper(newReader, fstream);
    //            ////pdfstamper.SetEncryption(true, null, "secret", PdfWriter.AllowScreenReaders);

    //             pdfstamper.SetEncryption(true, null, "secret", PdfWriter.ALLOW_PRINTING | PdfWriter.AllowScreenReaders | PdfWriter.ALLOW_MODIFY_ANNOTATIONS);



    //            pdfstamper.Close();
    //            newReader.Close();

    //            File.Delete(oldFile);
    //            File.Copy(newFile, oldFile);
    //            File.Delete(newFile);
    //            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.DigitalRights",
    //                            "Digital rights added to " + fileNameWithPath + " successfully", null, strPurchaseModuleName);
    //            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.DigitalRights",
    //                            "End adding Digital rights to " + fileNameWithPath, null, strPurchaseModuleName);
    //        }
    //    }
    //    catch (Exception ex)
    //    {
    //        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.DigitalRights",
    //                    "Error while adding Digital rights to" + fileNameWithPath + ex.Message, ex, strPurchaseModuleName);
    //    }
    //}


    public void DigitalRights(string fileNameWithPath, string userName, int totalCopies, string purchasedDate, int intOrderId, string instituteName)
    {
        try
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.DigitalRights",
                        "Start adding Digital rights to " + fileNameWithPath, null, strPurchaseModuleName);

            string oldFile = fileNameWithPath;
            string newFile = fileNameWithPath + "1";

            // Open the existing PDF
            PdfDocument document = PdfReader.Open(oldFile, PdfDocumentOpenMode.Modify);

            string text = ConfigurationManager.AppSettings["HiddenText"] + intOrderId;
            string txt = string.Empty;
            if (string.IsNullOrEmpty(instituteName))
            {
                txt = Convert.ToString(ConfigurationManager.AppSettings["DRMText1"]) + userName + ".";
            }
            else
            {
                txt = Convert.ToString(ConfigurationManager.AppSettings["DRMText1"]) + userName +
                      Convert.ToString(ConfigurationManager.AppSettings["DRMText2"]) + instituteName + ".";
            }

            string txtLine = Convert.ToString(ConfigurationManager.AppSettings["DRMText3"]);

            // Define the watermark text
            string watermarkText = txt + "\n" + txtLine;

            foreach (PdfPage page in document.Pages)
            {
                // Get the page dimensions
                XGraphics gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Prepend);

                // Create a font for the watermark
                XFont font = new XFont("Times New Roman", 10);

                // Define the brush for watermark color (semi-transparent gray)
                XBrush brush = new XSolidBrush(XColor.FromArgb(255, 0, 0, 0));

                // Get page dimensions
                double pageHeight = page.Height.Point;
                double pageWidth = page.Width.Point;

                // Calculate usable height with 10% margin from top and bottom
                double topMargin = pageWidth * 0.03; // 10% from the top
                double bottomMargin = pageWidth * 0.10; // 10% from the bottom
                double usableHeight = pageWidth - (topMargin + bottomMargin); // Usable space between margins

                // Save the current graphics state
                gfx.Save();

                // Rotate the graphics context to make the text vertical
                gfx.RotateTransform(-90); // Rotate counterclockwise 90 degrees
                gfx.TranslateTransform(-pageHeight + 150, topMargin); // Move closer to the left edge

                // Draw the watermark text in paragraph format
                XTextFormatter textFormatter = new XTextFormatter(gfx);
                XRect textRect = new XRect(0, 0, usableHeight, pageHeight - 80); // Define the text area bounds
                textFormatter.Alignment = XParagraphAlignment.Center;
                textFormatter.DrawString(watermarkText, font, brush, textRect);

                // Restore the graphics state
                gfx.Restore();
            }
            // Save the document to a temporary file
            document.Save(newFile);

            // Replace the original file with the new one
            File.Delete(oldFile);
            File.Copy(newFile, oldFile);
            File.Delete(newFile);

            // Encrypt the PDF with permissions
            PdfDocument encryptedDocument = PdfReader.Open(oldFile, PdfDocumentOpenMode.Modify);
            encryptedDocument.SecuritySettings.OwnerPassword = "secret";
            encryptedDocument.SecuritySettings.PermitPrint = true;
            encryptedDocument.SecuritySettings.PermitAnnotations = false; // Allows adding comments or annotations
            encryptedDocument.SecuritySettings.PermitModifyDocument = false;
            encryptedDocument.SecuritySettings.PermitExtractContent = false; // Disallow extracting text or images
            encryptedDocument.SecuritySettings.PermitFormsFill = false;         // Disallow form filling
            encryptedDocument.SecuritySettings.PermitFullQualityPrint = true;   //Disallow full-quality printing
            // Save the encrypted file
            encryptedDocument.Save(newFile);

            // Replace the original file with the encrypted one
            File.Delete(oldFile);
            File.Copy(newFile, oldFile);
            File.Delete(newFile);

            Console.WriteLine("Watermarked PDF saved.");

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.DigitalRights",
                "Digital rights added to " + fileNameWithPath + " successfully", null, strPurchaseModuleName);
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.DigitalRights",
                        "Error while adding Digital rights to" + fileNameWithPath + ex.Message, ex, strPurchaseModuleName);
        }
    }
    //private bool IsValidPdf(string filepath)
    //{
    //    bool Ret = true;

    //    PdfReader reader = null;

    //    try
    //    {
    //        reader = new PdfReader(filepath);
    //    }
    //    catch
    //    {
    //        Ret = false;
    //    }

    //    return Ret;
    //}

    //public void VerticalText(PdfWriter writer, string text, Document doc)
    //{
    //    doc.Open();
    //    PdfContentByte cb = writer.DirectContent;
    //    ColumnText ct = new ColumnText(cb);
    //    ct.SetSimpleColumn(new Phrase(new Chunk(text, FontFactory.GetFont(FontFactory.HELVETICA, 18, Font.NORMAL))),
    //                       46, 190, 530, 36, 25, Element.ALIGN_LEFT | Element.ALIGN_TOP);
    //    ct.Go();
    //    doc.Close();
    //}

    #endregion

    #region Email

    public void sendMailToAdministratorFileNotFound(string productName, string parentDirectoryPath, int intOrderId, string slugName)
    {
        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.sendMailToAdministratorFileNotFound",
                    "Sending email to Administrator because product: (" + productName + ") not found on physical location"
                    + parentDirectoryPath + " for order: " + intOrderId, null, strPurchaseModuleName);

        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.sendMailToAdministratorFileNotFound",
                   "slugName=" + slugName, null, strPurchaseModuleName);
        try
        {
            string productURL = ConfigurationManager.AppSettings["HomePageUrl"].ToString() + slugName;
            string oldPath = ConfigurationManager.AppSettings["SendMailProductNotFound"].ToString();
            StreamReader reader = new StreamReader(oldPath);
            string content = reader.ReadToEnd();
            reader.Close();

            productURL = "<a href=" + productURL + ">" + productName + "</a>";

            content = content.Replace("<strProductName>", productURL);
            content = content.Replace("<strPhysicalPath>", parentDirectoryPath);
            content = content.Replace("<DateTime>", DateTime.Now.ToString());
            content = content.Replace("<strOrderNumber>", Convert.ToString(intOrderId));

            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient(ConfigurationManager.AppSettings["SmtpHost"]);
            mail.From = new MailAddress(ConfigurationManager.AppSettings["MailFrom"]);
            //DataTable custInfo = GetCustomerInfo(customerId);
            mail.To.Add(ConfigurationManager.AppSettings["SendProductNotfoundMail"].ToString());
            mail.Subject = ConfigurationManager.AppSettings["ProductNotFoundSubject"];
            mail.IsBodyHtml = true;
            mail.Body = Convert.ToString(content);

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            SmtpServer.Port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]);
            SmtpServer.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["NetworkCredentialMailID"],
                       ConfigurationManager.AppSettings["NetworkCredentialMailIDPassword"]);
            SmtpServer.EnableSsl = true;

            Console.WriteLine("Start mail sending for product not found on physical location.");
            SmtpServer.Send(mail);
            Console.WriteLine("Email sending for product not found on physical location.");
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.sendMailToAdministratorFileNotFound",
                        "Error while sending email to Administrator if file not found. Error:" + ex.Message, ex, strPurchaseModuleName);
        }
    }


    public bool SendMailToCustomerProductsPurchased(string FilePath, int OrderId, string strFileExtension, string ParentDirectoryPath, Hashtable countHashTable)
    {

        string cardName = string.Empty;
        string cardType = string.Empty;
        string creditCardMasked = string.Empty;
        string cardExpirationMonth = string.Empty;
        string cardExpirationYear = string.Empty;
        string[] lstfile;
        List<string> strExten = new List<string>();
        List<string> strJSUrl = new List<string>();
        string strExtention = string.Empty;
        string strExtent;
        string strJsFilePath = string.Empty;
        //string strJsFileExt = ".js";
        string strJsFileExtention = string.Empty;
        string strfileName = string.Empty;
        string strSupportingFileName = string.Empty;
        StringBuilder mailBody = new StringBuilder();
        string _emailFormatFileLocation = ConfigurationManager.AppSettings["EmailBody"];
        string strZipFileName = string.Empty;
        string strJsFileName = string.Empty;
        string strDownloadProduct = string.Empty;
        String line = String.Empty;
        int intProductId;
        //video          
        string _VideoURL = string.Empty;

        try
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased", "Sending mail for the OrderNumber:"
                        + OrderId, null, strPurchaseModuleName);

            int shippingAddressId = 0;
            DataTable OrderData = GetOrderProductData(OrderId);
            DataTable dtOrderDetailsCP = GetOrderDetails(OrderId);
            int customerId = Convert.ToInt32(OrderData.Rows[0]["CustomerId"]);
            int billingAddressId = Convert.ToInt32(OrderData.Rows[0]["BillingAddressId"]);
            string customerCurrencyCode = Convert.ToString(OrderData.Rows[0]["CustomerCurrencyCode"]);

            if (!DBNull.Value.Equals((OrderData.Rows[0]["ShippingAddressId"])))
            {
                shippingAddressId = Convert.ToInt32(OrderData.Rows[0]["ShippingAddressId"]);
            }
            int CustomerId = Convert.ToInt32(OrderData.Rows[0]["CustomerId"]);

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                        "Getting email id from customer table for CustomerId: " + customerId, null, strPurchaseModuleName);
            DataTable customerInfo = GetCustomerEmail(customerId);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                        "Getting billing address from Address table", null, strPurchaseModuleName);
            DataTable billingAddress = GetCustomerAddress(billingAddressId);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                        "Getting shipping address from Address table", null, strPurchaseModuleName);
            DataTable shippingAddress = GetCustomerAddress(shippingAddressId);

            DataTable productDetails = GetProductDetails(OrderId);
            DataTable productUrl = GetOrderDetails(OrderId);
            //DataTable VideoCoursePack = FilOrderDeatilsCP(OrderId, intCPId);



            if (Convert.ToString(OrderData.Rows[0]["CardName"]) != string.Empty)
            {
                cardName = DecryptMaskedCard(Convert.ToString(OrderData.Rows[0]["CardName"]));
                cardType = DecryptMaskedCard(Convert.ToString(OrderData.Rows[0]["CardType"]));
                creditCardMasked = DecryptMaskedCard(Convert.ToString(OrderData.Rows[0]["MaskedCreditCardNumber"]));
                cardExpirationMonth = DecryptMaskedCard(Convert.ToString(OrderData.Rows[0]["CardExpirationMonth"]));
                cardExpirationYear = DecryptMaskedCard(Convert.ToString(OrderData.Rows[0]["CardExpirationYear"]));
            }//string CardType = encryptionService.DecryptText(cipherText, string.Empty);, [Order].CardExpirationYear
            else
            {
                if (Convert.ToString(OrderData.Rows[0]["MaskedCreditCardNumber"]) == "")
                {
                    creditCardMasked = ConfigurationManager.AppSettings["ShippingMethodsStatus"].ToString();
                }
                else
                {
                    creditCardMasked = DecryptMaskedCard(Convert.ToString(OrderData.Rows[0]["MaskedCreditCardNumber"]));
                }

            }

            List<string> ProductVariantId = new List<string>();
            List<string> ProductName = new List<string>();

            string paymentMethodFromDB = Convert.ToString(OrderData.Rows[0]["PaymentMethodSystemName"]);
            string paymentMethod = string.Empty;
            if (paymentMethodFromDB != "")
            {
                paymentMethod = paymentMethodFromDB.Remove(0, 9);
            }
            else
            {
                paymentMethod = ConfigurationManager.AppSettings["ShippingMethodsStatus"].ToString();
            }

            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient(ConfigurationManager.AppSettings["SmtpHost"]);
            mail.From = new MailAddress(ConfigurationManager.AppSettings["MailFrom"]);
            mail.To.Add(customerInfo.Rows[0]["Email"].ToString());
            mail.Subject = ConfigurationManager.AppSettings["EmailSubject"];
            mail.IsBodyHtml = true;

            ////ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased", 
            //"In try catch block of SendMailToCustomerProductsPurchased() method assigned network credentials.", null, strPurchaseModuleName);

            strExtent = "*" + Convert.ToString(strFileExtension).ToLower();

            // Added by Harish Patil for download file.            
            if (File.Exists(FilePath + ".zip"))
            {
                strZipFileName = Path.GetFileName(FilePath + ".zip");
                strDownloadProduct = ConfigurationManager.AppSettings["DownloadProduct"].ToString() + strZipFileName;
            }

            if (File.Exists(_emailFormatFileLocation))
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                            "HTML Format of email body exists on its location.", null, strPurchaseModuleName);

                string emailFormatFileLocation = ConfigurationManager.AppSettings["EmailBody"].ToString();
                StreamReader reader = new StreamReader(emailFormatFileLocation);
                string content = reader.ReadToEnd();
                reader.Close();


                char cr = (char)13;

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                            "Reading each line using stream reader of HTML body in SendMailToCustomerProductsPurchased() method.", null, strPurchaseModuleName);

                #region MailBody Append


                //lstfile = String.Empty;//strFileExtension));
                DataTable table = new DataTable("ProductsTableURL");
                DataColumn URL = new DataColumn("URL", typeof(System.String));
                //ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTax"])));
                table.Columns.Add(" ");

                //FROM CONFIGURATION FILE
                content = content.Replace("<homepageurl>", Convert.ToString(ConfigurationManager.AppSettings["HomePageUrl"]));
                content = content.Replace("<YourAccountOrders>", Convert.ToString(ConfigurationManager.AppSettings["YourAccountOrders"]));
                content = content.Replace("<youraccountordersname>", Convert.ToString(ConfigurationManager.AppSettings["YourAccountOrdersName"]));

                if (strDownloadProduct != "" || (int)countHashTable["extentCount"] > 0)
                {
                    strDownloadProduct = "\"" + strDownloadProduct + "\"";
                    string strReplace = "<a href=" + strDownloadProduct + " style='width: 400px' >Please click here to download PDF or ePub product(s).</a>";
                    //   content = content.Replace("<downloadfile>", strReplace);

                    DataRow downloadRow = table.NewRow();
                    downloadRow.ItemArray = new object[] { strReplace };
                    table.Rows.Add(downloadRow);
                }
                // Payment Method data
                if (Convert.ToString(OrderData.Rows[0]["PurchaseOrderNumber"]) != string.Empty)
                {
                    content = content.Replace("Card Type:", "Purchase Order No.");
                    content = content.Replace("<cardtype>", Convert.ToString(OrderData.Rows[0]["PurchaseOrderNumber"]));
                    //Card Number:Expiry:
                    content = content.Replace("Card Number:", cr.ToString());
                    content = content.Replace("Expiry:", cr.ToString());
                    content = content.Replace("<cardExpirationMonth>", cr.ToString());
                    content = content.Replace("<cardExpirationYear>", cr.ToString());

                }
                else
                {
                    content = content.Replace("Card Type:", "");
                    content = content.Replace("Expiry:", "");
                    content = content.Replace("<creditCardMasked>", creditCardMasked);
                }

                //TEMPORARY FROM SHIPPING ADDRESS TABLE AFTERWARDS CHANGE IT AND ACCESS USERNAME FROM GENERICATTRIBUTE TABLE
                //content = content.Replace("<UserName>", Convert.ToString(billingAddress.Rows[0]["UserName"]));
                content = content.Replace("<UserName>", Convert.ToString(displayName));
                //Billing Address
                content = content.Replace("<busername>", Convert.ToString(billingAddress.Rows[0]["UserName"]));
                content = content.Replace("<BPhone>", Convert.ToString(billingAddress.Rows[0]["PhoneNumber"]));
                content = content.Replace("<BFax>", Convert.ToString(billingAddress.Rows[0]["FaxNumber"]));

                if (content.Contains("<BCompany>"))
                {
                    if (Convert.ToString(billingAddress.Rows[0]["Company"]).Equals(string.Empty))
                    {
                        content = content.Replace("<BCompany><br />", cr.ToString());
                    }
                    else
                    {
                        content = content.Replace("<BCompany>", Convert.ToString(billingAddress.Rows[0]["Company"]));
                    }
                }
                content = content.Replace("<BAddress1>", Convert.ToString(billingAddress.Rows[0]["Address1"]));

                if (content.Contains("BAddress2"))
                {
                    if (Convert.ToString(billingAddress.Rows[0]["Address2"]).Equals(string.Empty))
                    {
                        content = content.Replace("<BAddress2><br />", cr.ToString());
                    }
                    else
                    {
                        content = content.Replace("<BAddress2>", Convert.ToString(billingAddress.Rows[0]["Address2"]));
                    }
                }
                if (Convert.ToString(billingAddress.Rows[0]["City"]) != string.Empty)
                {
                    content = content.Replace("<BCity,>", Convert.ToString(billingAddress.Rows[0]["City"]) + ",");
                }

                if (Convert.ToString(billingAddress.Rows[0]["State"]) != string.Empty)
                {
                    content = content.Replace("<BState,>", Convert.ToString(billingAddress.Rows[0]["State"]) + ",");
                }
                content = content.Replace("<BZip>", Convert.ToString(billingAddress.Rows[0]["ZipPostalCode"]));
                content = content.Replace("<BCountry>", Convert.ToString(billingAddress.Rows[0]["Country"]));
                content = content.Replace("<id>", Convert.ToString(OrderId));
                content = content.Replace("<orderdate>", Convert.ToString(OrderData.Rows[0]["CreatedOnUtc"]));
                content = content.Replace("<ordertotal>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTotal"])));
                content = content.Replace("<ordersubtotalexcltax>", ToCurrencyString(Convert.ToDecimal(OrderData.Rows[0]["OrderSubtotalExclTax"])));
                content = content.Replace("<paymentmethod>", paymentMethod);

                if (content.Contains("<shippingmethod>"))
                {
                    if (Convert.ToString(OrderData.Rows[0]["ShippingMethod"]).Equals(string.Empty))
                    {
                        content = content.Replace("<shippingmethod>", ConfigurationManager.AppSettings["ShippingMethodsStatus"].ToString());
                        content = content.Replace("<susername>", ConfigurationManager.AppSettings["ShippingMethodsStatus"].ToString());
                        content = content.Replace("<SPhone>", "");
                        content = content.Replace("<SFax>", "");
                    }
                    else
                    {
                        content = content.Replace("<shippingmethod>", Convert.ToString(OrderData.Rows[0]["ShippingMethod"]));

                        //Shipping Address
                        if (!shippingAddressId.Equals(0))
                        {
                            content = content.Replace("<susername>", Convert.ToString(shippingAddress.Rows[0]["UserName"]));
                            content = content.Replace("<SPhone>", "Phone: " + Convert.ToString(shippingAddress.Rows[0]["PhoneNumber"]));
                            content = content.Replace("<SFax>", "Fax  :" + Convert.ToString(shippingAddress.Rows[0]["FaxNumber"]));

                            if (content.Contains("<SCompany>"))
                            {
                                if (Convert.ToString(shippingAddress.Rows[0]["Company"]).Equals(string.Empty))
                                {
                                    content = content.Replace("<SCompany><br />", cr.ToString());
                                }
                                else
                                {
                                    content = content.Replace("<SCompany>", Convert.ToString(shippingAddress.Rows[0]["Company"]));
                                }
                            }
                            content = content.Replace("<SAddress1>", Convert.ToString(shippingAddress.Rows[0]["Address1"]));
                            if (content.Contains("SAddress2"))
                            {
                                if (Convert.ToString(shippingAddress.Rows[0]["Address2"]).Equals(string.Empty))
                                {
                                    content = content.Replace("<SAddress2><br />", cr.ToString());
                                }
                                else
                                {
                                    content = content.Replace("<SAddress2>", Convert.ToString(shippingAddress.Rows[0]["Address2"]));

                                }
                            }

                            if (Convert.ToString(shippingAddress.Rows[0]["City"]) != string.Empty)
                            {
                                content = content.Replace("<SCity,>", Convert.ToString(shippingAddress.Rows[0]["City"]) + ",");

                            }
                            if (Convert.ToString(shippingAddress.Rows[0]["State"]) != string.Empty)
                            {
                                content = content.Replace("<SState,>", Convert.ToString(shippingAddress.Rows[0]["State"]) + ",");

                            }
                            content = content.Replace("<SZip>", Convert.ToString(shippingAddress.Rows[0]["ZipPostalCode"]));

                            content = content.Replace("<SCountry>", Convert.ToString(shippingAddress.Rows[0]["Country"]));

                        }
                    }
                }

                content = content.Replace("<ordershippingexcltax>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderShippingExclTax"])));
                content = content.Replace("<ordertax>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTax"])));
                content = content.Replace("<ordertotal>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTotal"])));
                content = content.Replace("<orderdiscount>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderDiscount"])));

                //mailBody.Append(line);

                if (content.Contains("<productTable>"))
                {
                    string productLine = string.Empty;

                    int productCount = 0;

                    DataTable prodtable = new DataTable("ProductsTable");
                    DataColumn name = new DataColumn("Name", typeof(System.String));
                    DataColumn price = new DataColumn("Unit Price", typeof(System.String));
                    //ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTax"])));
                    DataColumn quantity = new DataColumn("Quantity", typeof(System.String));
                    DataColumn total = new DataColumn("Total Price", typeof(System.String));
                    //ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTax"])));
                    prodtable.Columns.Add(name);
                    prodtable.Columns.Add(price);
                    prodtable.Columns.Add(quantity);
                    prodtable.Columns.Add(total);

                    while (productDetails.Rows.Count > productCount)
                    {
                        string strProductName = productDetails.Rows[productCount]["Product"].ToString();
                        string strPrice = productDetails.Rows[productCount]["UnitPriceExclTax"].ToString();
                        string strDiscountedPrice = productDetails.Rows[productCount]["Discounted Price"].ToString();
                        if (Convert.ToString(strProductName) == string.Empty || Convert.ToString(strProductName) == Convert.ToString(""))
                        {
                            strProductName = productDetails.Rows[productCount]["CPProduct"].ToString();
                        }
                        else
                        {
                            strProductName = productDetails.Rows[productCount]["Product"].ToString();
                        }
                        //if (!String.IsNullOrEmpty(strDiscountedPrice) && Convert.ToString(strDiscountedPrice) != "")
                        //{
                        //    strPrice = Convert.ToString(strDiscountedPrice);
                        //}
                        //else if (Convert.ToString(strPrice) == string.Empty || Convert.ToString(strPrice) == Convert.ToString("0.0000"))
                        //{
                        //    strPrice = productDetails.Rows[productCount]["UnitPriceInclTax"].ToString();
                        //}
                        //else
                        //{
                        //    strPrice = productDetails.Rows[productCount]["Price"].ToString();
                        //}

                        DataRow row = prodtable.NewRow();
                        row.ItemArray = new object[] {
                                                            strProductName.ToString(),
                                                            strPrice.ToString(),
                                                            productDetails.Rows[productCount]["Quantity"].ToString(),
                                                            productDetails.Rows[productCount]["UnitPriceExclTax"].ToString()
                                                           // productDetails.Rows[productCount]["PriceInclTax"].ToString() 
                                                         };
                        prodtable.Rows.Add(row);
                        productCount++;
                    }
                    StringBuilder sb = new StringBuilder();
                    string tab = "\t";
                    sb.AppendLine("<table frame='box'class='MsoNormalTable'border='0'cellspacing='0'cellpadding='0'width='100%'style='width: 100%;border-collapse: collapse; font-size: 10.0pt; font-family: Verdana,sans-serif'>");

                    // Header of dynamic table.
                    sb.Append(tab + "<tr style='background: #EEEEEE; text-align:center; border: 1px solid black; font-size: 10.0pt; font-family: Verdana,sans-serif'>");

                    // Columns in dynamic table
                    foreach (DataColumn dc in prodtable.Columns)
                    {
                        sb.AppendFormat("<td style='border: 1px solid black; font-size: 10.0pt; font-family: Verdana,sans-serif'><strong>{0}</strong></td>", dc.ColumnName);
                    }
                    sb.AppendLine("</tr>");

                    // Total data rows in table.
                    foreach (DataRow dr in prodtable.Rows)
                    {
                        sb.Append(tab + "<tr style='border: 1px solid black; font-size: 10.0pt; font-family: Verdana,sans-serif'>");
                        //Actual Data in columns
                        foreach (DataColumn dc in prodtable.Columns)
                        {
                            int rowIndex = prodtable.Columns.IndexOf(dc);
                            string alignment = Alignment(rowIndex);

                            if (rowIndex.Equals(1) || rowIndex.Equals(3))
                            {
                                string cellValue = dr[dc] != null ? dr[dc].ToString() : "";
                                sb.AppendFormat("<td style='border: 1px solid black; font-size: 10.0pt; font-family: Verdana,sans-serif; text-align:"
                                                + alignment + ";'>{0}</td>", ToCurrencyString(Convert.ToDecimal(cellValue)));
                            }
                            else
                            {
                                string cellValue = dr[dc] != null ? dr[dc].ToString() : "";
                                sb.AppendFormat("<td style='border: 1px solid black; font-size: 10.0pt; font-family: Verdana,sans-serif; text-align:"
                                                + alignment + ";'>{0}</td>", cellValue);
                            }
                        }
                        sb.AppendLine("</tr>");
                    }
                    sb.AppendLine(tab + "</table>");
                    //mailBody.Append(sb);
                    content = content.Replace("<productTable>", sb.ToString());
                }


                #region ProductTable for URL
                if (content.Contains("<productTableURL>"))
                {
                    string URLLine = string.Empty;
                    string jsFileParentPath = string.Empty;
                    int urlProductCount = 0;


                    jsFileParentPath = String.Empty;
                    string strJsExtention = String.Empty;
                    string strVideoProductType = String.Empty;
                    string strInsertVideoURL = String.Empty;

                    while (productUrl.Rows.Count > urlProductCount)
                    {
                        //string strProductName = productUrl.Rows[urlProductCount]["Product"].ToString();
                        string strMask_VideoURL = String.Empty;
                        strJsExtention = productUrl.Rows[urlProductCount]["FileExtension"].ToString();
                        strVideoProductType = productUrl.Rows[urlProductCount]["ProductType"].ToString();
                        DateTime createdDate = Convert.ToDateTime(productUrl.Rows[urlProductCount]["CreatedOnUTC"].ToString());
                        string zipDate = createdDate.ToString("yyMMdd");
                        string strUserName = System.Text.RegularExpressions.Regex.Replace(
                                                Convert.ToString(productUrl.Rows[urlProductCount]["UserName"]),
                                                @"[*?:\|<>\\/]", "").Replace("\"", "");
                        string strFilteredUserName = strUserName.Replace(" ", "");
                        String zipFileName = strFilteredUserName + "_" + zipDate + "_" + OrderId;
                        //If file extension is JS or VarinatName is Playlist and Transcript exist for  product
                        if (strJsExtention == appJSFileExtension || strVideoProductType == appPlaylist || strVideoProductType == strMicrosite)
                        {   //Unique Identifier
                            Guid guid = Guid.NewGuid();
                            string strGuid = guid.ToString();

                            intProductId = Convert.ToInt32(productUrl.Rows[urlProductCount]["PRDID"].ToString());
                            //Get .jas file parent folder path
                            jsFileParentPath = Convert.ToString(ConfigurationManager.AppSettings["parentFileLocation"])
                                              + productUrl.Rows[urlProductCount]["Slug"].ToString();
                            //Get .js File path
                            if (strVideoProductType != strMicrosite)
                            {
                                lstfile = Directory.GetFiles(jsFileParentPath, "*" + productUrl.Rows[urlProductCount]["FileExtension"].ToString());

                                //Get .js File name
                                strfileName = System.IO.Path.GetFileName(lstfile[0]);
                                //create URL to be inserted in wm_VideoSubscription
                                strInsertVideoURL = ConfigurationManager.AppSettings["FileLocation"] + productUrl.Rows[urlProductCount]["Slug"].ToString()
                                      + "/" + strfileName;
                                //unique URL to be sent with Email
                                _VideoURL = ConfigurationManager.AppSettings["SubscriptionURL"] + strGuid;
                            }
                            else
                            {

                                lstfile = Directory.GetFiles(jsFileParentPath, "*" + msFileExtention1);
                                if (lstfile.Length == 0)
                                {
                                    lstfile = Directory.GetFiles(jsFileParentPath, "*" + msFileExtention2);
                                }
                                //Get .js File name
                                strfileName = System.IO.Path.GetFileName(lstfile[0]);
                                //create URL to be inserted in wm_VideoSubscription
                                if (strVideoProductType != strMicrosite)
                                {
                                    strInsertVideoURL = ConfigurationManager.AppSettings["FileLocation"] + productUrl.Rows[urlProductCount]["Slug"].ToString()
                                          + "/" + strfileName;
                                }
                                else
                                {

                                    strInsertVideoURL = "~" + purchasedMicrositePath + zipFileName + "/" + productUrl.Rows[urlProductCount]["Slug"].ToString()
                                      + "/" + strfileName;
                                }//unique URL to be sent with Email

                            }
                            if (strVideoProductType == ConfigurationManager.AppSettings["VideoType"])
                            {
                                _VideoURL = ConfigurationManager.AppSettings["SubscriptionURL"] + strGuid;
                                //Mask URL with product text + Product Name
                                strMask_VideoURL = "<a href=" + _VideoURL + " style='width: 400px' >Please click here to play your subscribed video '" + productUrl.Rows[urlProductCount]["Product"].ToString() + "'.</a>";
                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                                    "Email Video link = " + strMask_VideoURL, null, strPurchaseModuleName);
                            }
                            else if (strVideoProductType == ConfigurationManager.AppSettings["Playlist"])
                            {
                                _VideoURL = ConfigurationManager.AppSettings["SubscriptionURL"] + strGuid;
                                //Mask URL for Playlist
                                strMask_VideoURL = "<a href=" + _VideoURL + " style='width: 400px' >Please click here to play your subscribed Playlist '" + productUrl.Rows[urlProductCount]["Product"].ToString() + "'.</a>";

                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                                    "Email Video link = " + strMask_VideoURL, null, strPurchaseModuleName);
                            }
                            else
                            {
                                _VideoURL = ConfigurationManager.AppSettings["MicrositeSubscriptionURL"] + strGuid;

                                //Mask URL for Microsite
                                strMask_VideoURL = "<a href=" + _VideoURL + " style='width: 400px' >Please click here to visit your subscribed Microsite '" + productUrl.Rows[urlProductCount]["Product"].ToString() + "'.</a>";

                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                                    "Email Video link = " + strMask_VideoURL, null, strPurchaseModuleName);
                            }
                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                         "ProductId:" + intProductId + ",JsExtention:" + strJsExtention + ",jsFileParentPath" + jsFileParentPath + ",URLOrderItemGuid:" +
                            ",fileName:" + strfileName + ",VideoURL:" + strInsertVideoURL + ",VideoURL:" + _VideoURL, null, strPurchaseModuleName);

                            try
                            {
                                //Insert Video Details in wm_VideoSubscription Table
                                InsertVideoSubscription(customerId, OrderId, intProductId, strInsertVideoURL, strGuid);
                            }
                            catch (Exception ex)
                            {
                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.SendEmail", ex.Message, ex, strPurchaseModuleName + "Cannot Insert record Into wm_VideoSubscription");
                            }
                        }//end video product

                        //Check if product is Course Pack
                        else if (productUrl.Rows[urlProductCount]["ProductVariant"].ToString() == ConfigurationManager.AppSettings["CoursePack"])
                        {
                            int intCPProductId = Convert.ToInt32(productUrl.Rows[urlProductCount]["ProductVariantId"].ToString());
                            DataTable VideoInCoursePack = GetOrderDetailsCP(OrderId, intCPProductId);
                            int CPurlProductCount = 0;
                            while (VideoInCoursePack.Rows.Count > CPurlProductCount)
                            {
                                string strCPMask_VideoURL = String.Empty;
                                string strCPIncludeProdProductType = VideoInCoursePack.Rows[CPurlProductCount]["ProductType"].ToString();
                                //Check if Course Pack contains Video Product
                                if (VideoInCoursePack.Rows[CPurlProductCount]["FileExtension"].ToString() == appJSFileExtension)
                                {
                                    //Unique Identifier
                                    Guid guid = Guid.NewGuid();
                                    string strGuid = guid.ToString();
                                    intProductId = Convert.ToInt32(VideoInCoursePack.Rows[CPurlProductCount]["id"].ToString());
                                    strJsExtention = VideoInCoursePack.Rows[CPurlProductCount]["FileExtension"].ToString();
                                    //Get .jas file parent folder path
                                    jsFileParentPath = Convert.ToString(ConfigurationManager.AppSettings["parentFileLocation"])
                                                      + VideoInCoursePack.Rows[CPurlProductCount]["Slug"].ToString();
                                    string cpDetailsProdType = VideoInCoursePack.Rows[CPurlProductCount]["ProductType"].ToString();
                                    if (cpDetailsProdType == strMicrosite)
                                    {
                                        lstfile = Directory.GetFiles(jsFileParentPath, "*" + msFileExtention1);
                                        if (lstfile.Length == 0)
                                        {
                                            lstfile = Directory.GetFiles(jsFileParentPath, "*" + msFileExtention2);
                                        }
                                        strfileName = System.IO.Path.GetFileName(lstfile[0]);
                                        //DateTime createdDate = Convert.ToDateTime(VideoInCoursePack.Rows[CPurlProductCount]["CreatedOnUTC"].ToString());
                                        //string zipDate = createdDate.ToString("yyMMdd");
                                        //String zipFileName = VideoInCoursePack.Rows[CPurlProductCount]["FirstName"].ToString() + VideoInCoursePack.Rows[CPurlProductCount]["LastName"].ToString() + "_" + zipDate + "_" + OrderId;
                                        strInsertVideoURL = "~" + purchasedMicrositePath + zipFileName + "/" + VideoInCoursePack.Rows[CPurlProductCount]["Slug"].ToString()
                                          + "/" + strfileName;
                                    }
                                    else
                                    {
                                        //Get .js File path
                                        lstfile = Directory.GetFiles(jsFileParentPath, "*" + VideoInCoursePack.Rows[CPurlProductCount]["FileExtension"].ToString());
                                        //Get .js File name
                                        strfileName = System.IO.Path.GetFileName(lstfile[0]);
                                        //create URL to be inserted in wm_VideoSubscription
                                        strInsertVideoURL = ConfigurationManager.AppSettings["FileLocation"] + VideoInCoursePack.Rows[CPurlProductCount]["Slug"].ToString()
                                              + "/" + strfileName;
                                        //unique URL to be sent with Email

                                    }
                                    if (strCPIncludeProdProductType == appPlaylist)
                                    {
                                        _VideoURL = ConfigurationManager.AppSettings["SubscriptionURL"] + strGuid;
                                        //Mask URL with product text + Product Name
                                        strCPMask_VideoURL = "<a href=" + _VideoURL + " style='width: 400px' >Please click here to play your subscribed video in Course Pack '" + VideoInCoursePack.Rows[CPurlProductCount]["Product"].ToString() + "'.</a>";
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                                            "Email Video link = " + strCPMask_VideoURL, null, strPurchaseModuleName);
                                    }
                                    else if (strCPIncludeProdProductType == strVideoType)
                                    {
                                        _VideoURL = ConfigurationManager.AppSettings["SubscriptionURL"] + strGuid;
                                        //Mask URL for Playlist
                                        strCPMask_VideoURL = "<a href=" + _VideoURL + " style='width: 400px' >Please click here to play your subscribed Playlist in Course Pack '" + VideoInCoursePack.Rows[CPurlProductCount]["Product"].ToString() + "'.</a>";

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                                            "Email Video link = " + strCPMask_VideoURL, null, strPurchaseModuleName);
                                    }
                                    else if (strCPIncludeProdProductType == strMicrosite)
                                    {
                                        _VideoURL = ConfigurationManager.AppSettings["MicrositeSubscriptionURL"] + strGuid;

                                        //Mask URL for Microsite
                                        //commented by akash Bug#3096
                                        // strMask_VideoURL = "<a href=" + _VideoURL + " style='width: 400px' >Please click here to visit your subscribed Microsite included in Course Pack '" + VideoInCoursePack.Rows[CPurlProductCount]["Product"].ToString() + "'.</a>";
                                        strCPMask_VideoURL = "<a href=" + _VideoURL + " style='width: 400px' >Please click here to visit your subscribed Microsite included in Course Pack '" + VideoInCoursePack.Rows[CPurlProductCount]["Product"].ToString() + "'.</a>";

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                                            "Email Video link = " + strMask_VideoURL, null, strPurchaseModuleName);

                                    }
                                    //Mask URL with product text + Product Name


                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                                 "ProductId:" + intProductId + ",JsExtention:" + strJsExtention + ",jsFileParentPath" + jsFileParentPath + ",URLOrderItemGuid:" +
                                    ",fileName:" + strfileName + ",VideoURL:" + strInsertVideoURL + ",VideoURL:" + _VideoURL, null, strPurchaseModuleName);
                                    try
                                    {
                                        //Insert Video Details in wm_VideoSubscription Table
                                        InsertVideoSubscription(customerId, OrderId, intProductId, strInsertVideoURL, strGuid);
                                    }
                                    catch (Exception ex)
                                    {
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.SendEmail", "Cannot Insert record into wm_VideoSubscription." + ex.Message, ex, strPurchaseModuleName);
                                        return false;
                                    }

                                }
                                CPurlProductCount++;
                                if (!string.IsNullOrEmpty(strCPMask_VideoURL))
                                {

                                    DataRow blankCPRow = table.NewRow();
                                    table.Rows.Add(blankCPRow);
                                    DataRow cpRow = table.NewRow();
                                    cpRow.ItemArray = new object[] { strCPMask_VideoURL };
                                    table.Rows.Add(cpRow);
                                }
                                // urlProductCount++;
                            }
                        }
                        if (!string.IsNullOrEmpty(strMask_VideoURL))
                        {
                            DataRow blankVideoRow = table.NewRow();
                            table.Rows.Add(blankVideoRow);
                            DataRow row = table.NewRow();
                            row.ItemArray = new object[] { strMask_VideoURL };
                            table.Rows.Add(row);
                        }
                        urlProductCount++;
                    }

                    StringBuilder sb = new StringBuilder();
                    string tab = "\t";
                    sb.AppendLine("<table frame='box'class='MsoNormalTable'border='0'cellspacing='0'cellpadding='0'width='100%'style='width: 100.0%;border-collapse: collapse; font-size: 10.0pt; font-family: Verdana,sans-serif'>");

                    // Header of dynamic table.
                    sb.Append(tab + "<tr style=' text-align:center; font-size: 10.0pt; font-family: Verdana,sans-serif'>");

                    // Columns in dynamic table.
                    foreach (DataColumn dc in table.Columns)
                    {
                        sb.AppendFormat("<td style='font-size: 10.0pt; font-family: Verdana,sans-serif'><strong>{0}</strong></td>", dc.ColumnName);
                    }
                    sb.AppendLine("</tr>");

                    // Total data rows in table.
                    foreach (DataRow dr in table.Rows)
                    {
                        sb.Append(tab + "<tr style='font-size: 10.0pt; font-family: Verdana,sans-serif'>");
                        //Actual Data in columns
                        foreach (DataColumn dc in table.Columns)
                        {
                            int rowIndex = table.Columns.IndexOf(dc);
                            string alignment = Alignment(rowIndex);

                            if (rowIndex.Equals(1) || rowIndex.Equals(3))
                            {
                                string cellValue = dr[dc] != null ? dr[dc].ToString() : "";
                                sb.AppendFormat("<td style=' font-size: 10.0pt; font-family: Verdana,sans-serif; text-align:"
                                                + alignment + ";'>{0}</td>", ToCurrencyString(Convert.ToDecimal(cellValue)));
                            }
                            else
                            {
                                string cellValue = dr[dc] != null ? dr[dc].ToString() : "";
                                sb.AppendFormat("<td style=' font-size: 10.0pt; font-family: Verdana,sans-serif; text-align:"
                                                + alignment + ";'>{0}</td>", cellValue);
                            }
                        }
                        sb.AppendLine("</tr>");
                    }
                    sb.AppendLine(tab + "</table>");
                    //mailBody.Append(sb);
                    content = content.Replace("<productTableURL>", sb.ToString());

                }
                #endregion ProductTable for URL

                #endregion

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                            "Completed the reading of each line using stream reader of HTML body in SendMailToCustomerProductsPurchased() method.",
                            null, strPurchaseModuleName);
                //}
                //mail.Body = Convert.ToString(mailBody);
                mail.Body = Convert.ToString(content);
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                            "Mail body formed and values assigned to respective parameters successfuly.", null, strPurchaseModuleName);
            }

            #region Send email

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            SmtpServer.Port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]);
            SmtpServer.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["NetworkCredentialMailID"],
                       ConfigurationManager.AppSettings["NetworkCredentialMailIDPassword"]);
            SmtpServer.EnableSsl = true;

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                        "SendMailToCustomerProductsPurchased() method user name and password validated. Sending mail to: "
                        + customerInfo.Rows[0]["Email"].ToString() + ". CustomerId: " + customerId, null, strPurchaseModuleName);
            try
            {
                SmtpServer.Send(mail);
            }
            catch (SmtpException ex)
            {
                if (Convert.ToString(ex.StatusCode) == "MailboxUnavailable")
                {
                    goto cleanup;
                }

            }


            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.SendMailToCustomerProductsPurchased",
                        "Mail successfully sent to customer email: " + customerInfo.Rows[0]["Email"].ToString()
                        + ". CustomerId: " + customerId, null, strPurchaseModuleName);

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.SendMailToCustomerProductsPurchased",
                        "Order Id: " + OrderId + " placed successfully.", null, strPurchaseModuleName);
            #endregion
            //function call to send Invoice Details to Admin key=OrderReceiptMailTo
            if (ConfigurationManager.AppSettings["StoreFrontInstance"] == "INDIASTOREFRONT")
            {
                SendInvoiceDetailsOfCustomerToAdmin(OrderId);
            }
        cleanup:
            return true;
        }
        catch (Exception ex)
        {

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.SendMailToCustomerProductsPurchased",
                        "SendMailToCustomerProductsPurchased() method. Error: " + ex.Message, null, strPurchaseModuleName);
            return false;
        }
    }
    public bool SendOrderStuckEmail(int OrderId, string ParentDirectoryPath, Hashtable countHashTable)
    {

        string cardName = string.Empty;
        string cardType = string.Empty;
        string creditCardMasked = string.Empty;
        string cardExpirationMonth = string.Empty;
        string cardExpirationYear = string.Empty;
        string[] lstfile;
        List<string> strExten = new List<string>();
        List<string> strJSUrl = new List<string>();
        string strExtention = string.Empty;
        string strExtent;
        string strJsFilePath = string.Empty;
        //string strJsFileExt = ".js";
        string strJsFileExtention = string.Empty;
        string strfileName = string.Empty;
        string strSupportingFileName = string.Empty;
        StringBuilder mailBody = new StringBuilder();
        string emailTemplate = ConfigurationManager.AppSettings["OrderStuckEmailBody"];
        string strZipFileName = string.Empty;
        string strJsFileName = string.Empty;
        string strDownloadProduct = string.Empty;
        String line = String.Empty;
        int intProductId;
        //video          
        string _VideoURL = string.Empty;

        try
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased", "Sending mail for the OrderNumber:"
                        + OrderId, null, strPurchaseModuleName);

            int shippingAddressId = 0;
            DataTable OrderData = GetOrderProductData(OrderId);
            DataTable dtOrderDetailsCP = GetOrderDetails(OrderId);
            int customerId = Convert.ToInt32(OrderData.Rows[0]["CustomerId"]);
            int billingAddressId = Convert.ToInt32(OrderData.Rows[0]["BillingAddressId"]);
            string customerCurrencyCode = Convert.ToString(OrderData.Rows[0]["CustomerCurrencyCode"]);

            if (!DBNull.Value.Equals((OrderData.Rows[0]["ShippingAddressId"])))
            {
                shippingAddressId = Convert.ToInt32(OrderData.Rows[0]["ShippingAddressId"]);
            }
            int CustomerId = Convert.ToInt32(OrderData.Rows[0]["CustomerId"]);

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                        "Getting email id from customer table for CustomerId: " + customerId, null, strPurchaseModuleName);
            DataTable customerInfo = GetCustomerEmail(customerId);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                        "Getting billing address from Address table", null, strPurchaseModuleName);
            DataTable billingAddress = GetCustomerAddress(billingAddressId);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                        "Getting shipping address from Address table", null, strPurchaseModuleName);
            DataTable shippingAddress = GetCustomerAddress(shippingAddressId);

            DataTable productDetails = GetProductDetails(OrderId);
            DataTable productUrl = GetOrderDetails(OrderId);
            //DataTable VideoCoursePack = FilOrderDeatilsCP(OrderId, intCPId);



            if (Convert.ToString(OrderData.Rows[0]["CardName"]) != string.Empty)
            {
                cardName = DecryptMaskedCard(Convert.ToString(OrderData.Rows[0]["CardName"]));
                cardType = DecryptMaskedCard(Convert.ToString(OrderData.Rows[0]["CardType"]));
                creditCardMasked = DecryptMaskedCard(Convert.ToString(OrderData.Rows[0]["MaskedCreditCardNumber"]));
                cardExpirationMonth = DecryptMaskedCard(Convert.ToString(OrderData.Rows[0]["CardExpirationMonth"]));
                cardExpirationYear = DecryptMaskedCard(Convert.ToString(OrderData.Rows[0]["CardExpirationYear"]));
            }//string CardType = encryptionService.DecryptText(cipherText, string.Empty);, [Order].CardExpirationYear
            else
            {
                if (Convert.ToString(OrderData.Rows[0]["MaskedCreditCardNumber"]) == "")
                {
                    creditCardMasked = ConfigurationManager.AppSettings["ShippingMethodsStatus"].ToString();
                }
                else
                {
                    creditCardMasked = DecryptMaskedCard(Convert.ToString(OrderData.Rows[0]["MaskedCreditCardNumber"]));
                }

            }

            List<string> ProductVariantId = new List<string>();
            List<string> ProductName = new List<string>();

            string paymentMethodFromDB = Convert.ToString(OrderData.Rows[0]["PaymentMethodSystemName"]);
            string paymentMethod = string.Empty;
            if (paymentMethodFromDB != "")
            {
                paymentMethod = paymentMethodFromDB.Remove(0, 9);
            }
            else
            {
                paymentMethod = ConfigurationManager.AppSettings["ShippingMethodsStatus"].ToString();
            }

            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient(ConfigurationManager.AppSettings["SmtpHost"]);
            mail.From = new MailAddress(ConfigurationManager.AppSettings["MailFrom"]);
            mail.To.Add(ConfigurationManager.AppSettings["SendProductNotfoundMail"].ToString());
            mail.Subject = ConfigurationManager.AppSettings["OrderstuckEmailSubject"] + ' ' + '#' + OrderId;
            mail.IsBodyHtml = true;

            ////ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased", 
            //"In try catch block of SendMailToCustomerProductsPurchased() method assigned network credentials.", null, strPurchaseModuleName);

            //strExtent = "*" + Convert.ToString(strFileExtension).ToLower();         

            if (File.Exists(emailTemplate))
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                            "HTML Format of email body exists on its location.", null, strPurchaseModuleName);

                string emailFormatFileLocation = ConfigurationManager.AppSettings["OrderStuckEmailBody"].ToString();
                StreamReader reader = new StreamReader(emailFormatFileLocation);
                string content = reader.ReadToEnd();
                reader.Close();


                char cr = (char)13;

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                            "Reading each line using stream reader of HTML body in SendMailToCustomerProductsPurchased() method.", null, strPurchaseModuleName);

                #region MailBody Append


                //lstfile = String.Empty;//strFileExtension));
                DataTable table = new DataTable("ProductsTableURL");
                DataColumn URL = new DataColumn("URL", typeof(System.String));
                //ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTax"])));
                table.Columns.Add(" ");

                //FROM CONFIGURATION FILE
                content = content.Replace("<homepageurl>", Convert.ToString(ConfigurationManager.AppSettings["HomePageUrl"]));
                content = content.Replace("<YourAccountOrders>", Convert.ToString(ConfigurationManager.AppSettings["YourAccountOrders"]));
                content = content.Replace("<youraccountordersname>", Convert.ToString(ConfigurationManager.AppSettings["YourAccountOrdersName"]));

                if (strDownloadProduct != "" || (int)countHashTable["extentCount"] > 0)
                {
                    strDownloadProduct = "\"" + strDownloadProduct + "\"";
                    string strReplace = "<a href=" + strDownloadProduct + " style='width: 400px' >Please click here to download PDF or ePub product(s).</a>";
                    //   content = content.Replace("<downloadfile>", strReplace);

                    DataRow downloadRow = table.NewRow();
                    downloadRow.ItemArray = new object[] { strReplace };
                    table.Rows.Add(downloadRow);
                }
                // Payment Method data
                if (Convert.ToString(OrderData.Rows[0]["PurchaseOrderNumber"]) != string.Empty)
                {
                    content = content.Replace("Card Type:", "Purchase Order No.");
                    content = content.Replace("<cardtype>", Convert.ToString(OrderData.Rows[0]["PurchaseOrderNumber"]));
                    //Card Number:Expiry:
                    content = content.Replace("Card Number:", cr.ToString());
                    content = content.Replace("Expiry:", cr.ToString());
                    content = content.Replace("<cardExpirationMonth>", cr.ToString());
                    content = content.Replace("<cardExpirationYear>", cr.ToString());

                }
                else
                {
                    content = content.Replace("Card Type:", "");
                    content = content.Replace("Expiry:", "");
                    content = content.Replace("<creditCardMasked>", creditCardMasked);
                }

                //TEMPORARY FROM SHIPPING ADDRESS TABLE AFTERWARDS CHANGE IT AND ACCESS USERNAME FROM GENERICATTRIBUTE TABLE
                content = content.Replace("<UserName>", Convert.ToString(billingAddress.Rows[0]["UserName"]));
                //Billing Address
                content = content.Replace("<busername>", Convert.ToString(billingAddress.Rows[0]["UserName"]));
                content = content.Replace("<BPhone>", Convert.ToString(billingAddress.Rows[0]["PhoneNumber"]));
                content = content.Replace("<BFax>", Convert.ToString(billingAddress.Rows[0]["FaxNumber"]));

                if (content.Contains("<BCompany>"))
                {
                    if (Convert.ToString(billingAddress.Rows[0]["Company"]).Equals(string.Empty))
                    {
                        content = content.Replace("<BCompany><br />", cr.ToString());
                    }
                    else
                    {
                        content = content.Replace("<BCompany>", Convert.ToString(billingAddress.Rows[0]["Company"]));
                    }
                }
                content = content.Replace("<BAddress1>", Convert.ToString(billingAddress.Rows[0]["Address1"]));

                if (content.Contains("BAddress2"))
                {
                    if (Convert.ToString(billingAddress.Rows[0]["Address2"]).Equals(string.Empty))
                    {
                        content = content.Replace("<BAddress2><br />", cr.ToString());
                    }
                    else
                    {
                        content = content.Replace("<BAddress2>", Convert.ToString(billingAddress.Rows[0]["Address2"]));
                    }
                }
                if (Convert.ToString(billingAddress.Rows[0]["City"]) != string.Empty)
                {
                    content = content.Replace("<BCity,>", Convert.ToString(billingAddress.Rows[0]["City"]) + ",");
                }

                if (Convert.ToString(billingAddress.Rows[0]["State"]) != string.Empty)
                {
                    content = content.Replace("<BState,>", Convert.ToString(billingAddress.Rows[0]["State"]) + ",");
                }
                content = content.Replace("<BZip>", Convert.ToString(billingAddress.Rows[0]["ZipPostalCode"]));
                content = content.Replace("<BCountry>", Convert.ToString(billingAddress.Rows[0]["Country"]));
                content = content.Replace("<id>", Convert.ToString(OrderId));
                content = content.Replace("<orderdate>", Convert.ToString(OrderData.Rows[0]["CreatedOnUtc"]));
                content = content.Replace("<ordertotal>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTotal"])));
                content = content.Replace("<ordersubtotalexcltax>", ToCurrencyString(Convert.ToDecimal(OrderData.Rows[0]["OrderSubtotalExclTax"])));
                content = content.Replace("<paymentmethod>", paymentMethod);

                if (content.Contains("<shippingmethod>"))
                {
                    if (Convert.ToString(OrderData.Rows[0]["ShippingMethod"]).Equals(string.Empty))
                    {
                        content = content.Replace("<shippingmethod>", ConfigurationManager.AppSettings["ShippingMethodsStatus"].ToString());
                        content = content.Replace("<susername>", ConfigurationManager.AppSettings["ShippingMethodsStatus"].ToString());
                        content = content.Replace("<SPhone>", "");
                        content = content.Replace("<SFax>", "");
                    }
                    else
                    {
                        content = content.Replace("<shippingmethod>", Convert.ToString(OrderData.Rows[0]["ShippingMethod"]));

                        //Shipping Address
                        if (!shippingAddressId.Equals(0))
                        {
                            content = content.Replace("<susername>", Convert.ToString(shippingAddress.Rows[0]["UserName"]));
                            content = content.Replace("<SPhone>", "Phone: " + Convert.ToString(shippingAddress.Rows[0]["PhoneNumber"]));
                            content = content.Replace("<SFax>", "Fax  :" + Convert.ToString(shippingAddress.Rows[0]["FaxNumber"]));

                            if (content.Contains("<SCompany>"))
                            {
                                if (Convert.ToString(shippingAddress.Rows[0]["Company"]).Equals(string.Empty))
                                {
                                    content = content.Replace("<SCompany><br />", cr.ToString());
                                }
                                else
                                {
                                    content = content.Replace("<SCompany>", Convert.ToString(shippingAddress.Rows[0]["Company"]));
                                }
                            }
                            content = content.Replace("<SAddress1>", Convert.ToString(shippingAddress.Rows[0]["Address1"]));
                            if (content.Contains("SAddress2"))
                            {
                                if (Convert.ToString(shippingAddress.Rows[0]["Address2"]).Equals(string.Empty))
                                {
                                    content = content.Replace("<SAddress2><br />", cr.ToString());
                                }
                                else
                                {
                                    content = content.Replace("<SAddress2>", Convert.ToString(shippingAddress.Rows[0]["Address2"]));

                                }
                            }

                            if (Convert.ToString(shippingAddress.Rows[0]["City"]) != string.Empty)
                            {
                                content = content.Replace("<SCity,>", Convert.ToString(shippingAddress.Rows[0]["City"]) + ",");

                            }
                            if (Convert.ToString(shippingAddress.Rows[0]["State"]) != string.Empty)
                            {
                                content = content.Replace("<SState,>", Convert.ToString(shippingAddress.Rows[0]["State"]) + ",");

                            }
                            content = content.Replace("<SZip>", Convert.ToString(shippingAddress.Rows[0]["ZipPostalCode"]));

                            content = content.Replace("<SCountry>", Convert.ToString(shippingAddress.Rows[0]["Country"]));

                        }
                    }
                }


                content = content.Replace("<ordershippingexcltax>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderShippingExclTax"])));

                content = content.Replace("<ordertax>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTax"])));

                content = content.Replace("<ordertotal>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTotal"])));
                content = content.Replace("<orderdiscount>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderDiscount"])));


                //mailBody.Append(line);

                if (content.Contains("<productTable>"))
                {
                    string productLine = string.Empty;

                    int productCount = 0;

                    DataTable prodtable = new DataTable("ProductsTable");
                    DataColumn name = new DataColumn("Name", typeof(System.String));
                    DataColumn price = new DataColumn("Unit Price", typeof(System.String));
                    //ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTax"])));
                    DataColumn quantity = new DataColumn("Quantity", typeof(System.String));
                    DataColumn total = new DataColumn("Total Price", typeof(System.String));
                    //ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTax"])));
                    prodtable.Columns.Add(name);
                    prodtable.Columns.Add(price);
                    prodtable.Columns.Add(quantity);
                    prodtable.Columns.Add(total);

                    while (productDetails.Rows.Count > productCount)
                    {
                        string strProductName = productDetails.Rows[productCount]["Product"].ToString();
                        string strPrice = productDetails.Rows[productCount]["UnitPriceExclTax"].ToString();
                        string strDiscountedPrice = productDetails.Rows[productCount]["Discounted Price"].ToString();
                        if (Convert.ToString(strProductName) == string.Empty || Convert.ToString(strProductName) == Convert.ToString(""))
                        {
                            strProductName = productDetails.Rows[productCount]["CPProduct"].ToString();
                        }
                        else
                        {
                            strProductName = productDetails.Rows[productCount]["Product"].ToString();
                        }
                        //if (!String.IsNullOrEmpty(strDiscountedPrice) || Convert.ToString(strDiscountedPrice) != Convert.ToString("0.0000"))
                        //{
                        //    strPrice = Convert.ToString(strDiscountedPrice);
                        //}
                        //else if (!String.IsNullOrEmpty(strDiscountedPrice) && Convert.ToString(strDiscountedPrice) != "")
                        //{
                        //    strPrice = productDetails.Rows[productCount]["UnitPriceInclTax"].ToString();
                        //}
                        //else
                        //{
                        //    strPrice = productDetails.Rows[productCount]["Price"].ToString();
                        //}
                        DataRow row = prodtable.NewRow();
                        row.ItemArray = new object[] { strProductName.ToString(),
                                                           strPrice.ToString(),
                                                           productDetails.Rows[productCount]["Quantity"].ToString(),
                                                           productDetails.Rows[productCount]["UnitPriceExclTax"].ToString() };
                        prodtable.Rows.Add(row);
                        productCount++;
                    }
                    StringBuilder sb = new StringBuilder();
                    string tab = "\t";
                    sb.AppendLine("<table frame='box'class='MsoNormalTable'border='0'cellspacing='0'cellpadding='0'width='100%'style='width: 100%;border-collapse: collapse; font-size: 10.0pt; font-family: Verdana,sans-serif'>");

                    // Header of dynamic table.
                    sb.Append(tab + "<tr style='background: #EEEEEE; text-align:center; border: 1px solid black; font-size: 10.0pt; font-family: Verdana,sans-serif'>");

                    // Columns in dynamic table
                    foreach (DataColumn dc in prodtable.Columns)
                    {
                        sb.AppendFormat("<td style='border: 1px solid black; font-size: 10.0pt; font-family: Verdana,sans-serif'><strong>{0}</strong></td>", dc.ColumnName);
                    }
                    sb.AppendLine("</tr>");

                    // Total data rows in table.
                    foreach (DataRow dr in prodtable.Rows)
                    {
                        sb.Append(tab + "<tr style='border: 1px solid black; font-size: 10.0pt; font-family: Verdana,sans-serif'>");
                        //Actual Data in columns
                        foreach (DataColumn dc in prodtable.Columns)
                        {
                            int rowIndex = prodtable.Columns.IndexOf(dc);
                            string alignment = Alignment(rowIndex);

                            if (rowIndex.Equals(1) || rowIndex.Equals(3))
                            {
                                string cellValue = dr[dc] != null ? dr[dc].ToString() : "";
                                sb.AppendFormat("<td style='border: 1px solid black; font-size: 10.0pt; font-family: Verdana,sans-serif; text-align:"
                                                + alignment + ";'>{0}</td>", ToCurrencyString(Convert.ToDecimal(cellValue)));
                            }
                            else
                            {
                                string cellValue = dr[dc] != null ? dr[dc].ToString() : "";
                                sb.AppendFormat("<td style='border: 1px solid black; font-size: 10.0pt; font-family: Verdana,sans-serif; text-align:"
                                                + alignment + ";'>{0}</td>", cellValue);
                            }
                        }
                        sb.AppendLine("</tr>");
                    }
                    sb.AppendLine(tab + "</table>");
                    //mailBody.Append(sb);
                    content = content.Replace("<productTable>", sb.ToString());

                }


                #region ProductTable for URL
                if (content.Contains("<productTableURL>"))
                {
                    string URLLine = string.Empty;
                    string jsFileParentPath = string.Empty;
                    int urlProductCount = 0;


                    jsFileParentPath = String.Empty;
                    string strJsExtention = String.Empty;
                    string strVideoProductType = String.Empty;


                    string strInsertVideoURL = String.Empty;


                    //lstfile = String.Empty;//strFileExtension));
                    //  DataTable table = new DataTable("ProductsTableURL");
                    //  DataColumn URL = new DataColumn("URL", typeof(System.String));
                    //ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTax"])));
                    // table.Columns.Add(" ");

                    while (productUrl.Rows.Count > urlProductCount)
                    {
                        //string strProductName = productUrl.Rows[urlProductCount]["Product"].ToString();
                        string strMask_VideoURL = String.Empty;
                        strJsExtention = productUrl.Rows[urlProductCount]["FileExtension"].ToString();
                        strVideoProductType = productUrl.Rows[urlProductCount]["ProductType"].ToString();
                        //If file extension is JS or VarinatName is Playlist and Transcript exist for  product
                        if (strJsExtention == appJSFileExtension || strVideoProductType == appPlaylist)
                        {   //Unique Identifier
                            Guid guid = Guid.NewGuid();
                            string strGuid = guid.ToString();

                            intProductId = Convert.ToInt32(productUrl.Rows[urlProductCount]["PRDID"].ToString());
                            //Get .jas file parent folder path
                            jsFileParentPath = Convert.ToString(ConfigurationManager.AppSettings["parentFileLocation"])
                                              + productUrl.Rows[urlProductCount]["Slug"].ToString();
                            //Get .js File path
                            lstfile = Directory.GetFiles(jsFileParentPath, "*" + productUrl.Rows[urlProductCount]["FileExtension"].ToString());
                            //Get .js File name
                            strfileName = System.IO.Path.GetFileName(lstfile[0]);
                            //create URL to be inserted in wm_VideoSubscription
                            strInsertVideoURL = ConfigurationManager.AppSettings["FileLocation"] + productUrl.Rows[urlProductCount]["Slug"].ToString()
                                  + "/" + strfileName;
                            //unique URL to be sent with Email
                            _VideoURL = ConfigurationManager.AppSettings["SubscriptionURL"] + strGuid;

                            if (strVideoProductType != ConfigurationManager.AppSettings["Playlist"])
                            {
                                //Mask URL with product text + Product Name
                                strMask_VideoURL = "<a href=" + _VideoURL + " style='width: 400px' >Please click here to play your subscribed video '" + productUrl.Rows[urlProductCount]["Product"].ToString() + "'.</a>";
                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMail",
                                    "Email Video link = " + strMask_VideoURL, null, strPurchaseModuleName);
                            }
                            else
                            {
                                //Mask URL for Playlist
                                strMask_VideoURL = "<a href=" + _VideoURL + " style='width: 400px' >Please click here to play your subscribed Playlist '" + productUrl.Rows[urlProductCount]["Product"].ToString() + "'.</a>";

                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                                    "Email Video link = " + strMask_VideoURL, null, strPurchaseModuleName);
                            }
                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                         "ProductId:" + intProductId + ",JsExtention:" + strJsExtention + ",jsFileParentPath" + jsFileParentPath + ",URLOrderItemGuid:" +
                            ",fileName:" + strfileName + ",VideoURL:" + strInsertVideoURL + ",VideoURL:" + _VideoURL, null, strPurchaseModuleName);

                            try
                            {
                                //Insert Video Details in wm_VideoSubscription Table
                                InsertVideoSubscription(customerId, OrderId, intProductId, strInsertVideoURL, strGuid);
                            }
                            catch (Exception ex)
                            {
                                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.SendEmail", ex.Message, ex, strPurchaseModuleName + "Cannot Insert record Into wm_VideoSubscription");
                            }
                        }//end video product

                        //Check if product is Course Pack
                        else if (productUrl.Rows[urlProductCount]["ProductVariant"].ToString() == ConfigurationManager.AppSettings["CoursePack"])
                        {
                            int intCPProductId = Convert.ToInt32(productUrl.Rows[urlProductCount]["ProductVariantId"].ToString());
                            DataTable VideoInCoursePack = GetOrderDetailsCP(OrderId, intCPProductId);
                            int CPurlProductCount = 0;
                            while (VideoInCoursePack.Rows.Count > CPurlProductCount)
                            {
                                string strCPMask_VideoURL = String.Empty;
                                //Check if Course Pack contains Video Product
                                if (VideoInCoursePack.Rows[CPurlProductCount]["FileExtension"].ToString() == appJSFileExtension)
                                {
                                    //Unique Identifier
                                    Guid guid = Guid.NewGuid();
                                    string strGuid = guid.ToString();
                                    intProductId = Convert.ToInt32(VideoInCoursePack.Rows[CPurlProductCount]["id"].ToString());
                                    strJsExtention = VideoInCoursePack.Rows[CPurlProductCount]["FileExtension"].ToString();
                                    //Get .jas file parent folder path
                                    jsFileParentPath = Convert.ToString(ConfigurationManager.AppSettings["parentFileLocation"])
                                                      + VideoInCoursePack.Rows[CPurlProductCount]["Slug"].ToString();
                                    //Get .js File path
                                    lstfile = Directory.GetFiles(jsFileParentPath, "*" + VideoInCoursePack.Rows[CPurlProductCount]["FileExtension"].ToString());
                                    //Get .js File name
                                    strfileName = System.IO.Path.GetFileName(lstfile[0]);
                                    //create URL to be inserted in wm_VideoSubscription
                                    strInsertVideoURL = ConfigurationManager.AppSettings["FileLocation"] + VideoInCoursePack.Rows[CPurlProductCount]["Slug"].ToString()
                                          + "/" + strfileName;
                                    //unique URL to be sent with Email
                                    _VideoURL = ConfigurationManager.AppSettings["SubscriptionURL"] + strGuid;
                                    if (strVideoProductType != ConfigurationManager.AppSettings["Playlist"])
                                    {
                                        //Mask URL with product text + Product Name
                                        strCPMask_VideoURL = "<a href=" + _VideoURL + " style='width: 400px' >Please click here to play your subscribed video in Course Pack '" + VideoInCoursePack.Rows[CPurlProductCount]["Product"].ToString() + "'.</a>";
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                                            "Email Video link = " + strCPMask_VideoURL, null, strPurchaseModuleName);
                                    }
                                    else
                                    {
                                        //Mask URL for Playlist
                                        strCPMask_VideoURL = "<a href=" + _VideoURL + " style='width: 400px' >Please click here to play your subscribed Playlist in Course Pack '" + VideoInCoursePack.Rows[CPurlProductCount]["Product"].ToString() + "'.</a>";

                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                                            "Email Video link = " + strCPMask_VideoURL, null, strPurchaseModuleName);
                                    }
                                    //Mask URL with product text + Product Name


                                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                                 "ProductId:" + intProductId + ",JsExtention:" + strJsExtention + ",jsFileParentPath" + jsFileParentPath + ",URLOrderItemGuid:" +
                                    ",fileName:" + strfileName + ",VideoURL:" + strInsertVideoURL + ",VideoURL:" + _VideoURL, null, strPurchaseModuleName);
                                    try
                                    {
                                        //Insert Video Details in wm_VideoSubscription Table
                                        InsertVideoSubscription(customerId, OrderId, intProductId, strInsertVideoURL, strGuid);
                                    }
                                    catch (Exception ex)
                                    {
                                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.SendEmail", "Cannot Insert record into wm_VideoSubscription." + ex.Message, ex, strPurchaseModuleName);
                                        return false;
                                    }

                                }
                                CPurlProductCount++;
                                if (!string.IsNullOrEmpty(strCPMask_VideoURL))
                                {

                                    DataRow blankCPRow = table.NewRow();
                                    table.Rows.Add(blankCPRow);
                                    DataRow cpRow = table.NewRow();
                                    cpRow.ItemArray = new object[] { strCPMask_VideoURL };
                                    table.Rows.Add(cpRow);
                                }
                                // urlProductCount++;
                            }
                        }
                        if (!string.IsNullOrEmpty(strMask_VideoURL))
                        {
                            DataRow blankVideoRow = table.NewRow();
                            table.Rows.Add(blankVideoRow);
                            DataRow row = table.NewRow();
                            row.ItemArray = new object[] { strMask_VideoURL };
                            table.Rows.Add(row);


                        }
                        urlProductCount++;
                    }

                    StringBuilder sb = new StringBuilder();
                    string tab = "\t";
                    sb.AppendLine("<table frame='box'class='MsoNormalTable'border='0'cellspacing='0'cellpadding='0'width='100%'style='width: 100.0%;border-collapse: collapse; font-size: 10.0pt; font-family: Verdana,sans-serif'>");

                    // Header of dynamic table.
                    sb.Append(tab + "<tr style=' text-align:center; font-size: 10.0pt; font-family: Verdana,sans-serif'>");

                    // Columns in dynamic table.
                    foreach (DataColumn dc in table.Columns)
                    {
                        sb.AppendFormat("<td style='font-size: 10.0pt; font-family: Verdana,sans-serif'><strong>{0}</strong></td>", dc.ColumnName);
                    }
                    sb.AppendLine("</tr>");

                    // Total data rows in table.
                    foreach (DataRow dr in table.Rows)
                    {
                        sb.Append(tab + "<tr style='font-size: 10.0pt; font-family: Verdana,sans-serif'>");
                        //Actual Data in columns
                        foreach (DataColumn dc in table.Columns)
                        {
                            int rowIndex = table.Columns.IndexOf(dc);
                            string alignment = Alignment(rowIndex);

                            if (rowIndex.Equals(1) || rowIndex.Equals(3))
                            {
                                string cellValue = dr[dc] != null ? dr[dc].ToString() : "";
                                sb.AppendFormat("<td style=' font-size: 10.0pt; font-family: Verdana,sans-serif; text-align:"
                                                + alignment + ";'>{0}</td>", ToCurrencyString(Convert.ToDecimal(cellValue)));
                            }
                            else
                            {
                                string cellValue = dr[dc] != null ? dr[dc].ToString() : "";
                                sb.AppendFormat("<td style=' font-size: 10.0pt; font-family: Verdana,sans-serif; text-align:"
                                                + alignment + ";'>{0}</td>", cellValue);
                            }
                        }
                        sb.AppendLine("</tr>");
                    }
                    sb.AppendLine(tab + "</table>");
                    //mailBody.Append(sb);
                    content = content.Replace("<productTableURL>", sb.ToString());

                }
                #endregion ProductTable for URL

                #endregion

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                            "Completed the reading of each line using stream reader of HTML body in SendMailToCustomerProductsPurchased() method.",
                            null, strPurchaseModuleName);
                //}
                //mail.Body = Convert.ToString(mailBody);
                mail.Body = Convert.ToString(content);
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                            "Mail body formed and values assigned to respective parameters successfuly.", null, strPurchaseModuleName);
            }

            #region Send email

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            SmtpServer.Port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]);
            SmtpServer.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["NetworkCredentialMailID"],
                       ConfigurationManager.AppSettings["NetworkCredentialMailIDPassword"]);
            SmtpServer.EnableSsl = true;

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased",
                        "SendMailToCustomerProductsPurchased() method user name and password validated. Sending mail to: "
                        + customerInfo.Rows[0]["Email"].ToString() + ". CustomerId: " + customerId, null, strPurchaseModuleName);
            Console.WriteLine("Start order detail mail sending.");
            SmtpServer.Send(mail);
            Console.WriteLine("Order detail mail sent.");

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.SendMailToCustomerProductsPurchased",
                        "Mail successfully sent to customer email: " + customerInfo.Rows[0]["Email"].ToString()
                        + ". CustomerId: " + customerId, null, strPurchaseModuleName);
            Console.WriteLine("Order Id:" + OrderId + " placed successfully.");

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.SendMailToCustomerProductsPurchased",
                        "Order Id: " + OrderId + " placed successfully.", null, strPurchaseModuleName);
            #endregion

            return true;
        }
        catch (Exception ex)
        {

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.SendMailToCustomerProductsPurchased",
                        "SendMailToCustomerProductsPurchased() method. Error: " + ex.Message, null, strPurchaseModuleName);
            return false;
        }
    }

    /// <summary>
    /// 
    /// </summary>
    /// <param name="sessionId"></param>
    /// <param name="facutlyName"></param>
    /// <param name="FacultyEmail"></param>
    /// <param name="simulationUrl"></param>
    /// <param name="simulationName"></param>
    /// <param name="orderNumber"></param>
    /// <param name="classDescription"></param>
    public void SendSimulationMailToFaculty(string sessionId, string facutlyName, string FacultyEmail, string simulationUrl,
                                            string simulationName, string orderNumber, string classDescription)
    {
        classDescription = classDescription.Split(':')[1];
        StringBuilder emailsubject = new StringBuilder();
        DataSet ds = new DataSet();
        SqlCommand cmdd = new SqlCommand("pr_uva_GetSimulationEmailTemplate", thisConnection);
        cmdd.CommandType = CommandType.StoredProcedure;
        cmdd.Parameters.Add("@temaplateName", SqlDbType.VarChar).Value = ConfigurationManager.AppSettings["simulatinEmailTemplateName"];
        cmdd.Parameters.Add("@storename", SqlDbType.VarChar).Value = ConfigurationManager.AppSettings["storenameDB"];
        SqlDataAdapter da = new SqlDataAdapter(cmdd);
        da.Fill(ds);

        MailMessage mail = new MailMessage();
        SmtpClient SmtpServer = new SmtpClient(ConfigurationManager.AppSettings["SmtpHost"]);
        mail.From = new MailAddress(ConfigurationManager.AppSettings["MailFrom"]);
        mail.To.Add(FacultyEmail);
        //emailsubject = dt.Rows[0]["Subject"].ToString();
        if (ds.Tables[0].Rows[0]["Subject"].ToString() != "")
        {
            String line = ds.Tables[0].Rows[0]["Subject"].ToString();
            line = line.Replace(ConfigurationManager.AppSettings["Product.SimulationSessionId"], sessionId);
            line = line.Replace(ConfigurationManager.AppSettings["CustomerName"], facutlyName);
            line = line.Replace(ConfigurationManager.AppSettings["ProductName"], simulationName);
            line = line.Replace(ConfigurationManager.AppSettings["simulationLink"], simulationUrl);
            line = line.Replace(ConfigurationManager.AppSettings["StoreName"], ds.Tables[1].Rows[0]["Value"].ToString());
            line = line.Replace(ConfigurationManager.AppSettings["OrderNumber"], orderNumber);
            line = line.Replace(ConfigurationManager.AppSettings["ClassDescription"], classDescription);
            emailsubject.Append(line.ToString());
        }
        mail.Subject = Convert.ToString(emailsubject);//ConfigurationManager.AppSettings["SimulationEmailSubject"];
        mail.IsBodyHtml = true;

        //ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailToCustomerProductsPurchased", 
        //"In try catch block of SendMailToCustomerProductsPurchased() method assigned network credentials.", null, strPurchaseModuleName);

        StringBuilder mailBody = new StringBuilder();
        string _emailTest = ds.Tables[0].Rows[0]["Body"].ToString();


        if (_emailTest != "")
        {
            String line = _emailTest;
            line = line.Replace(ConfigurationManager.AppSettings["Product.SimulationSessionId"], sessionId);
            line = line.Replace(ConfigurationManager.AppSettings["CustomerName"], facutlyName);
            line = line.Replace(ConfigurationManager.AppSettings["ProductName"], simulationName);
            line = line.Replace(ConfigurationManager.AppSettings["simulationLink"], simulationUrl);
            line = line.Replace(ConfigurationManager.AppSettings["StoreName"], ds.Tables[1].Rows[0]["Value"].ToString());
            line = line.Replace(ConfigurationManager.AppSettings["OrderNumber"], orderNumber);
            line = line.Replace(ConfigurationManager.AppSettings["ClassDescription"], classDescription);
            mailBody.Append(line.ToString());
        }
        mail.Body = Convert.ToString(mailBody);

        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        SmtpServer.Port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"].ToString());
        SmtpServer.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["NetworkCredentialMailID"].ToString(),
                                 ConfigurationManager.AppSettings["NetworkCredentialMailIDPassword"].ToString());
        SmtpServer.EnableSsl = true;

        SmtpServer.Send(mail);
        SmtpServer.Dispose();
        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                    "Email send to faculty, EmailId:" + FacultyEmail + ", simluation link: " + simulationUrl, null, strPurchaseModuleName);
    }

    public void SendEpicenterAuthenticationMailToUser(string FacultyEmail, string simulationUrl, string simulationName,
                                                          string username, string password)
    {
        String line = string.Empty;
        StringBuilder emailsubject = new StringBuilder();
        DataSet ds = new DataSet();
        SqlCommand cmdd = new SqlCommand("pr_uva_GetSimulationEmailTemplate", thisConnection);
        cmdd.CommandType = CommandType.StoredProcedure;
        cmdd.Parameters.Add("@temaplateName", SqlDbType.VarChar).Value = ConfigurationManager.AppSettings["EpicenterSimulatinEmailTemplateName"];
        cmdd.Parameters.Add("@storename", SqlDbType.VarChar).Value = ConfigurationManager.AppSettings["storenameDB"];
        SqlDataAdapter da = new SqlDataAdapter(cmdd);
        da.Fill(ds);

        MailMessage mail = new MailMessage();
        SmtpClient SmtpServer = new SmtpClient(ConfigurationManager.AppSettings["SmtpHost"]);
        mail.From = new MailAddress(ConfigurationManager.AppSettings["MailFrom"]);
        mail.To.Add(FacultyEmail);

        if (ds.Tables[0].Rows[0]["Subject"].ToString() != "")
        {
            line = ds.Tables[0].Rows[0]["Subject"].ToString();
            line = line.Replace(ConfigurationManager.AppSettings["EpicenterProductName"], simulationName);
            line = line.Replace(ConfigurationManager.AppSettings["EpicenterUsername"], username);
            line = line.Replace(ConfigurationManager.AppSettings["EpicenterPassword"], password);
            line = line.Replace(ConfigurationManager.AppSettings["EpicenterSimulationUrl"], simulationUrl);

            emailsubject.Append(line.ToString());
        }
        mail.Subject = Convert.ToString(emailsubject);//ConfigurationManager.AppSettings["SimulationEmailSubject"];
        mail.IsBodyHtml = true;

        StringBuilder mailBody = new StringBuilder();
        string _emailTest = ds.Tables[0].Rows[0]["Body"].ToString();


        if (_emailTest != "")
        {
            line = _emailTest;
            line = line.Replace(ConfigurationManager.AppSettings["EpicenterProductName"], simulationName);
            line = line.Replace(ConfigurationManager.AppSettings["EpicenterUsername"], username);
            line = line.Replace(ConfigurationManager.AppSettings["EpicenterPassword"], password);
            line = line.Replace(ConfigurationManager.AppSettings["EpicenterSimulationUrl"], simulationUrl);
            mailBody.Append(line.ToString());
        }
        mail.Body = Convert.ToString(mailBody);

        ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
        SmtpServer.Port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"].ToString());
        SmtpServer.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["NetworkCredentialMailID"].ToString(),
                                 ConfigurationManager.AppSettings["NetworkCredentialMailIDPassword"].ToString());
        SmtpServer.EnableSsl = true;
        SmtpServer.Send(mail);
        SmtpServer.Dispose();
        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.getConnection",
                    "Email send to faculty, EmailId:" + FacultyEmail + ", simluation link: " + simulationUrl, null, strPurchaseModuleName);
    }

    public void SendMailOrderPlacedReceiptToAdmin(int OrderId)
    {
        try
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                "GetConnection.SendMailOrderPlacedReceiptToAdmin", "Sending mail for the OrderNumber:"
                        + OrderId, null, strPurchaseModuleName);

            int shippingAddressId = 0;
            //string DisplayName = string.Empty;
            DataTable OrderData = GetOrderProductData(OrderId);
            // DataTable dtOrderDataDetails = GetOrderDetails(OrderId);
            string Uname = displayName;

            if (OrderData.Rows.Count > 0)
            {

                int customerId = Convert.ToInt32(OrderData.Rows[0]["CustomerId"]);
                int billingAddressId = Convert.ToInt32(OrderData.Rows[0]["BillingAddressId"]);
                string customerCurrencyCode = Convert.ToString(OrderData.Rows[0]["CustomerCurrencyCode"]);

                if (!DBNull.Value.Equals((OrderData.Rows[0]["ShippingAddressId"])))
                {
                    shippingAddressId = Convert.ToInt32(OrderData.Rows[0]["ShippingAddressId"]);
                }
                int CustomerId = Convert.ToInt32(OrderData.Rows[0]["CustomerId"]);


                DataTable customerInfo = GetCustomerEmail(customerId);

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailOrderPlacedReceiptToAdmin",
                           "Getting email id from customer table for CustomerId: " + customerId, null, strPurchaseModuleName);

                DataTable billingAddress = GetCustomerAddress(billingAddressId);//vijay.nagarsoge@workmethods.com

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailOrderPlacedReceiptToAdmin",
                            "Getting billing address from Address table", null, strPurchaseModuleName);

                DataTable shippingAddress = GetCustomerAddress(shippingAddressId);

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailOrderPlacedReceiptToAdmin",
                            "Getting shipping address from Address table", null, strPurchaseModuleName);

                DataTable productDetails = GetProductDetails(OrderId);
                List<string> ProductVariantId = new List<string>();
                List<string> ProductName = new List<string>();
                string paymentMethodFromDB = Convert.ToString(OrderData.Rows[0]["PaymentMethodSystemName"]);
                string paymentMethod = string.Empty;
                if (paymentMethodFromDB != "")
                {
                    paymentMethod = paymentMethodFromDB.Remove(0, 9);
                }
                else
                {
                    //Not Applicable
                    paymentMethod = ConfigurationManager.AppSettings["ShippingMethodsStatus"].ToString();
                }
                MailMessage mail = new MailMessage();
                SmtpClient SmtpServer = new SmtpClient(ConfigurationManager.AppSettings["SmtpHost"]);
                mail.From = new MailAddress(ConfigurationManager.AppSettings["MailFrom"]);//uvacomm-test@workmethods.com
                mail.To.Add(ConfigurationManager.AppSettings["OrderReceiptMailTo"]);
                mail.Subject = ConfigurationManager.AppSettings["OrderReceiptEmailSubject"] + "#" + OrderId;//Darden Business Publishing Storefront. Purchase Receipt for Order#241504
                mail.IsBodyHtml = true;
                StringBuilder mailBody = new StringBuilder();
                string _emailFormatFileLocation = ConfigurationManager.AppSettings["OrderReceiptEmailBody"];
                //_emailFormatFileLocation=D:\USER_DATA\Vaidehi\StoreFront\trunk\Modules\Sales Completion\SalesCompletion\bin\Debug\IndiaStorefrontOrderReceiptEmailAdminTemplate.html
                if (File.Exists(_emailFormatFileLocation))
                {
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailOrderPlacedReceiptToAdmin",
                                "HTML format of email body exists on its location. " + _emailFormatFileLocation, null, strPurchaseModuleName);
                    //emailFormatFileLocation =D:\USER_DATA\Vaidehi\StoreFront\trunk\Modules\Sales Completion\SalesCompletion\bin\Debug\IndiaStorefrontOrderReceiptEmailAdminTemplate.html
                    string emailFormatFileLocation = ConfigurationManager.AppSettings["OrderReceiptEmailBody"].ToString();
                    StreamReader reader = new StreamReader(emailFormatFileLocation);
                    string content = reader.ReadToEnd();
                    //<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m='http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"> <head> <meta http-equiv="Content-Type" content="text/html; charset=utf-8"> <meta name="Generator" content="Microsoft Word 14 (filtered medium)"> <!--[if !mso]><style>v\:* {behavior:url(#default#VML);}o\:* {behavior:url(#default#VML);}w\:* {behavior:url(#default#VML);}.shape {behavior:url(#default#VML);}</style><![endif]--> <title> Placed An Order From Your Store.</title> <style> <!-- /* Font Definitions */ @font-face { font-family: "Cambria Math"; panose-1: 2 4 5 3 5 4 6 3 2 4; } @font-face { font-family: Calibri; panose-1: 2 15 5 2 2 2 4 3 2 4; } @font-face { font-family: Tahoma; panose-1: 2 11 6 4 3 5 4 4 2 4; } @font-face { font-family: Times New Roman; panose-1: 2 11 6 4 3 5 4 4 2 4; } /* Style Definitions */ p.MsoNormal, li.MsoNormal, div.MsoNormal { margin: 0in; margin-bottom: .0001pt; font-size: 12.0pt; font-family: "Times New Roman"; } a:link, span.MsoHyperlink { mso-style-priority: 99; color: #0066CC; text-decoration: underline; } a:visited, span.MsoHyperlinkFollowed { mso-style-priority: 99; color: purple; text-decoration: underline; } p { mso-style-priority: 99; mso-margin-top-alt: auto; margin-right: 0in; mso-margin-bottom-alt: auto; margin-left: 0in; font-size: 12.0pt; font-family: "Times New Roman"; } span.small { mso-style-name: small; } span.price { mso-style-name: price; } span.EmailStyle20 { mso-style-type: personal-reply; font-family: "Calibri","sans-serif"; color: windowtext; } .MsoChpDefault { mso-style-type: export-only; font-size: 12.0pt; } @page WordSection1 { size: 8.5in 11.0in; margin: 1.0in 1.0in 1.0in 1.0in; } div.WordSection1 { page: WordSection1; } #_x0000_i1025 { height: 43px; } .MsoNormalTable { margin-right: 0px; } .auto-style1 { height: 108px; } .style1 { width: 535px; } .style2 { width: 14px; } .style3 { width: 97px; height: 15px; } .auto-style3 { width: 97px; height: 15px; } .auto-style4 { width: 74px; height: 15px; } .auto-style5 { width: 74px; height: 15px; } --> </style> </head> <body bgcolor="white" lang="EN-US" link=blue vlink=purple> <div align="left" class="WordSection1" > <!-- <table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" style='margin-left: .5in; > <tr> <td style='padding: .75pt .75pt .75pt .75pt'> <div align="center"> <table width="100%"> <tr> <td>--> <p> <a href="<homepageurl>">Darden Business Publishing IndiaStorefront</a><br><br> <span style='font-size: 12.0pt; font-family: "Times New Roman"'><UserName> <a href="mailto:<Email>"> (<Email>) </a> has just placed an order from your store. Below is the summary of the order.</span> </p> <!-- <table width="80%" style='width: 100.0%'> <tr> <td nowrap style='padding: 3.0pt 3.0pt 3.0pt 3.0pt' class="style1">--> <p class="MsoNormal"> <span style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'>Order Number: </span><span style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'> </span> <span style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'> <id> </span> </p> <p class="MsoNormal"> <span style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'>Date Ordered:</span><span style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'> </span> <orderdate> </p> <br /><br /> <div> <p class="MsoNormal" style='margin-bottom: 12.0pt'> <span style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'>Billing Address:</span><span style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'><br> <busername><br /> <BCompany><br /> <BAddress1><br /> <BAddress2><br /> <BCity,> <BState,> <BZip><br /> <BCountry><br /> </span> <BPhone> <BFax> </p> </div> <div> <p class="MsoNormal" style='margin-bottom: 12.0pt'> <span style='font-size: 12.0pt; font-family: "Times New Roman"'>Shipping Address:</span><span style='font-size: 12.0pt; font-family: "Times New Roman"'><br> <susername><br /> <SCompany><br /> <SAddress1><br /> <SAddress2><br /> <SCity,> <SState,> <SZip><br /> <SCountry><br /> <SPhone><br /> <SFax> </span> </p> </div> <div class="shipping-method" style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'> <div > Shipping Method: </div> <div class="shipping-method-value"> <shippingmethod> </div> </div> <!-- </td> </tr> <tr> <td colspan="2" style='padding: 3.0pt 3.0pt 3.0pt 3.0pt;'>--> <productTable> <!--<table border="0" style='width: 68%; font-size: 12.0pt; font-family: "Times New Roman","sans-serif"; margin-left: 0px;text-align:left'>--> <tr style='background:#DDE2E6'> <!-- <td style='background:white'></td>--><td style='background:white'></td> <td colspan="2"class="auto-style3"> <strong>Sub-Total: </strong> </td> <td class="auto-style4" style="text-align:right"> <currency>&nbsp;&nbsp;<ordersubtotalexcltax> </td> </tr> <tr style='background:#DDE2E6'> <!--<td style='background:white'></td>--><td style='background:white'></td> <td colspan="2" class="style3"> <strong>GST(18%):</strong> </td> <td class="auto-style5" style="text-align:right"> <currency>&nbsp;&nbsp;<ordertax> </td> </tr> <tr style='background:#DDE2E6'> <!--<td style='background:white'></td>--><td style='background:white'></td> <td colspan="2" class="style3"> <strong>Order Total:</strong> </td> <td class="auto-style5" style="text-align:right"> <currency>&nbsp;&nbsp;<ordertotal> </td> </tr> </table> <!--</p>--> <!-- </td> </tr> <tr> <td colspan="2" style='padding: 3.0pt 3.0pt 3.0pt 3.0pt' class="auto-style1"> &nbsp;</td> </tr> </table> </td> </tr> </table> </div> </td> </tr> </table>--> </div> </body> </html>
                    reader.Close();
                    String line = String.Empty;
                    char cr = (char)13;
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailOrderPlacedReceiptToAdmin",
                                "Reading each line using stream reader of HTML body in SendMailOrderPlacedReceiptToAdmin() method.", null, strPurchaseModuleName);

                    #region MailBody Append
                    content = content.Replace("<homepageurl>", Convert.ToString(ConfigurationManager.AppSettings["HomePageUrl"]));
                    content = content.Replace("<YourAccountOrders>", Convert.ToString(ConfigurationManager.AppSettings["YourAccountOrders"]));
                    content = content.Replace("<youraccountordersname>", Convert.ToString(ConfigurationManager.AppSettings["YourAccountOrdersName"]));
                    //TEMPORARY FROM SHIPPING ADDRESS TABLE AFTERWARDS CHANGE IT AND ACCESS USERNAME FROM GENERICATTRIBUTE TABLE
                    //  content = content.Replace("<UserName>", Convert.ToString(billingAddress.Rows[0]["UserName"]));
                    content = content.Replace("<UserName>", Uname.ToString());
                    content = content.Replace("<Email>", Convert.ToString(billingAddress.Rows[0]["Expr1"]));
                    //Billing Address
                    content = content.Replace("<busername>", Convert.ToString(billingAddress.Rows[0]["UserName"]));
                    content = content.Replace("<BPhone>", Convert.ToString(billingAddress.Rows[0]["PhoneNumber"]));
                    content = content.Replace("<BFax>", Convert.ToString(billingAddress.Rows[0]["FaxNumber"]));
                    if (content.Contains("<BCompany>"))
                    {
                        if (Convert.ToString(billingAddress.Rows[0]["Company"]).Equals(string.Empty))
                        {
                            content = content.Replace("<BCompany><br />", cr.ToString());
                        }
                        else
                        {
                            content = content.Replace("<BCompany>", Convert.ToString(billingAddress.Rows[0]["Company"]));
                        }
                    }
                    content = content.Replace("<BAddress1>", Convert.ToString(billingAddress.Rows[0]["Address1"]));
                    if (content.Contains("BAddress2"))
                    {
                        if (Convert.ToString(billingAddress.Rows[0]["Address2"]).Equals(string.Empty))
                        {
                            content = content.Replace("<BAddress2><br />", cr.ToString());
                        }
                        else
                        {
                            content = content.Replace("<BAddress2>", Convert.ToString(billingAddress.Rows[0]["Address2"]));
                        }
                    }
                    if (Convert.ToString(billingAddress.Rows[0]["City"]) != string.Empty)
                    {
                        content = content.Replace("<BCity,>", Convert.ToString(billingAddress.Rows[0]["City"]) + ",");

                    }
                    if (Convert.ToString(billingAddress.Rows[0]["State"]) != string.Empty)
                    {
                        content = content.Replace("<BState,>", Convert.ToString(billingAddress.Rows[0]["State"]) + ",");
                    }
                    content = content.Replace("<BZip>", Convert.ToString(billingAddress.Rows[0]["ZipPostalCode"]));
                    content = content.Replace("<BCountry>", Convert.ToString(billingAddress.Rows[0]["Country"]));
                    content = content.Replace("<id>", Convert.ToString(OrderId));
                    content = content.Replace("<orderdate>", Convert.ToString(String.Format("{0:dddd, MMMM d, yyyy}", (OrderData.Rows[0]["CreatedOnUtc"]))));
                    content = content.Replace("<ordertotal>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTotal"])));
                    content = content.Replace("<ordersubtotalexcltax>", ToCurrencyString(Convert.ToDecimal(OrderData.Rows[0]["OrderSubtotalExclTax"])));
                    content = content.Replace("<paymentmethod>", paymentMethod);
                    if (content.Contains("<shippingmethod>"))
                    {
                        if (Convert.ToString(OrderData.Rows[0]["ShippingMethod"]).Equals(string.Empty))
                        {
                            content = content.Replace("<shippingmethod>", ConfigurationManager.AppSettings["ShippingMethodsStatus"].ToString());
                            content = content.Replace("<susername>", ConfigurationManager.AppSettings["ShippingMethodsStatus"].ToString());
                            content = content.Replace("<SPhone>", "");
                            content = content.Replace("<SFax>", "");
                        }
                        else
                        {
                            content = content.Replace("<shippingmethod>", Convert.ToString(OrderData.Rows[0]["ShippingMethod"]));
                            //Shipping Address
                            if (!shippingAddressId.Equals(0))
                            {
                                content = content.Replace("<susername>", Convert.ToString(shippingAddress.Rows[0]["UserName"]));
                                content = content.Replace("<SPhone>", "Phone: " + Convert.ToString(shippingAddress.Rows[0]["PhoneNumber"]));
                                content = content.Replace("<SFax>", "Fax  :" + Convert.ToString(shippingAddress.Rows[0]["FaxNumber"]));

                                if (content.Contains("<SCompany>"))
                                {
                                    if (Convert.ToString(shippingAddress.Rows[0]["Company"]).Equals(string.Empty))
                                    {
                                        content = content.Replace("<SCompany><br />", cr.ToString());
                                    }
                                    else
                                    {
                                        content = content.Replace("<SCompany>", Convert.ToString(shippingAddress.Rows[0]["Company"]));
                                    }
                                }

                                content = content.Replace("<SAddress1>", Convert.ToString(shippingAddress.Rows[0]["Address1"]));

                                if (content.Contains("SAddress2"))
                                {
                                    if (Convert.ToString(shippingAddress.Rows[0]["Address2"]).Equals(string.Empty))
                                    {
                                        content = content.Replace("<SAddress2><br />", cr.ToString());
                                    }
                                    else
                                    {
                                        content = content.Replace("<SAddress2>", Convert.ToString(shippingAddress.Rows[0]["Address2"]));
                                    }
                                }

                                if (Convert.ToString(shippingAddress.Rows[0]["City"]) != string.Empty)
                                {
                                    content = content.Replace("<SCity,>", Convert.ToString(shippingAddress.Rows[0]["City"]) + ",");

                                }
                                if (Convert.ToString(shippingAddress.Rows[0]["State"]) != string.Empty)
                                {
                                    content = content.Replace("<SState,>", Convert.ToString(shippingAddress.Rows[0]["State"]) + ",");

                                }
                                content = content.Replace("<SZip>", Convert.ToString(shippingAddress.Rows[0]["ZipPostalCode"]));
                                content = content.Replace("<SCountry>", Convert.ToString(shippingAddress.Rows[0]["Country"]));

                            }
                        }
                    }
                    content = content.Replace("<ordershippingexcltax>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderShippingExclTax"])));

                    content = content.Replace("<ordertax>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTax"])));

                    content = content.Replace("<ordertotal>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTotal"])));
                    content = content.Replace("<orderdiscount>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderDiscount"])));

                    if (content.Contains("<productTable>"))
                    {
                        string productLine = string.Empty;
                        int productCount = 0;
                        DataTable table = new DataTable("ProductsTable");
                        DataColumn name = new DataColumn("Name", typeof(System.String));
                        DataColumn price = new DataColumn("Unit Price", typeof(System.String));
                        DataColumn quantity = new DataColumn("Quantity", typeof(System.String));
                        DataColumn total = new DataColumn("Total", typeof(System.String));

                        table.Columns.Add(name);
                        table.Columns.Add(price);
                        table.Columns.Add(quantity);
                        table.Columns.Add(total);

                        while (productDetails.Rows.Count > productCount)
                        {
                            string strProductName = productDetails.Rows[productCount]["Product"].ToString();
                            string strPrice = productDetails.Rows[productCount]["UnitPriceExclTax"].ToString();
                            string strDiscountedPrice = productDetails.Rows[productCount]["Discounted Price"].ToString();
                            if (Convert.ToString(strProductName) == string.Empty || Convert.ToString(strProductName) == Convert.ToString(""))
                            {
                                strProductName = productDetails.Rows[productCount]["CPProduct"].ToString();
                            }
                            else
                            {
                                strProductName = productDetails.Rows[productCount]["Product"].ToString();
                            }
                            //if (!String.IsNullOrEmpty(strDiscountedPrice))
                            //{
                            //    strPrice = Convert.ToString(strDiscountedPrice);

                            //}
                            ////else if (!String.IsNullOrEmpty(strDiscountedPrice) && Convert.ToString(strDiscountedPrice) != "")
                            ////{
                            ////    strPrice = productDetails.Rows[productCount]["UnitPriceInclTax"].ToString();
                            ////}
                            //else
                            //{
                            //    strPrice = productDetails.Rows[productCount]["Price"].ToString();
                            //}
                            DataRow row = table.NewRow();
                            row.ItemArray = new object[] {
                                                                strProductName.ToString(),
                                                                strPrice.ToString(),
                                                                productDetails.Rows[productCount]["Quantity"].ToString(),
                                                                productDetails.Rows[productCount]["UnitPriceExclTax"].ToString()
                                                                //productDetails.Rows[productCount]["PriceInclTax"].ToString() 
                                                            };
                            table.Rows.Add(row);
                            productCount++;
                        }
                        StringBuilder sb = new StringBuilder();
                        string tab = "\t";
                        sb.AppendLine("<table border=0 class='MsoNormalTable'cellspacing='2'cellpadding='3'style='width:100%;font-size: 12.0pt; font-family: Times New Roman'>");

                        // Header of dynamic table.
                        sb.Append(tab + "<tr style=' text-align:center; font-size: 12.0pt; font-family: Times New Roman'>");

                        // Columns in dynamic table
                        foreach (DataColumn dc in table.Columns)
                        {
                            sb.AppendFormat("<td style=' background-color:#b9babe;font-size: 12.0pt; font-family: Times New Roman;'><strong>{0}</strong></td>", dc.ColumnName);
                        }
                        sb.AppendLine("</tr>");

                        // Total data rows in table.
                        foreach (DataRow dr in table.Rows)
                        {
                            sb.Append(tab + "<tr style='background-color: #ebecee;font-size: 12.0pt; font-family:Times New Roman'>");
                            //Actual Data in columns
                            foreach (DataColumn dc in table.Columns)
                            {
                                int rowIndex = table.Columns.IndexOf(dc);
                                string alignment = Alignment(rowIndex);

                                if (rowIndex.Equals(1) || rowIndex.Equals(3))
                                {
                                    string cellValue = dr[dc] != null ? dr[dc].ToString() : "";
                                    sb.AppendFormat("<td style='background-color: #ebecee; font-size: 12.0pt; font-family:Times New Roman;; text-align:"
                                                    + alignment + ";'>{0}</td>", ToCurrencyString(Convert.ToDecimal(cellValue)));
                                }
                                else
                                {
                                    string cellValue = dr[dc] != null ? dr[dc].ToString() : "";
                                    sb.AppendFormat("<td background-color: #ebecee;style='font-size: 12.0pt; font-family:Times New Roman; text-align:"
                                                    + alignment + ";'>{0}</td>", cellValue);
                                }
                            }

                            sb.AppendLine("</tr>");
                        }
                        //sb.AppendLine(tab + "</table>");
                        //mailBody.Append(sb);
                        //sb=<table border=0 class='MsoNormalTable'cellspacing='2'cellpadding='3'style='width:100%;font-size: 12.0pt; font-family: Times New Roman'> <tr style=' text-align:center; font-size: 12.0pt; font-family: Times New Roman'><td style=' background-color:#b9babe;font-size: 12.0pt; font-family: Times New Roman;'><strong>Name</strong></td><td style=' background-color:#b9babe;font-size: 12.0pt; font-family: Times New Roman;'><strong>Price</strong></td><td style=' background-color:#b9babe;font-size: 12.0pt; font-family: Times New Roman;'><strong>Quantity</strong></td><td style=' background-color:#b9babe;font-size: 12.0pt; font-family: Times New Roman;'><strong>Total</strong></td></tr> <tr style='background-color: #ebecee;font-size: 12.0pt; font-family:Times New Roman'><td background-color: #ebecee;style='font-size: 12.0pt; font-family:Times New Roman; text-align:left;'>Uber Pricing Strategies and Marketing Communications (PDF Download)</td><td style='background-color: #ebecee; font-size: 12.0pt; font-family:Times New Roman;; text-align:right;'>₹ 0.00</td><td background-colo}
                        content = content.Replace("<productTable>", sb.ToString());

                    }
                    #endregion

                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailOrderPlacedReceiptToAdmin",
                                "Completed the reading of each line using stream reader of HTML body in SendMailOrderPlacedReceiptToAdmin() method.",
                                null, strPurchaseModuleName);
                    //content =<html xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns:m="http://schemas.microsoft.com/office/2004/12/omml" xmlns="http://www.w3.org/TR/REC-html40"> <head> <meta http-equiv="Content-Type" content="text/html; charset=utf-8"> <meta name="Generator" content="Microsoft Word 14 (filtered medium)"> <!--[if !mso]><style>v\:* {behavior:url(#default#VML);}o\:* {behavior:url(#default#VML);}w\:* {behavior:url(#default#VML);}.shape {behavior:url(#default#VML);}</style><![endif]--> <title> Placed An Order From Your Store.</title> <style> <!-- /* Font Definitions */ @font-face { font-family: "Cambria Math"; panose-1: 2 4 5 3 5 4 6 3 2 4; } @font-face { font-family: Calibri; panose-1: 2 15 5 2 2 2 4 3 2 4; } @font-face { font-family: Tahoma; panose-1: 2 11 6 4 3 5 4 4 2 4; } @font-face { font-family: Times New Roman; panose-1: 2 11 6 4 3 5 4 4 2 4; } /* Style Definitions */ p.MsoNormal, li.MsoNormal, div.MsoNormal { margin: 0in; margin-bottom: .0001pt; font-size: 12.0pt; font-family: "Times New Roman"; } a:link, span.MsoHyperlink { mso-style-priority: 99; color: #0066CC; text-decoration: underline; } a:visited, span.MsoHyperlinkFollowed { mso-style-priority: 99; color: purple; text-decoration: underline; } p { mso-style-priority: 99; mso-margin-top-alt: auto; margin-right: 0in; mso-margin-bottom-alt: auto; margin-left: 0in; font-size: 12.0pt; font-family: "Times New Roman"; } span.small { mso-style-name: small; } span.price { mso-style-name: price; } span.EmailStyle20 { mso-style-type: personal-reply; font-family: "Calibri","sans-serif"; color: windowtext; } .MsoChpDefault { mso-style-type: export-only; font-size: 12.0pt; } @page WordSection1 { size: 8.5in 11.0in; margin: 1.0in 1.0in 1.0in 1.0in; } div.WordSection1 { page: WordSection1; } #_x0000_i1025 { height: 43px; } .MsoNormalTable { margin-right: 0px; } .auto-style1 { height: 108px; } .style1 { width: 535px; } .style2 { width: 14px; } .style3 { width: 97px; height: 15px; } .auto-style3 { width: 97px; height: 15px; } .auto-style4 { width: 74px; height: 15px; } .auto-style5 { width: 74px; height: 15px; } --> </style> </head> <body bgcolor="white" lang="EN-US" link=blue vlink=purple> <div align="left" class="WordSection1" > <!-- <table class="MsoNormalTable" border="0" cellspacing="0" cellpadding="0" style='margin-left: .5in; > <tr> <td style='padding: .75pt .75pt .75pt .75pt'> <div align="center"> <table width="100%"> <tr> <td>--> <p> <a href="http://192.168.10.143:8989//">Darden Business Publishing IndiaStorefront</a><br><br> <span style='font-size: 12.0pt; font-family: "Times New Roman"'> <a href="mailto:Vijay.Nagarsoge@workmethods.com"> (Vijay.Nagarsoge@workmethods.com) </a> has just placed an order from your store. Below is the summary of the order.</span> </p> <!-- <table width="80%" style='width: 100.0%'> <tr> <td nowrap style='padding: 3.0pt 3.0pt 3.0pt 3.0pt' class="style1">--> <p class="MsoNormal"> <span style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'>Order Number: </span><span style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'> </span> <span style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'> 241504 </span> </p> <p class="MsoNormal"> <span style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'>Date Ordered:</span><span style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'> </span> Friday, September 13, 2019 </p> <br /><br /> <div> <p class="MsoNormal" style='margin-bottom: 12.0pt'> <span style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'>Billing Address:</span><span style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'><br> Vijay Nagarsoge<br /> workmethods pune<br /> pune, <BState,> 411038<br /> India<br /> </span> 9096998615 </p> </div> <div> <p class="MsoNormal" style='margin-bottom: 12.0pt'> <span style='font-size: 12.0pt; font-family: "Times New Roman"'>Shipping Address:</span><span style='font-size: 12.0pt; font-family: "Times New Roman"'><br> Not Applicable<br /> <SCompany><br /> <SAddress1><br /> <SAddress2><br /> <SCity,> <SState,> <SZip><br /> <SCountry><br /> <br /> </span> </p> </div> <div class="shipping-method" style='font-size: 12.0pt; font-family: "Times New Roman","sans-serif"'> <div > Shipping Method: </div> <div class="shipping-method-value"> Not Applicable </div> </div> <!-- </td> </tr> <tr> <td colspan="2" style='padding: 3.0pt 3.0pt 3.0pt 3.0pt;'>--> <table border=0 class='MsoNormalTable'cellspacing='2'cellpadding='3'style='width:100%;font-size: 12.0pt; font-family: Times New Roman'> <tr style=' text-align:center; font-size: 12.0pt; font-family: Times New Roman'><td style=' background-color:#b9babe;font-size: 12.0pt; font-family: Times New Roman;'><strong>Name</strong></td><td style=' background-color:#b9babe;font-size: 12.0pt; font-family: Times New Roman;'><strong>Price</strong></td><td style=' background-color:#b9babe;font-size: 12.0pt; font-family: Times New Roman;'><strong>Quantity</strong></td><td style=' background-color:#b9babe;font-size: 12.0pt; font-family: Times New Roman;'><strong>Total</strong></td></tr> <tr style='background-color: #ebecee;font-size: 12.0pt; font-family:Times New Roman'><td background-color: #ebecee;style='font-size: 12.0pt; font-family:Times New Roman; text-align:left;'>Uber Pricing Strategies and Marketing Communications (PDF Download)</td><td style='background-color: #ebecee; font-size: 12.0pt; font-family:Times New Roman;; text-align:right;'>₹ 0.00</td><td background-color: #ebecee;style='font-size: 12.0pt; font-family:Times New Roman; text-align:right;'>1</td><td style='background-color: #ebecee; font-size: 12.0pt; font-family:Times New Roman;; text-align:right;'>₹ 0.00</td></tr> <!--<table border="0" style='width: 68%; font-size: 12.0pt; font-family: "Times New Roman","sans-serif"; margin-left: 0px;text-align:left'>--> <tr style='background:#DDE2E6'> <!-- <td style='background:white'></td>--><td style='background:white'></td> <td colspan="2"class="auto-style3"> <strong>Sub-Total: </strong> </td> <td class="auto-style4" style="text-align:right"> <currency>&nbsp;&nbsp;₹ 0.00 </td> </tr> <tr style='background:#DDE2E6'> <!--<td style='background:white'></td>--><td style='background:white'></td> <td colspan="2" class="style3"> <strong>GST(18%):</strong> </td> <td class="auto-style5" style="text-align:right"> <currency>&nbsp;&nbsp;₹ 0.00 </td> </tr> <tr style='background:#DDE2E6'> <!--<td style='background:white'></td>--><td style='background:white'></td> <td colspan="2" class="style3"> <strong>Order Total:</strong> </td> <td class="auto-style5" style="text-align:right"> <currency>&nbsp;&nbsp;₹ 0.00 </td> </tr> </table> <!--</p>--> <!-- </td> </tr> <tr> <td colspan="2" style='padding: 3.0pt 3.0pt 3.0pt 3.0pt' class="auto-style1"> &nbsp;</td> </tr> </table> </td> </tr> </table> </div> </td> </tr> </table>--> </div> </body> </html>
                    mail.Body = Convert.ToString(content);
                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendMailOrderPlacedReceiptToAdmin",
                                "Mail body formed and values assigned to respective parameters successfully.", null, strPurchaseModuleName);
                }

                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
                SmtpServer.Port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]);//25
                SmtpServer.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["NetworkCredentialMailID"],
                           ConfigurationManager.AppSettings["NetworkCredentialMailIDPassword"]);
                SmtpServer.EnableSsl = true;

                Console.WriteLine("Start order Placed mail sending.");
                SmtpServer.Send(mail);

                Console.WriteLine("Order Placed mail sent.");
                this.UpdateOrderReceiptSend(OrderId);
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Success, "GetConnection.SendMailOrderPlacedReceiptToAdmin",
                            "Order Placed Mail successfully sent ", null, strPurchaseModuleName);
            }

        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.SendMailOrderPlacedReceiptToAdmin",
                        "SendMailOrderPlacedReceiptToAdmin() method. Error: " + ex.Message, ex, strPurchaseModuleName);
        }
    }
    #endregion

    public bool SendInvoiceDetailsOfCustomerToAdmin(int OrderId)
    {
        try
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "Sending mail for the OrderNumber:"
                      + OrderId, null, strPurchaseModuleName);

            DataTable OrderData = GetOrderProductData(OrderId);
            DataTable productDetails = GetProductDetails(OrderId);
            int billingAddressId = Convert.ToInt32(OrderData.Rows[0]["BillingAddressId"]);
            DataTable billingAddress = GetCustomerAddress(billingAddressId);

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "After GetCustomerAddress", null, strPurchaseModuleName);

            MailMessage mail = new MailMessage();
            SmtpClient SmtpServer = new SmtpClient(ConfigurationManager.AppSettings["SmtpHost"]);

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "After SmtpHost", null, strPurchaseModuleName);

            mail.From = new MailAddress(ConfigurationManager.AppSettings["MailFrom"]);
            mail.To.Add(ConfigurationManager.AppSettings["OrderInvoiceReceiptMailTo"]);
            mail.Subject = ConfigurationManager.AppSettings["InvoiceCopyDetails"];
            mail.IsBodyHtml = true;

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "1", null, strPurchaseModuleName);

            string emailFormatFileLocation = ConfigurationManager.AppSettings["EmailBodyInvoice"].ToString();
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "2", null, strPurchaseModuleName);

            StreamReader reader = new StreamReader(emailFormatFileLocation);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "3", null, strPurchaseModuleName);

            string content = reader.ReadToEnd();
            reader.Close();
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "4", null, strPurchaseModuleName);

            content = content.Replace("<bemailid>", Convert.ToString(billingAddress.Rows[0]["Expr1"]));
            content = content.Replace("<id>", Convert.ToString(OrderId));
            content = content.Replace("<invoicenum>", Convert.ToString(ConfigurationManager.AppSettings["InvoiceNumber"] + Convert.ToString(OrderId)));

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "5", null, strPurchaseModuleName);

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
              "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "5.0 -" + OrderData.Rows[0]["CreatedOnUtc"] + "", null, strPurchaseModuleName);

            content = content.Replace("<orderdate>", Convert.ToString(String.Format("{0:dddd, MMMM d, yyyy}", (OrderData.Rows[0]["CreatedOnUtc"]))));

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
               "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "5.1 -" + OrderData.Rows[0]["CreatedOnUtc"] + "", null, strPurchaseModuleName);

            content = content.Replace("<busername>", Convert.ToString(billingAddress.Rows[0]["UserName"]));

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
           "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "5.2", null, strPurchaseModuleName);

            content = content.Replace("<buniversity>", Convert.ToString(billingAddress.Rows[0]["University"]));

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
           "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "5.3", null, strPurchaseModuleName);

            content = content.Replace("<customergstin>", Convert.ToString(billingAddress.Rows[0]["GSTINNumber"]));
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
           "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "5.4 - " + OrderData.Rows[0]["OrderSubtotalExclTax"] + "", null, strPurchaseModuleName);

            content = content.Replace("<ordersubtotalexcltax>", ToCurrencyString(Convert.ToDecimal(OrderData.Rows[0]["OrderSubtotalExclTax"])));

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
           "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "5.5 - " + productDetails.Rows[0]["OrderTax"] + "", null, strPurchaseModuleName);

            content = content.Replace("<ordertax>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTax"])));

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
           "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "5.6 - " + productDetails.Rows[0]["OrderTotal"] + "", null, strPurchaseModuleName);

            content = content.Replace("<ordertotal>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderTotal"])));

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
           "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "5.7 - " + productDetails.Rows[0]["OrderDiscount"] + "", null, strPurchaseModuleName);

            content = content.Replace("<orderdiscount>", ToCurrencyString(Convert.ToDecimal(productDetails.Rows[0]["OrderDiscount"])));

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
           "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "5.8-end", null, strPurchaseModuleName);

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "6", null, strPurchaseModuleName);

            if (content.Contains("<productTable>"))
            {
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                    "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "7", null, strPurchaseModuleName);

                string productLine = string.Empty;
                int productCount = 0;
                DataTable table = new DataTable("ProductsTable");
                DataColumn name = new DataColumn("Name", typeof(System.String));
                DataColumn price = new DataColumn("Unit Price", typeof(System.String));
                DataColumn quantity = new DataColumn("Quantity", typeof(System.String));
                DataColumn total = new DataColumn("Total", typeof(System.String));

                table.Columns.Add(name);
                table.Columns.Add(price);
                table.Columns.Add(quantity);
                table.Columns.Add(total);

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information,
                   "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "8", null, strPurchaseModuleName);

                while (productDetails.Rows.Count > productCount)
                {
                    string strProductName = productDetails.Rows[productCount]["Product"].ToString();
                    string strPrice = productDetails.Rows[productCount]["UnitPriceExclTax"].ToString();
                    string strDiscountedPrice = productDetails.Rows[productCount]["Discounted Price"].ToString();
                    if (Convert.ToString(strProductName) == string.Empty || Convert.ToString(strProductName) == Convert.ToString(""))
                    {
                        strProductName = productDetails.Rows[productCount]["CPProduct"].ToString();
                    }
                    else
                    {
                        strProductName = productDetails.Rows[productCount]["Product"].ToString();
                    }
                    //if (!String.IsNullOrEmpty(strDiscountedPrice))
                    //{
                    //    strPrice = Convert.ToString(strDiscountedPrice);
                    //}
                    ////else if (!String.IsNullOrEmpty(strDiscountedPrice) && Convert.ToString(strDiscountedPrice) != "")
                    ////{
                    ////    strPrice = productDetails.Rows[productCount]["UnitPriceInclTax"].ToString();
                    ////}
                    //else
                    //{
                    //    strPrice = productDetails.Rows[productCount]["Price"].ToString();
                    //}
                    DataRow row = table.NewRow();
                    row.ItemArray = new object[] {
                                                    strProductName.ToString(),
                                                    strPrice.ToString(),
                                                    productDetails.Rows[productCount]["Quantity"].ToString(),
                                                    productDetails.Rows[productCount]["UnitPriceExclTax"].ToString()
                                                 };
                    table.Rows.Add(row);
                    productCount++;
                }
                StringBuilder sb = new StringBuilder();
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "out of while loop-9",
                null, strPurchaseModuleName);

                content = content.Replace("<productTable>", sb.ToString());
                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "10",
                null, strPurchaseModuleName);
            }
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "11",
            null, strPurchaseModuleName);


            mail.Body = Convert.ToString(content);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "12",
            null, strPurchaseModuleName);

            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;
            SmtpServer.Port = Convert.ToInt32(ConfigurationManager.AppSettings["Port"]);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "13",
            null, strPurchaseModuleName);
            SmtpServer.Credentials = new System.Net.NetworkCredential(ConfigurationManager.AppSettings["NetworkCredentialMailID"],
            ConfigurationManager.AppSettings["NetworkCredentialMailIDPassword"]);
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "14",
            null, strPurchaseModuleName);
            SmtpServer.EnableSsl = true;

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.SendInvoiceDetailsOfCustomerToAdmin", "15",
            null, strPurchaseModuleName);
            SmtpServer.Send(mail);

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: SendInvoiceDetailsOfCustomerToAdmin",
                                                                         "Method SendInvoiceDetailsOfCustomerToAdmin executed successfully.", null, strPurchaseModuleName);
            return true;
        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.SendInvoiceDetailsOfCustomerToAdmin",
                       "SendInvoiceDetailsOfCustomerToAdmin() method. Error: " + ex.Message, null, strPurchaseModuleName);
            return false;
        }
    }//end of function

    //To get all the Supplement files from SF-Support folder.
    public void CopySupplementFiles(string folderName, string supportingFile, DirectoryInfo myDirSubFolder)
    {
        //To get all the supporting files from SF-Support folder.
        myDirSupport = new DirectoryInfo(ConfigurationManager.AppSettings["childFileLocation"] + folderName);

        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                    "Checking if supporting folder: " + myDirSupport + " exists", null, strPurchaseModuleName);

        //SF-Support Directory
        if (myDirSupport.Exists)
        {

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                        "Start getting supporting files. Path: " + myDirSupport, null, strPurchaseModuleName);

            if (myDirSubFolder == null)
            {
                myDirSubFolder = new DirectoryInfo(strCopyZIPFile + @"\" + folderName);

                myDirSubFolder.Create();

                ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                            "Created sub folder: " + myDirSubFolder, null, strPurchaseModuleName);
            }

            string[] allSupportingFiles = System.IO.Directory.GetFiles(myDirSupport.ToString());
            for (int i = 0; i < allSupportingFiles.Length; i++)
            {
                string strFile = Path.GetFullPath(allSupportingFiles[i]);
                string strFileNameExtension = Path.GetExtension(allSupportingFiles[i]).ToUpper();
                if (strFileNameExtension != ".MP4" && strFileNameExtension != ".JS")
                {
                    //Use static Path methods to extract only the file name from the path.
                    supportingFile = System.IO.Path.GetFileName(strFile);

                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess",
                             "Supporting file name: " + supportingFile, null, strPurchaseModuleName);

                    System.IO.File.Copy(strFile, myDirSubFolder.ToString() + @"\" + supportingFile, true);

                    ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection: ProductProcess",
                        "File " + supportingFile + " copied to location " + myDirSubFolder.ToString(), null, strPurchaseModuleName);
                }
            }

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                        "End getting supporting files. Path: " + myDirSupport, null, strPurchaseModuleName);

        }
        else
        {

            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                        "Supporting folder: " + strSlug.ToString() + " does not exist on location: " + childFileLocation,
                        null, strPurchaseModuleName);
        }
    }

    //To Insert supplementory video details
    public void InsertSupplementVideoDetails(int customerId, int intOrderId, int intParentgroupedProductId,
        string strFileLocation, string strVideoFileLocation, string strProductType)
    {
        if (strProductType != strConfigCoursePack)
        {
            string strProductName = Path.GetFileName(strFileLocation);
            if (Directory.Exists(strFileLocation))
            {
                //Identifying Video Supplements
                string[] strVideoFiles = Directory.GetFiles(strFileLocation, "*" + ".mp4");
                for (int i = 0; i <= strVideoFiles.Count() - 1; i++)
                {
                    if (strVideoFiles[i] != null)
                    {
                        int intRecords = 0;
                        Guid guid = Guid.NewGuid();
                        string strGuid = guid.ToString();
                        string strVideoName = Path.GetFileNameWithoutExtension(Path.GetFileName(strVideoFiles[i]));
                        try
                        {
                            SqlCommand myCommandSave = new SqlCommand();
                            myCommandSave.CommandType = CommandType.StoredProcedure;
                            myCommandSave.CommandText = "pr_uva_InsertSupplementVideoSubscriptionDetails";
                            myCommandSave.Parameters.AddWithValue("@OrderGuid", SqlDbType.VarChar).Value = strGuid;
                            myCommandSave.Parameters.AddWithValue("@CustomerId", SqlDbType.Int).Value = customerId;
                            myCommandSave.Parameters.AddWithValue("@OrderId", SqlDbType.Int).Value = intOrderId;
                            myCommandSave.Parameters.AddWithValue("@SubscriptionURL", SqlDbType.VarChar).Value =
                                                    strFacultySupport + strProductName + "/" + strVideoName + ".js";
                            myCommandSave.Parameters.AddWithValue("@ProductId", SqlDbType.Int).Value = intParentgroupedProductId;
                            myCommandSave.Parameters.AddWithValue("@SubscriptionDays", SqlDbType.Int).Value = intSubscriptionDays;
                            myCommandSave.Connection = thisConnection;
                            thisConnection.Open();
                            intRecords = (int)myCommandSave.ExecuteNonQuery();

                            //change log just pass mention the return value
                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                                        "Information saved in table (wm_SubscriptionDetails), for OrderId: " + intOrderId + ", CustomerID: "
                                        + customerId + ", No of days Subscribed:  ", null, strPurchaseModuleName);
                            thisConnection.Close();

                        }
                        catch (Exception ex)
                        {
                            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.ProductProcess",
                                         ex.Message, ex, strPurchaseModuleName);
                        }
                    }
                    else
                    {
                        // no video available
                        ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.ProductProcess",
                                        "No video available", null, strPurchaseModuleName);

                    }
                }
            }
        }

    }

    //Insert Supplement Video Details for Cource PAck
    public void InsertCoursePackSupplementVideoDetails(int intOrderId, int intCoursePackId, int CustomerId)
    {
        int intRecords = 0;
        int customerId = CustomerId;
        try
        {
            SqlCommand myCommandSave = new SqlCommand();
            myCommandSave.CommandType = CommandType.StoredProcedure;
            myCommandSave.CommandText = "pr_uva_InsertCoursePackSupplementVideoDetails";
            myCommandSave.Parameters.AddWithValue("@OrderId", SqlDbType.Int).Value = intOrderId;
            myCommandSave.Parameters.AddWithValue("@ProductId", SqlDbType.Int).Value = intCoursePackId;
            myCommandSave.Connection = thisConnection;
            thisConnection.Open();
            intRecords = (int)myCommandSave.ExecuteNonQuery();

            //change log just pass mention the return value
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Information, "GetConnection.singleOrderCompletion",
                        "Information saved in table (wm_SubscriptionDetails), for OrderId: " + intOrderId + ", CustomerID: "
                        + customerId + ", No of days Subscribed:  ", null, strPurchaseModuleName);
            thisConnection.Close();

        }
        catch (Exception ex)
        {
            ActivityLog.Log.LogActivity(ActivityLog.GlobalEnum.ActivityLog.Error, "GetConnection.ProductProcess",
                         ex.Message, ex, strPurchaseModuleName);
        }
    }
}