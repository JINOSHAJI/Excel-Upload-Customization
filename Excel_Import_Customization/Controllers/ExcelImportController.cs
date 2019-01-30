using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using static Enums;

namespace Excel_Import_Customization.Controllers
{
    [AllowAnonymous]
    public class ExcelImportController : Controller
    {
        static DataTable dtExclUpldPblc = new DataTable();
        private static ExcelPackage ExcelPackage_DemandSending = new ExcelPackage();
        DataTable dtGlobal = new DataTable();


        // GET: ExcelImport
        public ActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UploadExcel()
        {
            List<UploadedFileInfo> fileList = new List<UploadedFileInfo>();
            UploadedFileInfo filedata = new UploadedFileInfo();
            String Savedfilepath = null;
            ServiceResult response = new ServiceResult();

            dtExclUpldPblc.PrimaryKey = null; dtGlobal.Clear();
            for (int i = 0; i < dtExclUpldPblc.Columns.Count;)
                dtExclUpldPblc.Columns.RemoveAt(0);
            for (int i = 0; i < dtGlobal.Columns.Count;)
                dtGlobal.Columns.RemoveAt(0);
            for (int i = 0; i < dtExclUpldPblc.Rows.Count;)
                dtExclUpldPblc.Rows.RemoveAt(0);
            for (int i = 0; i < dtGlobal.Rows.Count;)
                dtGlobal.Rows.RemoveAt(0);

            if (Request.Files.Count > 0)
            {
                try
                {
                    HttpFileCollectionBase files = Request.Files;
                    int i = 0;

                    string path = AppDomain.CurrentDomain.BaseDirectory + "Uploads/";
                    String Tagname = HttpContext.Request.Params["tags"];

                    string filename = Path.GetFileName(Request.Files[i].FileName);
                    string File_Title = Path.GetFileName(Request.Files[i].FileName);
                    string ext = System.IO.Path.GetExtension((Request.Files[i].FileName));

                    if (ext != ".xlsx")
                    {
                        response.Message = "Excel File Not Found or Invalid Excel File! Upload a valid excel file (.xlsx). It Support Only .xlsx! ";
                        response.Status = Enums.ServiceStatus.Failure;
                        return Json(response, JsonRequestBehavior.AllowGet);
                    }


                    if (!(Tagname != null || Tagname != ""))
                    {
                        filename = Tagname + "_" + filename;
                        File_Title = Tagname;
                    }
                    HttpPostedFileBase file = files[i];
                    string fname;
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                        fname = testfiles[testfiles.Length - 1];
                    }
                    else
                    {
                        string date = DateTime.Now.ToString();
                        fname = date.Replace(" ", "").Replace("-", "").Replace(":", "") + filename;
                        //fname = file.FileName;
                    }

                    DirectoryInfo di = Directory.CreateDirectory(Server.MapPath("~/" + "srvUploads/Demand_Collection/DemandCollectionExcel"));
                    Savedfilepath = Path.Combine(Server.MapPath("~/" + "srvUploads/Demand_Collection/DemandCollectionExcel"), fname);
                    file.SaveAs(Savedfilepath);

                    FileInfo fi = new FileInfo(Server.MapPath("~/" + "srvUploads/Demand_Collection/DemandCollectionExcel") + "/" + fname);
                    if (!fi.Exists)
                    {
                        response.Message = "File Transfer Failed!File Transfer to the server failed.Try again!";
                        response.Status = Enums.ServiceStatus.Failure;
                        return Json(response, JsonRequestBehavior.AllowGet);
                    }

                    filedata.Filetype = Path.GetExtension(Savedfilepath);
                    filedata.Filename = fname;
                    filedata.Title = File_Title;
                    filedata.DocumentUrl = "srvUploads/Demand_Collection/DemandCollectionExcel" + "/" + fname;
                    // DataTable dtExcel = ImportToDataTable(Server.MapPath("~/excelTemp/" + FileUpload1.FileName));

                   

                    // decimal InterestRate =Convert.ToDecimal(HttpContext.Request.Params["txtInterestRate"]);
 
                    //  response = Service.getDemandCollectionExcelFile(Month, filedata.DocumentUrl);

                    

                    if (response.Status == Enums.ServiceStatus.Failure)
                    {
                        response.Data = 1;

                        response.Message = "Unexpected Error while saving file details to database! ";
                        response.Status = Enums.ServiceStatus.Failure;
                        return Json(response, JsonRequestBehavior.AllowGet);
                    }

                    DataTable dtExcel = ImportToDataTable(Server.MapPath("~/srvUploads/Demand_Collection/DemandCollectionExcel") + "/" + fname);

                    // response.Message = "Invalid Excel Template! This is not the valid excel template for student. please use the excel template downloaded from this site";

                    //response.Status = Enums.ServiceStatus.Failure;
                    //response.Data = dtExcel;
                    //return Json(response, JsonRequestBehavior.AllowGet);


                    Int32 cnt = dtExcel.Rows.Count;
                    int countRows = 0;
                    if (cnt == 0)
                    {
                        response.Message = "Invalid Excel Template! This is not the valid excel template for demand collection. please use the excel template downloaded from this site";
                        response.Data = 1;
                        response.Status = Enums.ServiceStatus.Failure;
                        return Json(response, JsonRequestBehavior.AllowGet);
                    }
                    

                    System.Web.Script.Serialization.JavaScriptSerializer serializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                    List<Dictionary<string, object>> rows = new List<Dictionary<string, object>>();
                    Dictionary<string, object> row;
                    foreach (DataRow dr in dtExcel.Rows)
                    {
                        row = new Dictionary<string, object>();
                        foreach (DataColumn col in dtExcel.Columns)
                        {
                            row.Add(col.ColumnName, dr[col]);
                        }
                        rows.Add(row);
                    }
                  
                    response.Data = (rows);

                    JsonResult jsonResult = new JsonResult();
                    jsonResult.MaxJsonLength = Int32.MaxValue;

                    jsonResult = Json(response);
                    jsonResult.MaxJsonLength = Int32.MaxValue;
                    jsonResult.JsonRequestBehavior = JsonRequestBehavior.AllowGet;
                    return jsonResult;
                     

                  //  response = ImportDataFrom_Excel(dtExcel, Month);

                    if (response.Status != Enums.ServiceStatus.Success)
                        return Json(response, JsonRequestBehavior.AllowGet);

                   
                    return Json(response, JsonRequestBehavior.AllowGet);

                    //return Json(filedata);
                }
                catch (Exception ex)
                {
                    response.Data = 1;

                    response.Message = "Error occurred. Error details: " + ex.Message;
                    response.Status = Enums.ServiceStatus.Failure;
                    //  response.Data = errList;
                    return Json(response, JsonRequestBehavior.AllowGet);
                }
            }
            else
            {
                response.Data = 1;

                response.Message = "Upload a valid excel file!. Excel File Not Found!";
                response.Status = Enums.ServiceStatus.Failure;
                return Json(response);
            }
        }


        public ServiceResult ImportDataFrom_Excel(DataTable dtExcel, string Month)
        {
            ServiceResult response = new ServiceResult();
           
            /*
            List<DemandCollectionModel> errList = new List<DemandCollectionModel>();
            List<DemandCollectionModel> DemandCollection_List = new List<DemandCollectionModel>();
            for (Int32 i = 0; i < dtExcel.Rows.Count; i++)
            {

                DemandCollectionModel DemandCollection = new DemandCollectionModel();
                DemandCollection.BadgeNo = dtExcel.Rows[i]["Badge No"].ToString().Trim();
                DemandCollection.Name = dtExcel.Rows[i]["Name"].ToString().Trim().Replace("'", "''");

                List<string> Errorreason = new List<string>();




                if (DemandCollection.BadgeNo.Trim() == "")
                    Errorreason.Add("Badge No is Mandator");
                if (DemandCollection.Name.Trim() == "")
                    Errorreason.Add("Name is Mandatory");

                try
                {
                    DemandCollection.Limit = Convert.ToDecimal(dtExcel.Rows[i]["Limit"]);
                    DemandCollection.Approved_Amount = DemandCollection.Limit;

                }
                catch (FormatException e)
                {
                    Limit = dtExcel.Rows[i]["Limit"].ToString();

                    Errorreason.Add("Invalid Limit");
                }

                DemandCollection.MonthAndYear = Month;

                if (DemandCollection.Limit == null)
                    Errorreason.Add("Limit is Mandatory");

                if (Errorreason.Count <= 0)
                    DemandCollection_List.Add(DemandCollection);
                else
                {

                    errList.Add(new DemandCollectionModel
                    {
                        BadgeNo = DemandCollection.BadgeNo,
                        Name = DemandCollection.Name,
                        Limit = DemandCollection.Limit,
                        Month = DemandCollection.Month,
                        Year = DemandCollection.Year
                    });
                    //write error
                    string ErrorReasons = String.Join(", ", Errorreason);
                    dtExclUpldPblc.Rows.Add(++countRows, DemandCollection.BadgeNo, DemandCollection.Name, DemandCollection.Limit, ErrorReasons);
                }

            }
*/
            

                ServiceResult saveUserResult = new ServiceResult();
               
               if (saveUserResult.Status == Enums.ServiceStatus.Failure)
                {
                    response.Message = "Upload Error !Some Error Occured During Uploading. Please download and refer the log file to view the demand collections that has not uploaded.";
                    response.Status = Enums.ServiceStatus.Failure;
                  
                    return response;
                }

                response.Message = "Demand Collection details uploaded successfully!";
                response.Status = Enums.ServiceStatus.Success;
               
                // BindStudent();
                return response;
           

        }

        public DataTable ImportToDataTable(string FilePath)
        {
            DataTable dt = new DataTable();
            FileInfo fi = new FileInfo(FilePath);
            //   DataTable dt1 = new DataTable();
            // Check if the file exists
            if (!fi.Exists)
            {
                throw new Exception("File " + FilePath + " Does Not Exists");
            }
            using (ExcelPackage xlPackage = new ExcelPackage(fi))
            {
                // get the first worksheet in the workbook
                ExcelWorksheet worksheet = xlPackage.Workbook.Worksheets[1];

                // Fetch the WorkSheet size
                ExcelCellAddress startCell = worksheet.Dimension.Start;
                ExcelCellAddress endCell = worksheet.Dimension.End;

                // create all the needed DataColumn

                for (int col = 1; col <= endCell.Column; col++)
                {
                    //if (string.IsNullOrEmpty(worksheet.Cells[startCell.Row, col].Text))
                    //    return dt;
                    if (worksheet.Cells[startCell.Row, col].Text == "")
                    {
                        dt.Columns.Add(worksheet.Cells[startCell.Row, col].Address);
                        dtGlobal.Columns.Add(worksheet.Cells[startCell.Row, col].Address);
                    }
                    else
                    {
                        dt.Columns.Add(worksheet.Cells[startCell.Row, col].Text);
                        dtGlobal.Columns.Add(worksheet.Cells[startCell.Row, col].Text);
                    }
                    //   dt1.Columns.Add(col.ToString());
                }

                try
                {
                    dtExclUpldPblc.Columns.Add("ErrorReason");
                }
                catch (DuplicateNameException)
                {

                }
                // place all the data into DataTable
                for (int row = 2; row <= endCell.Row; row++)
                {
                    DataRow dr = dt.NewRow();
                    DataRow drglb = dtGlobal.NewRow();
                    //  DataRow dr1 = dt1.NewRow();
                    int x = 0;

                    for (int col = startCell.Column; col <= endCell.Column; col++)
                    {
                        //if (string.IsNullOrEmpty(worksheet.Cells[row, col].Text))
                        //{

                        //    if ((col != 1) && (col != 10) && (col != 11) && (col != 13) && (col != 14))
                        //    {
                        //        isExcelCellEmpty = true; break;
                        //    }

                        //}


                        dr[x] = worksheet.Cells[row, col].Text;
                        drglb[x] = worksheet.Cells[row, col].Text.Trim(); x++;
                    }
                    //if (!isExcelCellEmpty)
                    //{
                    dt.Rows.Add(dr);
                    dtGlobal.Rows.Add(drglb);
                    //}
                }
            }
            fi.Delete();

            return dt;
        }

    }
}


public class ServiceResult
{
    /// <summary>
    /// Gets or sets the data.
    /// </summary>
    /// <value>
    /// The data is an object type member variable.
    /// </value>
    public object Data { get; set; }
    /// <summary>
    /// Gets or sets the status.
    /// </summary>
    /// <value>
    /// The status.
    /// </value>
    public Enums.ServiceStatus Status { get; set; }
    /// <summary>
    /// Gets or sets the message.
    /// </summary>
    /// <value>
    /// The message.
    /// </value>
    public string Message { get; set; }
    /// <summary>
    /// Gets or sets the token.
    /// </summary>
    /// <value>
    /// The token.
    /// </value>
    public string Token { get; set; }
    /// <summary>
    /// Gets a value indicating whether this <see cref="ServiceResult"/> is success.
    /// </summary>
    /// <value>
    ///   <c>true</c> if success; otherwise, <c>false</c>.
    /// </value>
    public bool Success
    {
        get
        {
            if (Status == ServiceStatus.Success)
            {
                return true;
            }
            else
            {
                return false;
            };
        }
    }
}

public class Enums
{


    /// <summary>
    /// UserRole
    /// </summary>
    public enum UserRole
    {
        SUPER_ADMIN = 10,
        SYSTEM_ADMIN = 20,
        SHOP_ADMIN = 30,
        STAFF = 40
    }

    /// <summary>
    /// RoleStatus
    /// </summary>
    public enum RoleStatus
    {
        /// <summary>
        /// The active
        /// </summary>
        Active = 1,
        /// <summary>
        /// The disabled
        /// </summary>
        Disabled = 2
    }

    /// <summary>
    /// ServiceStatus
    /// </summary>
    public enum ServiceStatus
    {
        /// <summary>
        /// The success
        /// </summary>
        Success = 1,
        /// <summary>
        /// The failure
        /// </summary>
        Failure = 2,
        /// <summary>
        /// The record exists
        /// </summary>
        RecordExists = 3,
        /// <summary>
        /// The no record exists
        /// </summary>
        NoRecordExists = 4,
        /// <summary>
        /// The reference exists
        /// </summary>
        ReferenceExists = 5
    }

    /// <summary>
    /// ResponseMessageType
    /// </summary>
    public enum ResponseMessageType
    {
        /// <summary>
        /// The information
        /// </summary>
        Info = 1,
        /// <summary>
        /// The error
        /// </summary>
        Error = 0,
        /// <summary>
        /// The warning
        /// </summary>
        Warning = 2,
        /// <summary>
        /// The confirmation
        /// </summary>
        Confirmation = -1
    }

    /// <summary>
    /// ResponseStatus
    /// </summary>
    public enum ResponseStatus
    {
        /// <summary>
        /// The failure
        /// </summary>
        Failure = 0,
        /// <summary>
        /// The success
        /// </summary>
        Success
    }



    /// <summary>
    /// LeadStatus
    /// </summary>
    public enum LeadStatus
    {
        /// <summary>
        /// The hot
        /// </summary>
        Hot = 1,
        /// <summary>
        /// The warm
        /// </summary>
        Warm = 2,
        /// <summary>
        /// The cold
        /// </summary>
        Cold = 3,
        /// <summary>
        /// The booked
        /// </summary>
        Booked = 4
    }

    /// <summary>
    /// LeadCategory
    /// </summary>
    public enum LeadCategory
    {
        /// <summary>
        /// All
        /// </summary>
        All = 0,
        /// <summary>
        /// The direct
        /// </summary>
        Direct = 1,
        /// <summary>
        /// The enquiry
        /// </summary>
        Enquiry = 2
    }
    /// <summary>
    /// NotificationStatus
    /// </summary>
    public enum NotificationStatus
    {
        /// <summary>
        /// The pending
        /// </summary>
        Pending = 0,
        /// <summary>
        /// The processed
        /// </summary>
        Processed = 1
    }
    /// <summary>
    /// Privilege
    /// </summary>
    public enum Privilege
    {
        /// <summary>
        /// The save customer
        /// </summary>
        SaveCustomer = 1,
        /// <summary>
        /// The update customer
        /// </summary>
        UpdateCustomer,
        /// <summary>
        /// The delete customer
        /// </summary>
        DeleteCustomer
    }
}

public class UploadedFileInfo
{

    public string DocumentUrl { get; set; }
    public string Title { get; set; }
    public string Filename { get; set; }
    public string Filetype { get; set; }

}