using ExcelFingerPrint.Models;
using LinqKit;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.Entity;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace ExcelFingerPrint.Controllers
{
    public class HomeController : Controller
    {
        private readonly DataContext db;
        public HomeController()
        {
            db = new DataContext();
        }
        public async Task<ActionResult> Index()
        {
            var data = await db.FingerPrintDatas.OrderBy(x => x.GuestID).Take(100).ToListAsync();
            var listGuestID = await db.FingerPrintDatas.GroupBy(x => x.GuestID.Trim()).Select(x => x.Key).ToListAsync();
            //var listEntryDoor = await db.FingerPrintDatas.GroupBy(x => x.EntryDoor.Substring(x.EntryDoor.IndexOf(":") + 1)).Select(x => x.Key).ToListAsync();
            var listEntryDoor = await db.FingerPrintDatas.GroupBy(x => x.EntryDoor.Trim()).Select(x => x.Key).ToListAsync();
            var result = new HomeViewModel
            {
                FingerPrintData = data,
                ListGuestID = listGuestID,
                ListEntryDoor = listEntryDoor
            };
            return View(result);
        }

        [HttpPost]
        public async Task<JsonResult> ImportExcel()
        {
            HttpPostedFileBase file = Request.Files[0];

            // Xóa tất cả dữ liệu tháng cũ
            int countColumn = 0;
            string filePath = string.Empty;
            if (file != null)
            {
                try
                {
                    string extension = Path.GetExtension(file.FileName);
                    if (extension == ".xls" || extension == ".xlsx")
                    {
                        string filename = string.Empty;
                        if (extension == ".xls")
                        {
                            filename = "Excel_FingerPrint" + ".xls";
                        }
                        else
                        {
                            filename = "Excel_FingerPrint" + ".xlsx";
                        }
                        string path = Server.MapPath("~/upload/") + filename;
                        //Nếu chưa tồn tại thư mục thì tạo thư mục
                        if (!Directory.Exists(Server.MapPath("~/upload/")))
                        {
                            Directory.CreateDirectory(path);
                        }
                        filePath = path;

                        //Nếu đã tồn tại file cũ xóa file cũ đi
                        if (System.IO.File.Exists(path))
                        {
                            System.IO.File.Delete(path);
                        }

                        file.SaveAs(filePath);
                        string conString = string.Empty;
                        switch (extension)
                        {
                            case ".xlsxxx":
                                conString = ConfigurationManager.ConnectionStrings["Excel03ConString"].ConnectionString;
                                break;

                            case ".xls":
                                conString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                                break;
                            case ".xlsx":
                                conString = ConfigurationManager.ConnectionStrings["Excel07ConString"].ConnectionString;
                                break;
                            default:
                                break;
                        }
                        DataTable dt = new DataTable();
                        conString = string.Format(conString, filePath);

                        //thêm cột Id vào file exel
                        dt.Columns.Add("Id");

                        using (OleDbConnection connExcel = new OleDbConnection(conString))
                        {
                            using (OleDbCommand cmdExcel = new OleDbCommand())
                            {
                                using (OleDbDataAdapter odaExcel = new OleDbDataAdapter())
                                {
                                    cmdExcel.Connection = connExcel;

                                    //Get the name of First Sheet.
                                    connExcel.Open();
                                    DataTable dtExcelSchema;
                                    dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                                    string sheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
                                    connExcel.Close();

                                    //Read Data from First Sheet.
                                    connExcel.Open();
                                    cmdExcel.CommandText = "SELECT * From [" + sheetName + "]";
                                    odaExcel.SelectCommand = cmdExcel;
                                    odaExcel.Fill(dt);
                                    countColumn = dt.Columns.Count;
                                    connExcel.Close();
                                }
                            }

                            //// xóa hết dữ liệu cũ đi
                            //var data = db.FingerPrintDatas.ToList();
                            //db.FingerPrintDatas.RemoveRange(data);
                            //await db.SaveChangesAsync();

                            //List<FingerPrintData> listFingerPrintData = new List<FingerPrintData>();
                            //foreach (DataRow dr in dt.Rows)
                            //{
                            //    FingerPrintData fingerPrintData = new FingerPrintData();

                            //    fingerPrintData.GuestID = dr[1].ToString().Trim();
                            //    fingerPrintData.CardNo = dr[2].ToString().Trim();
                            //    fingerPrintData.GuestName = dr[3].ToString().Trim();
                            //    fingerPrintData.Department = dr[4].ToString().Trim();
                            //    fingerPrintData.Date = dr[5].ToString().Trim();
                            //    fingerPrintData.Time = Convert.ToDateTime(dr[6].ToString().Trim());
                            //    fingerPrintData.EntryDoor = dr[7].ToString().Trim();
                            //    fingerPrintData.EventDescription = dr[8].ToString().Trim();
                            //    fingerPrintData.VerificationSource = dr[9].ToString().Trim();
                            //    fingerPrintData.Id = Guid.NewGuid().ToString();

                            //    listFingerPrintData.Add(fingerPrintData);
                            //}
                            //db.FingerPrintDatas.AddRange(listFingerPrintData);
                            //await db.SaveChangesAsync();

                            //chạy từng dòng cho id về kiểu guid
                            foreach (DataRow row in dt.Rows)
                            {
                                row["ID"] = Guid.NewGuid();
                            }
                            conString = ConfigurationManager.ConnectionStrings["DataContext"].ConnectionString;
                            using (SqlConnection con = new SqlConnection(conString))
                            {
                                using (SqlBulkCopy sqlBulkCopy = new SqlBulkCopy(con))
                                {
                                    if (countColumn == 10)
                                    {
                                        // xóa hết dữ liệu cũ đi
                                        var data = db.FingerPrintDatas.ToList();
                                        db.FingerPrintDatas.RemoveRange(data);
                                        await db.SaveChangesAsync();
                                        //Set the database table name.
                                        sqlBulkCopy.DestinationTableName = "dbo.FingerPrintData";

                                        //[OPTIONAL]: Map the Excel columns with that of the database table
                                        sqlBulkCopy.ColumnMappings.Add("ID", "Id");
                                        sqlBulkCopy.ColumnMappings.Add("GuestID", "GuestID");
                                        sqlBulkCopy.ColumnMappings.Add("CardNo", "CardNo");
                                        sqlBulkCopy.ColumnMappings.Add("GuestName", "GuestName");
                                        sqlBulkCopy.ColumnMappings.Add("Department", "Department");
                                        sqlBulkCopy.ColumnMappings.Add("Date", "Date");
                                        sqlBulkCopy.ColumnMappings.Add("Time", "Time");
                                        sqlBulkCopy.ColumnMappings.Add("EntryDoor", "EntryDoor");
                                        sqlBulkCopy.ColumnMappings.Add("EventDescription", "EventDescription");
                                        sqlBulkCopy.ColumnMappings.Add("VerificationSource", "VerificationSource");

                                        con.Open();
                                        sqlBulkCopy.WriteToServer(dt);
                                        con.Close();
                                    }
                                    else
                                    {
                                        object result = new
                                        {
                                            status = false,
                                            message = "Vui lòng chọn file đúng mẫu nhé!"
                                        };
                                        return Json(result, JsonRequestBehavior.AllowGet);
                                    }
                                }
                            }
                        }

                    }
                    else
                    {
                        object result = new
                        {
                            status = false,
                            message = "Vui lòng chọn đúng định dạng file excel nhé!"
                        };
                        return Json(result, JsonRequestBehavior.AllowGet);
                    }
                }
                catch (Exception ex)
                {
                    object result = new
                    {
                        status = false,
                        message = "Lỗi hệ thống rồi nhé!"
                    };
                    return Json(result, JsonRequestBehavior.AllowGet);
                    throw ex;
                }
            }
            else
            {
                object result = new
                {
                    status = false,
                    message = "Vui lòng chọn file để cập nhật hệ thống nhé!"
                };
                return Json(result, JsonRequestBehavior.AllowGet);
            }

            object result1 = new
            {
                status = true,
                message = "Upload dữ liệu thành công"
            };
            return Json(result1, JsonRequestBehavior.AllowGet);
        }

        [HttpPost]
        [Obsolete]
        public async Task<PartialViewResult> Search(string guestID, string[] listEntryDoor)
        {
            var query = db.FingerPrintDatas.AsQueryable();
            if (guestID != string.Empty && guestID != null)
            {
                query = query.Where(x => x.GuestID.Trim() == guestID.Trim());
            }
            if (listEntryDoor != null)
            {
                var predicate = PredicateBuilder.False<FingerPrintData>();
                foreach (var item in listEntryDoor)
                {
                    //predicate = predicate.Or(x => x.EntryDoor.Substring(x.EntryDoor.IndexOf(":") + 1) == item);
                    predicate = predicate.Or(x => x.EntryDoor.Trim() == item.Trim());
                }
                query = query.Where(predicate);
            }


            var data = await query.OrderBy(x => x.EntryDoor).ToListAsync();
            TempData["DataSearchFingerPrint"] = data;
            return PartialView("_Data", data);
        }

        [HttpPost]
        public ActionResult ExportExcel(string excelTitle, string excelName)
        {
            var data = TempData["DataSearchFingerPrint"] as List<FingerPrintData>;
            using (ExcelPackage excel = new ExcelPackage())
            {
                var workSheet = excel.Workbook.Worksheets.Add("Sheet1");
                workSheet.TabColor = System.Drawing.Color.Black;
                workSheet.DefaultRowHeight = 12;
                //Header of table  
                // 
                workSheet.Cells[1, 1].Value = excelTitle;
                workSheet.Row(1).Height = 30;
                workSheet.Cells["A1:I1"].Merge = true;
                workSheet.Row(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Row(1).Style.Font.Bold = true;
                workSheet.Row(1).Style.Font.Size = 20;

                workSheet.Row(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                workSheet.Row(2).Style.Font.Bold = true;

                workSheet.Cells[2, 1].Value = "GuestID";
                workSheet.Cells[2, 2].Value = "CardNo";
                workSheet.Cells[2, 3].Value = "GuestName";
                workSheet.Cells[2, 4].Value = "Department";
                workSheet.Cells[2, 5].Value = "Date";
                workSheet.Cells[2, 6].Value = "Time";
                workSheet.Cells[2, 7].Value = "EntryDoor";
                workSheet.Cells[2, 8].Value = "EventDescription";
                workSheet.Cells[2, 9].Value = "VerificationSource";
                //Body of table  
                //  
                int recordIndex = 3;
                foreach (var item in data)
                {
                    workSheet.Cells[recordIndex, 1].Value = item.GuestID;
                    workSheet.Cells[recordIndex, 2].Value = item.CardNo;
                    workSheet.Cells[recordIndex, 3].Value = item.GuestName;
                    workSheet.Cells[recordIndex, 4].Value = item.Department;
                    workSheet.Cells[recordIndex, 5].Value = item.Date;
                    workSheet.Cells[recordIndex, 6].Value = Convert.ToDateTime(item.Time).ToString("HH:mm:ss");
                    workSheet.Cells[recordIndex, 7].Value = item.EntryDoor;
                    workSheet.Cells[recordIndex, 8].Value = item.EventDescription;
                    workSheet.Cells[recordIndex, 9].Value = item.VerificationSource;
                    recordIndex++;
                }
                workSheet.Column(1).AutoFit();
                workSheet.Column(2).AutoFit();
                workSheet.Column(3).AutoFit();
                workSheet.Column(4).AutoFit();
                workSheet.Column(5).AutoFit();
                workSheet.Column(6).AutoFit();
                workSheet.Column(7).AutoFit();
                workSheet.Column(8).AutoFit();
                workSheet.Column(9).AutoFit();
                excel.Save();
                //Create buffer memory stream to catch excel file
                using (var buffer = excel.Stream as MemoryStream)
                {
                    var name = excelName == string.Empty ? "FingerPrintData" : excelName;
                    //This is the content type for excel file
                    Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                    Response.AddHeader("content-disposition", "attachment; filename=" + name + ".xlsx");
                    //Save this excel file as a byte array to return response
                    Response.BinaryWrite(buffer.ToArray());
                    //Sent all output bytes to clients
                    Response.Flush();
                    Response.End();
                }
            }
            return RedirectToAction("Index");
        }
    }
}