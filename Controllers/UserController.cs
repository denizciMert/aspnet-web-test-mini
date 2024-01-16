using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using akademetre_cs_test.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.EMMA;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace akademetre_cs_test.Controllers
{
    public class UserController : Controller
    {
        [HttpPost]
        public ActionResult Create(UserModel model)
        {
            if (ModelState.IsValid)
            {
                List<UserModel> userInfo = Session["KullaniciBilgileri"] as List<UserModel>;
                if (userInfo == null)
                {
                    userInfo = new List<UserModel>();
                    Session["KullaniciBilgileri"] = userInfo;
                }
                userInfo.Add(model);
                return RedirectToAction("Index");
            }
            return View(model);
        }

        [HttpPost]
        public ActionResult CreateExcel(List<List<string>> dataList)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();
                    WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet(new SheetData());
                    Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());
                    Sheet sheet = new Sheet()
                    {
                        Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = "Kullanıcı Bilgileri"
                    };
                    sheets.Append(sheet);
                    SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                    Row titleRow = new Row();
                    titleRow.Append(
                        new Cell { CellValue = new CellValue("Ad"), DataType = CellValues.String },
                        new Cell { CellValue = new CellValue("Soyad"), DataType = CellValues.String },
                        new Cell { CellValue = new CellValue("Adres"), DataType = CellValues.String },
                        new Cell { CellValue = new CellValue("Email"), DataType = CellValues.String }
                    );
                    sheetData.AppendChild(titleRow);
                    foreach (var veriRow in dataList)
                    {
                        Row row = new Row();
                        foreach (string veri in veriRow)
                        {
                            row.Append(new Cell
                            {
                                CellValue = new CellValue(veri),
                                DataType = CellValues.String
                            });
                        }
                        sheetData.AppendChild(row);
                    }
                    workbookPart.Workbook.Save();
                    memoryStream.Position = 0;                 
                }
                byte[] content = memoryStream.ToArray();
                return File(content, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "KullaniciBilgileri.xlsx");
            }
        }
    }
}