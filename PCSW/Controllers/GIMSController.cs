
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using PCSW.Models;
using System.Configuration;
using System.IO;



namespace PCSW.Controllers
{
    public class GIMSController : Controller
    {

        public static List<SelectListItem> GetDropDownListForYears()
        {
            List<SelectListItem> ls = new List<SelectListItem>();

            int currYear = DateTime.Now.Year;
            for (int i = currYear - 20 ; i <= currYear+5; i++)
            {
                ls.Add(new SelectListItem() { Text = i.ToString(), Value = i.ToString() });
            }

            return ls;
        }

        // GET: GIMS
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult Part1()
        {            
            return View();
        }

        public ActionResult Part2()
        {
            //@Html.DropDownList("yearPickerPart2" , (IEnumerable<SelectListItem>)ViewBag.Years , "Select Year" , new { @class = "dropdown-menu" })
            return View();
        }

        public ActionResult Part3()
        {
            
            return View();
        }

       
        
        public FileResult Download(string FileId)
        {
            try
            {             
                string virPath = "~/Exports/" + FileId + ".xls";
                string fullPath = Path.Combine(Server.MapPath(virPath));
                return File(fullPath, "application/vnd.ms-excel", "Report");
            }
            catch (Exception)
            {
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                return null;
            }
        }

        [HttpPost]
        public ActionResult ExportToXLS(DataModel d)
        {
            
            try
            {
                Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

                if (xlApp == null)
                {

                    return Json(new { success = false, responseText = "App is not installed on the server..." }, JsonRequestBehavior.AllowGet);
                }


                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                object misValue = System.Reflection.Missing.Value;

                xlWorkBook = xlApp.Workbooks.Add(misValue);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

                //---------------------Header---------------------
                xlWorkSheet.Range["A1:E1"].Merge();
                xlWorkSheet.Range["A1:E1"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[1, 1] = "Department Name:";
                xlWorkSheet.Cells[1, 1].Font.Bold = true;


                xlWorkSheet.Range["F1:I1"].Merge();
                xlWorkSheet.Range["F1:I1"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[1, 6] = "Focal Person:" + d.focalPerson;
                xlWorkSheet.Cells[1, 6].Font.Bold = true;

                xlWorkSheet.Range["J1:M1"].Merge();
                xlWorkSheet.Range["J1:M1"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[1, 10] = "Contact No:" + d.contactNumber;
                xlWorkSheet.Cells[1, 10].Font.Bold = true;

                //-----------------Sub-Header---------------------
                xlWorkSheet.Cells[2, 1] = "Sr.#";
                xlWorkSheet.Cells[2, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[2, 1].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[2, 1].Font.Bold = true;
                xlWorkSheet.Range["B2:E2"].Merge();
                xlWorkSheet.Range["B2:E2"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[2, 2] = "Indicator";
                xlWorkSheet.Cells[2, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[2, 2].Font.Bold = true;

                xlWorkSheet.Range["F2:G2"].Merge();
                xlWorkSheet.Range["F2:G2"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[2, 6] = "Male";
                xlWorkSheet.Cells[2, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[2, 6].Font.Bold = true;

                xlWorkSheet.Range["H2:I2"].Merge();
                xlWorkSheet.Range["H2:I2"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[2, 8] = "Female";
                xlWorkSheet.Cells[2, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[2, 8].Font.Bold = true;

                xlWorkSheet.Range["J2:K2"].Merge();
                xlWorkSheet.Range["J2:K2"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[2, 10] = "Total";
                xlWorkSheet.Cells[2, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[2, 10].Font.Bold = true;


                //-------------------P1----------------------------
                xlWorkSheet.Cells[3, 1] = "1";
                xlWorkSheet.Cells[3, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[3, 1].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[3, 1].Font.Bold = true;
                xlWorkSheet.Range["B3:E3"].Merge();
                xlWorkSheet.Range["B3:E3"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[3, 2] = "Number of Employees";
                xlWorkSheet.Cells[3, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F3:G3"].Merge();
                xlWorkSheet.Range["F3:G3"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[3, 6] = Int32.Parse(d.contractMale) + Int32.Parse(d.gazettedMale) + Int32.Parse(d.noneGazettedMale);
                xlWorkSheet.Cells[3, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                xlWorkSheet.Range["H3:I3"].Merge();
                xlWorkSheet.Range["H3:I3"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[3, 8] = Int32.Parse(d.contractFemale) + Int32.Parse(d.gazettedFemale) + Int32.Parse(d.noneGazettedFemale);
                xlWorkSheet.Cells[3, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["J3:K3"].Merge();
                xlWorkSheet.Range["J3:K3"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[3, 10] = Int32.Parse(d.contractMale) + Int32.Parse(d.gazettedMale) + Int32.Parse(d.noneGazettedMale) + Int32.Parse(d.contractFemale) + Int32.Parse(d.gazettedFemale) + Int32.Parse(d.noneGazettedFemale);
                xlWorkSheet.Cells[3, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                xlWorkSheet.Range["A3:A3"].Merge();
                xlWorkSheet.Range["A3:A3"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                //-------------------P2----------------------------
                xlWorkSheet.Cells[4, 1] = "2";
                xlWorkSheet.Cells[4, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[4, 1].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[4, 1].Font.Bold = true;
                xlWorkSheet.Range["B4:E4"].Merge();
                xlWorkSheet.Range["B4:E4"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[4, 2] = "Number of Gazetted Officers";
                xlWorkSheet.Cells[4, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F4:G4"].Merge();
                xlWorkSheet.Range["F4:G4"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[4, 6] = d.gazettedMale;
                xlWorkSheet.Cells[4, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                xlWorkSheet.Range["H4:I4"].Merge();
                xlWorkSheet.Range["H4:I4"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[4, 8] = d.gazettedFemale;
                xlWorkSheet.Cells[4, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["J4:K4"].Merge();
                xlWorkSheet.Range["J4:K4"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                xlWorkSheet.Cells[4, 10] = Int32.Parse(d.gazettedFemale) + Int32.Parse(d.gazettedMale);
                xlWorkSheet.Cells[4, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                //-------------------P3----------------------------
                xlWorkSheet.Range["A5:A6"].Merge();
                xlWorkSheet.Range["A5:A6"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[5, 1] = "3";
                xlWorkSheet.Cells[5, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[5, 1].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[5, 1].Font.Bold = true;
                xlWorkSheet.Range["B5:E6"].Merge();
                xlWorkSheet.Range["B5:E6"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[5, 2] = "Number of Non-Gazetted Officers/Officials";
                xlWorkSheet.Range["B5:E6"].WrapText = true;
                xlWorkSheet.Cells[5, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F5:G6"].Merge();
                xlWorkSheet.Range["F5:G6"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[5, 6] = d.noneGazettedMale;
                xlWorkSheet.Cells[5, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[5, 6].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;


                xlWorkSheet.Range["H5:I6"].Merge();
                xlWorkSheet.Cells[5, 8] = d.noneGazettedFemale;
                xlWorkSheet.Cells[5, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[5, 8].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["J5:K6"].Merge();
                xlWorkSheet.Range["J5:K6"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                xlWorkSheet.Cells[5, 10] = Int32.Parse(d.noneGazettedFemale) + Int32.Parse(d.noneGazettedMale);
                xlWorkSheet.Cells[5, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[5, 10].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //-------------------P4----------------------------
                xlWorkSheet.Range["A7:A8"].Merge();
                xlWorkSheet.Range["A7:A8"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[7, 1] = "4";
                xlWorkSheet.Cells[7, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[7, 1].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[7, 1].Font.Bold = true;
                xlWorkSheet.Range["B7:E8"].Merge();
                xlWorkSheet.Range["B7:E8"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["B7:E8"].WrapText = true;
                xlWorkSheet.Cells[7, 2] = "Number of Employees inducted on contractual basis";
                xlWorkSheet.Cells[7, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F7:G8"].Merge();
                xlWorkSheet.Range["F7:G8"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[7, 6] = d.contractMale;
                xlWorkSheet.Cells[7, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[7, 6].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;


                xlWorkSheet.Range["H7:I8"].Merge();
                xlWorkSheet.Range["H7:I8"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[7, 8] = d.contractFemale;
                xlWorkSheet.Cells[7, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[7, 8].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["J7:K8"].Merge();
                xlWorkSheet.Range["J7:K8"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[7, 10] = Int32.Parse(d.contractFemale) + Int32.Parse(d.contractMale);
                xlWorkSheet.Cells[7, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[7, 10].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;


                xlWorkSheet.Range["L2:M8"].Merge();
                xlWorkSheet.Range["L2:M8"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);


                //-------------------P5----------------------------
                xlWorkSheet.Range["A9:A12"].Merge();
                xlWorkSheet.Range["A9:A12"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[9, 1] = "5";
                xlWorkSheet.Cells[9, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[9, 1].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[9, 1].Font.Bold = true;
                xlWorkSheet.Range["B9:E12"].Merge();
                xlWorkSheet.Range["B9:E12"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["B9:E12"].WrapText = true;
                xlWorkSheet.Cells[9, 2] = "Number of Women Friendly amenities in public offices";
                xlWorkSheet.Cells[9, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[9, 2].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F9:G11"].Merge();
                xlWorkSheet.Range["F9:G11"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["F9:G11"].WrapText = true;
                xlWorkSheet.Cells[9, 6] = "No. of seperate washrooms for females";
                xlWorkSheet.Cells[9, 6].Font.Bold = true;
                xlWorkSheet.Cells[9, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[9, 6].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;


                xlWorkSheet.Range["H9:I11"].Merge();
                xlWorkSheet.Range["H9:I11"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["H9:I11"].WrapText = true;
                xlWorkSheet.Cells[9, 8] = "No. of seperate prayer areas for females";
                xlWorkSheet.Cells[9, 8].Font.Bold = true;
                xlWorkSheet.Cells[9, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[9, 8].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F12:G12"].Merge();
                xlWorkSheet.Range["F12:G12"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[12, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[12, 6] = d.washroomsFemale;

                xlWorkSheet.Range["H12:I12"].Merge();
                xlWorkSheet.Range["H12:I12"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[12, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[12, 8] = d.prayerRoomsFemale;

                xlWorkSheet.Range["J9:K12"].Merge();
                xlWorkSheet.Range["J9:K12"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                xlWorkSheet.Range["L9:M12"].Merge();
                xlWorkSheet.Range["L9:M12"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                ////////////////////////////////////////////////////////
                xlWorkSheet.Range["B13:E13"].Merge();
                xlWorkSheet.Range["B13:E13"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                xlWorkSheet.Range["F13:M13"].Merge();
                xlWorkSheet.Range["F13:M13"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                xlWorkSheet.Cells[13, 6] = d.year;
                xlWorkSheet.Cells[13, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;


                //-------------------P6----------------------------
                xlWorkSheet.Range["A14:A16"].Merge();
                xlWorkSheet.Range["A14:A16"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[14, 1] = "6";
                xlWorkSheet.Cells[14, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[14, 1].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[14, 1].Font.Bold = true;
                xlWorkSheet.Range["B14:E16"].Merge();
                xlWorkSheet.Range["B14:E16"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["B14:E16"].WrapText = true;
                xlWorkSheet.Cells[14, 2] = "Number of women appointed to whom age relaxation of upto 3 years was allowed";
                xlWorkSheet.Cells[14, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F14:M16"].Merge();
                xlWorkSheet.Range["F14:M16"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                xlWorkSheet.Cells[14, 6] = d.numAgeRelexation3;
                xlWorkSheet.Cells[14, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[14, 6].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //-------------------P7----------------------------
                xlWorkSheet.Range["A17:A18"].Merge();
                xlWorkSheet.Range["A17:A18"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[17, 1] = "7";
                xlWorkSheet.Cells[17, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[17, 1].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[17, 1].Font.Bold = true;
                xlWorkSheet.Range["B17:E18"].Merge();
                xlWorkSheet.Range["B17:E18"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["B17:E18"].WrapText = true;
                xlWorkSheet.Cells[17, 2] = "Number of Women who availed maternity leave";
                xlWorkSheet.Cells[17, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F17:M18"].Merge();
                xlWorkSheet.Range["F17:M18"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[17, 6] = d.numMaternityLeave;
                xlWorkSheet.Cells[17, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[17, 6].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //-------------------P8----------------------------
                xlWorkSheet.Range["A19:A20"].Merge();
                xlWorkSheet.Range["A19:A20"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[19, 1] = "8";
                xlWorkSheet.Cells[19, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[19, 1].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[19, 1].Font.Bold = true;
                xlWorkSheet.Range["B19:E20"].Merge();
                xlWorkSheet.Range["B19:E20"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["B19:E20"].WrapText = true;
                xlWorkSheet.Cells[19, 2] = "Number of Men who availed paternity leave";
                xlWorkSheet.Cells[19, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F19:M20"].Merge();
                xlWorkSheet.Range["F19:M20"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[19, 6] = d.numPaternityLeave;
                xlWorkSheet.Cells[19, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[19, 6].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //-------------------P9----------------------------
                xlWorkSheet.Range["A21:A24"].Merge();
                xlWorkSheet.Range["A21:A24"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[21, 1] = "9";
                xlWorkSheet.Cells[21, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[21, 1].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[21, 1].Font.Bold = true;
                xlWorkSheet.Range["B21:E24"].Merge();
                xlWorkSheet.Range["B21:E24"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["B21:E24"].WrapText = true;
                xlWorkSheet.Cells[21, 2] = "Number of Selection and Recruitment Committees for regular and contractual employment fulfilling the condition of at least one women representative";
                xlWorkSheet.Cells[21, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F21:M24"].Merge();
                xlWorkSheet.Range["F21:M24"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[21, 6] = d.numSelectionContractualCommittee;
                xlWorkSheet.Cells[21, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[21, 6].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                //-------------------P10----------------------------
                xlWorkSheet.Range["A25:A26"].Merge();
                xlWorkSheet.Range["A25:A26"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[25, 1] = "10";
                xlWorkSheet.Cells[25, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[25, 1].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[25, 1].Font.Bold = true;
                xlWorkSheet.Range["B25:E26"].Merge();
                xlWorkSheet.Range["B25:E26"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["B25:E26"].WrapText = true;
                xlWorkSheet.Cells[25, 2] = "Establishment of Gender Mainstreaming Committee";
                xlWorkSheet.Cells[25, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F25:G26"].Merge();
                xlWorkSheet.Range["F25:G26"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["F25:G26"].WrapText = true;
                xlWorkSheet.Cells[25, 6] = d.GMC;
                xlWorkSheet.Cells[25, 6].Font.Bold = true;
                xlWorkSheet.Cells[25, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[25, 6].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["J25:M29"].Merge();
                xlWorkSheet.Range["J25:M29"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                //-------------------P11----------------------------
                xlWorkSheet.Range["A27:A29"].Merge();
                xlWorkSheet.Range["A27:A29"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[27, 1] = "11";
                xlWorkSheet.Cells[27, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[27, 1].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[27, 1].Font.Bold = true;
                xlWorkSheet.Range["B27:E29"].Merge();
                xlWorkSheet.Range["B27:E29"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["B27:E29"].WrapText = true;
                xlWorkSheet.Cells[27, 2] = "Code of Conduct Implemented under Punjab Protection Against Harassment of Women at Workplace Act 2012";
                xlWorkSheet.Cells[27, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F27:G29"].Merge();
                xlWorkSheet.Range["F27:G29"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["F27:G29"].WrapText = true;
                xlWorkSheet.Cells[27, 6] = d.COCPunjabProtection;
                xlWorkSheet.Cells[27, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[27, 6].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["H27.I29"].Merge();
                xlWorkSheet.Range["H27.I29"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["H25.I26"].Merge();
                xlWorkSheet.Range["H25.I26"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);


                //-------------------P12----------------------------
                xlWorkSheet.Range["A30:A32"].Merge();
                xlWorkSheet.Range["A30:A32"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[30, 1] = "12";
                xlWorkSheet.Cells[30, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[30, 1].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[30, 1].Font.Bold = true;
                xlWorkSheet.Range["B30:E32"].Merge();
                xlWorkSheet.Range["B30:E32"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["B30:E32"].WrapText = true;
                xlWorkSheet.Cells[30, 2] = "Establishment of Workplace Harassment Committees";
                xlWorkSheet.Cells[30, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F30:G31"].Merge();
                xlWorkSheet.Range["F30:G31"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["F30:G31"].WrapText = true;
                xlWorkSheet.Cells[30, 6] = "Establishment of Committee";
                xlWorkSheet.Cells[30, 6].Font.Bold = true;
                xlWorkSheet.Cells[30, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;


                xlWorkSheet.Range["H30:I31"].Merge();
                xlWorkSheet.Range["H30:I31"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["H30:I31"].WrapText = true;
                xlWorkSheet.Cells[30, 8] = "No. of Complaints Received";
                xlWorkSheet.Cells[30, 8].Font.Bold = true;
                xlWorkSheet.Cells[30, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;

                xlWorkSheet.Range["J30:K31"].Merge();
                xlWorkSheet.Range["J30:K31"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["J30:K31"].WrapText = true;
                xlWorkSheet.Cells[30, 10] = "No. of Action Taken";
                xlWorkSheet.Cells[30, 10].Font.Bold = true;
                xlWorkSheet.Cells[30, 10].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F32:G32"].Merge();
                xlWorkSheet.Range["F32:G32"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["F32:G32"].WrapText = true;
                xlWorkSheet.Cells[32, 6] = d.workplaceHarassmentCommittees;
                xlWorkSheet.Cells[32, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["H32:I32"].Merge();
                xlWorkSheet.Range["H32:I32"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["H32:I32"].WrapText = true;
                xlWorkSheet.Cells[32, 8] = d.numComplaintsReceived;
                xlWorkSheet.Cells[32, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["J32:K32"].Merge();
                xlWorkSheet.Range["J32:K32"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["J32:K32"].WrapText = true;
                xlWorkSheet.Cells[32, 10] = d.numActionsTaken;
                xlWorkSheet.Cells[32, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["L30.M31"].Merge();
                xlWorkSheet.Range["L30.M31"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["L32.M32"].Merge();
                xlWorkSheet.Range["L32.M32"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                //-------------------P13----------------------------
                xlWorkSheet.Range["A33:A39"].Merge();
                xlWorkSheet.Range["A33:A39"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[33, 1] = "13";
                xlWorkSheet.Cells[33, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[33, 1].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[33, 1].Font.Bold = true;
                xlWorkSheet.Range["B33:E39"].Merge();
                xlWorkSheet.Range["B33:E39"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["B33:E39"].WrapText = true;
                xlWorkSheet.Cells[33, 2] = "Number of Women in all Boards, Committees and Special Purpose Taskforces";
                xlWorkSheet.Cells[33, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[33, 2].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F33:G36"].Merge();
                xlWorkSheet.Range["F33:G36"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                xlWorkSheet.Range["F37:G37"].Merge();
                xlWorkSheet.Range["F37:G37"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["F37:G37"].WrapText = true;
                xlWorkSheet.Cells[37, 6] = "Board(s)";
                xlWorkSheet.Cells[37, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F38:G38"].Merge();
                xlWorkSheet.Range["F38:G38"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["F38:G38"].WrapText = true;
                xlWorkSheet.Cells[38, 6] = "Committee(s)";
                xlWorkSheet.Cells[38, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F39:G39"].Merge();
                xlWorkSheet.Range["F39:G39"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["F39:G39"].WrapText = true;
                xlWorkSheet.Cells[39, 6] = "Taskforce(s)";
                xlWorkSheet.Cells[39, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                ////////////////////////////////////////

                xlWorkSheet.Range["H33:I36"].Merge();
                xlWorkSheet.Range["H33:I36"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["H33:I36"].WrapText = true;
                xlWorkSheet.Cells[33, 8] = "No. of Boards/Committees/ Task Forces";
                xlWorkSheet.Cells[33, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                xlWorkSheet.Cells[33, 8].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;


                xlWorkSheet.Range["H37:I37"].Merge();
                xlWorkSheet.Range["H37:I37"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["H37:I37"].WrapText = true;
                xlWorkSheet.Cells[37, 8] = d.numBoardBCT;
                xlWorkSheet.Cells[37, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["H38:I38"].Merge();
                xlWorkSheet.Range["H38:I38"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["H38:I38"].WrapText = true;
                xlWorkSheet.Cells[38, 8] = d.numCommitteeBCT;
                xlWorkSheet.Cells[38, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["H39:I39"].Merge();
                xlWorkSheet.Range["H39:I39"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["H39:I39"].WrapText = true;
                xlWorkSheet.Cells[39, 8] = d.numTaskForceBCT;
                xlWorkSheet.Cells[39, 8].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                ////////////////////////////////////////

                xlWorkSheet.Range["J33:K36"].Merge();
                xlWorkSheet.Range["J33:K36"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["J33:K36"].WrapText = true;
                xlWorkSheet.Cells[33, 10] = "No. of Male Members";
                xlWorkSheet.Cells[33, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                xlWorkSheet.Cells[33, 10].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;


                xlWorkSheet.Range["J37:K37"].Merge();
                xlWorkSheet.Range["J37:K37"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["J37:K37"].WrapText = true;
                xlWorkSheet.Cells[37, 10] = d.numMaleBoard;
                xlWorkSheet.Cells[37, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["J38:K38"].Merge();
                xlWorkSheet.Range["J38:K38"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["J38:K38"].WrapText = true;
                xlWorkSheet.Cells[38, 10] = d.numMaleCommittee;
                xlWorkSheet.Cells[38, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["J39:K39"].Merge();
                xlWorkSheet.Range["J39:K39"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["J39:K39"].WrapText = true;
                xlWorkSheet.Cells[39, 10] = d.numMaleTaskForce;
                xlWorkSheet.Cells[39, 10].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                ////////////////////////////////////////////
                xlWorkSheet.Range["L33:M36"].Merge();
                xlWorkSheet.Range["L33:M36"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["L33:M36"].WrapText = true;
                xlWorkSheet.Cells[33, 12] = "No. of Female Members";
                xlWorkSheet.Cells[33, 12].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                xlWorkSheet.Cells[33, 12].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;


                xlWorkSheet.Range["L37:M37"].Merge();
                xlWorkSheet.Range["L37:M37"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["L37:M37"].WrapText = true;
                xlWorkSheet.Cells[37, 12] = d.numFemaleBoard;
                xlWorkSheet.Cells[37, 12].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["L38:M38"].Merge();
                xlWorkSheet.Range["L38:M38"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["L38:M38"].WrapText = true;
                xlWorkSheet.Cells[38, 12] = d.numFemaleCommittee;
                xlWorkSheet.Cells[38, 12].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["L39:M39"].Merge();
                xlWorkSheet.Range["L39:M39"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["L39:M39"].WrapText = true;
                xlWorkSheet.Cells[39, 12] = d.numFemaleTaskForce;
                xlWorkSheet.Cells[39, 12].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                ////////////////////////////////////////
                xlWorkSheet.Range["N33:O36"].Merge();
                xlWorkSheet.Range["N33:O36"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["N33:O36"].WrapText = true;
                xlWorkSheet.Cells[33, 14] = "Tenure of Boards/Committees/ Task Forces";
                xlWorkSheet.Cells[33, 14].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                xlWorkSheet.Cells[33, 14].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;


                xlWorkSheet.Range["N37:O37"].Merge();
                xlWorkSheet.Range["N37:O37"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["N37:O37"].WrapText = true;
                xlWorkSheet.Cells[37, 14] = d.tenureOfBoard;
                xlWorkSheet.Cells[37, 14].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["N38:O38"].Merge();
                xlWorkSheet.Range["N38:O38"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["N38:O38"].WrapText = true;
                xlWorkSheet.Cells[38, 14] = d.tenureOfCommittee;
                xlWorkSheet.Cells[38, 14].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["N39:O39"].Merge();
                xlWorkSheet.Range["N39:O39"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["N39:O39"].WrapText = true;
                xlWorkSheet.Cells[39, 14] = d.tenureOfTaskForce;
                xlWorkSheet.Cells[39, 14].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                /////////////////////////////
                xlWorkSheet.Range["P33:P36"].Merge();
                xlWorkSheet.Range["P33:P36"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["P33:P36"].WrapText = true;
                xlWorkSheet.Cells[33, 16] = "Vacant Positions";
                xlWorkSheet.Cells[33, 16].HorizontalAlignment = Excel.XlHAlign.xlHAlignLeft;
                xlWorkSheet.Cells[33, 16].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Cells[37, 16] = d.vacantPositionsBoard;
                xlWorkSheet.Cells[38, 16] = d.vacantPositionsCommittee;
                xlWorkSheet.Cells[39, 16] = d.vacantPositionsTaskForce;

                xlWorkSheet.Range["P37:P37"].Merge();
                xlWorkSheet.Range["P37:P37"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                xlWorkSheet.Range["P38:P38"].Merge();
                xlWorkSheet.Range["P38:P38"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                xlWorkSheet.Range["P39:P39"].Merge();
                xlWorkSheet.Range["P39:P39"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);

                //-------------------P14----------------------------
                xlWorkSheet.Range["A40:A41"].Merge();
                xlWorkSheet.Range["A40:A41"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[40, 1] = "14";
                xlWorkSheet.Cells[40, 1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[40, 1].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[40, 1].Font.Bold = true;
                xlWorkSheet.Range["B40:E41"].Merge();
                xlWorkSheet.Range["B40:E41"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Range["B40:E41"].WrapText = true;
                xlWorkSheet.Cells[40, 2] = "No. of Trainings held for Board/Committee/Task Force Members";
                xlWorkSheet.Cells[40, 2].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

                xlWorkSheet.Range["F40:M41"].Merge();
                xlWorkSheet.Range["F40:M41"].BorderAround2(Excel.XlLineStyle.xlContinuous, Excel.XlBorderWeight.xlMedium, Excel.XlColorIndex.xlColorIndexAutomatic, Excel.XlColorIndex.xlColorIndexAutomatic);
                xlWorkSheet.Cells[40, 6] = d.numMemberTrainings;
                xlWorkSheet.Cells[40, 6].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                xlWorkSheet.Cells[40, 6].VerticalAlignment = Excel.XlHAlign.xlHAlignCenter;

                Guid g = Guid.NewGuid();

                string relPath = "~/Exports/" + g.ToString() + ".xls";
                string fullPath = Path.Combine(Server.MapPath(relPath));
                xlApp.DisplayAlerts = false;
                xlWorkBook.SaveAs(fullPath, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Save();

                xlWorkBook.Close(true, misValue, misValue);
                xlApp.Quit();

                Marshal.ReleaseComObject(xlWorkSheet);
                Marshal.ReleaseComObject(xlWorkBook);
                Marshal.ReleaseComObject(xlApp);

                return Json(new { success = true, fileId = g.ToString() }, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(new { success = false, fileId = "Something Went Wrong !!!" }, JsonRequestBehavior.AllowGet);
            }

}

        public ActionResult DownloadView()
        {
            return View();
        }
        
        public ActionResult ErrorView()
        {
            return View();
        }    
        
        public ActionResult SSNotSupported()
        {
            return View();
        }
    }
}