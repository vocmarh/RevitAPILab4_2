using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RevitAPILab4_2
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            List<Pipe> pipes = new FilteredElementCollector(doc)
                .OfClass(typeof(Pipe))
                .Cast<Pipe>()
                .ToList();

            string excelPath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "pipes.xlsx");

            using (FileStream stream = new FileStream(excelPath, FileMode.Create, FileAccess.Write))
            {
                IWorkbook workBook = new XSSFWorkbook();
                ISheet sheet = workBook.CreateSheet("Лист 1");

                int rowIndex = 0;
                foreach (var pipe in pipes)
                {
                    string pipeName = pipe.Name;
                    double outerDiamParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble();
                    double outerDiam = UnitUtils.ConvertFromInternalUnits(outerDiamParam, DisplayUnitType.DUT_MILLIMETERS);
                    double innerDiamParam = pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble();
                    double innerDiam = UnitUtils.ConvertFromInternalUnits(innerDiamParam, DisplayUnitType.DUT_MILLIMETERS);
                    double lengthParam = pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble();
                    double length = UnitUtils.ConvertFromInternalUnits(lengthParam, DisplayUnitType.DUT_MILLIMETERS);
                    sheet.SetCellValue(rowIndex, columnIndex: 0, pipeName);
                    sheet.SetCellValue(rowIndex, columnIndex: 1, outerDiam);
                    sheet.SetCellValue(rowIndex, columnIndex: 2, innerDiam);
                    sheet.SetCellValue(rowIndex, columnIndex: 3, length);
                    rowIndex++;
                }
                workBook.Write(stream);
                workBook.Close();
            }

            return Result.Succeeded;
        }
    }
}
