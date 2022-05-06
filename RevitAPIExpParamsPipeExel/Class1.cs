using Autodesk.Revit.ApplicationServices;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Plumbing;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;


namespace RevitAPIExpParamsPipeExel
{
    [Transaction(TransactionMode.Manual)]
    public class Main : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIApplication uiapp = commandData.Application;
            UIDocument uidoc = uiapp.ActiveUIDocument;
            Document doc = uidoc.Document;

            var saveFileDialog = new SaveFileDialog
            {
                OverwritePrompt = true,
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                Filter = "All files (*.*)|*.*",
                FileName = "PipeInfo.xlsx",
                DefaultExt = ".xlsx"
            };

            string selectedFilePath = string.Empty;
            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                selectedFilePath = saveFileDialog.FileName;
            }

            if (string.IsNullOrEmpty(selectedFilePath))
                return Result.Cancelled;

            List<Pipe> allPipes = new FilteredElementCollector(doc)
                 .OfClass(typeof(Pipe))
                 .Cast<Pipe>()
                 .ToList();

            List<PipeParam> pipeParams = new List<PipeParam>();

            string allText = string.Empty;
            foreach (var pipe in allPipes)
            {

                PipeParam pipeParam = new PipeParam
                {
                    PipeType = pipe.Name,
                    PipeLength = UnitUtils.ConvertFromInternalUnits(pipe.get_Parameter(BuiltInParameter.CURVE_ELEM_LENGTH).AsDouble(), UnitTypeId.Meters).ToString(),
                    PipeInnerDiam = UnitUtils.ConvertFromInternalUnits(pipe.get_Parameter(BuiltInParameter.RBS_PIPE_INNER_DIAM_PARAM).AsDouble(), UnitTypeId.Millimeters).ToString(),
                    PipeOuterDiam = UnitUtils.ConvertFromInternalUnits(pipe.get_Parameter(BuiltInParameter.RBS_PIPE_OUTER_DIAMETER).AsDouble(), UnitTypeId.Millimeters).ToString()
                };
                pipeParams.Add(pipeParam);

            }


            using (var fs = new FileStream(selectedFilePath, FileMode.Create, FileAccess.Write))
            {

                IWorkbook workbook = new XSSFWorkbook();

                ISheet sheet1 = workbook.CreateSheet("Sheet1");

                IRow row = sheet1.CreateRow(0);
                row.Height = 80 * 10;
                row.CreateCell(0).SetCellValue("Имя типа");
                row.CreateCell(1).SetCellValue("Длина трубы,м");
                row.CreateCell(2).SetCellValue("Внутренний диаметр, мм");
                row.CreateCell(3).SetCellValue("Внешний диаметр, мм");
                sheet1.AutoSizeColumn(0);

                var rowIndex = 1;

                foreach (var pipe in pipeParams)
                {

                    row = sheet1.CreateRow(rowIndex);
                    row.Height = 10 * 80;
                    row.CreateCell(0).SetCellValue(pipe.PipeType);
                    row.CreateCell(1).SetCellValue(pipe.PipeLength);
                    row.CreateCell(2).SetCellValue(pipe.PipeInnerDiam);
                    row.CreateCell(3).SetCellValue(pipe.PipeOuterDiam);
                    sheet1.AutoSizeColumn(0);
                    rowIndex++;
                }

                workbook.Write(fs);
                fs.Close();
            }

            return Result.Succeeded;
        }
    }
}
