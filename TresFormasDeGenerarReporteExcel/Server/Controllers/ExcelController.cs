using ClosedXML.Excel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Data;

namespace TresFormasDeGenerarReporteExcel.Server.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        [HttpGet]
        [Route("PorCelda")]
        public IActionResult ExportExcel()
        {

            try
            {
                using var workbook = new XLWorkbook();
                var worksheet = workbook.AddWorksheet("Hoja1");
                worksheet.Cell("A1").Value = "Hola mundo, hola blazor!";
                worksheet.Cell("A2").FormulaA1 = "MID(A1, 7, 5)";

                using var memoria = new MemoryStream();
                workbook.SaveAs(memoria);
                var nombreExcel = "Reporte.xlsx";
                var archivo = File(memoria.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombreExcel);
                return archivo;

            }
            catch (Exception)
            {
                throw;

            }
        }


        [HttpGet]
        [Route("PorColumna")]
        public IActionResult ExportExcelPorColumna()
        {

            try
            {
                DataTable table = new DataTable();//tabla general


                table.Columns.Add("Nombre");
                table.Columns.Add("Curso");
                

                //Aqui van las filas y puedo añadir las filas que quiera o traerlas de una base de datos y con foreach interar

                DataRow fila = table.NewRow();
                fila["Nombre"] = "Dagoberto";
                fila["Curso"] = "Blazor webassembly";
                DataRow fila2 = table.NewRow();
                fila2["Nombre"] = "Andres";
                fila2["Curso"] = "Angular";

                table.Rows.Add(fila);
                table.Rows.Add(fila2);


                using var libro = new XLWorkbook();
                table.TableName = "Registros";

                var hoja = libro.Worksheets.Add(table);

                hoja.ColumnsUsed().AdjustToContents();
                //agregar tablas de tanques al excel

                using var memoria = new MemoryStream();
                libro.SaveAs(memoria);
                var nombreExcel = "Reporte.xlsx";
                var archivo = File(memoria.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombreExcel);
                return archivo;
            }
            catch (Exception)
            {
                throw;

            }
        }




        [HttpGet]
        [Route("Plantilla{Valor}")]
        public IActionResult ExportExcelPlantilla(string valor)
        {

            try
            {
                using (var workbook = new XLWorkbook(@"C:\Users\Dagoberto\Downloads\Nomina.xlsx"))
                {
                    var SampleSheet = workbook.Worksheets.Where(x => x.Name == "Empresa").First();



                    string CeldaItem = "C9";


                    //*************************************************

                    SampleSheet.Cell(CeldaItem).Value = valor;



                    using var memoria = new MemoryStream();
                    workbook.SaveAs(memoria);
                    var nombreExcel = "Reporte.xlsx";
                    var archivo = File(memoria.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", nombreExcel);
                    return archivo;
                }

            }
            catch (Exception)
            {
                throw;

            }
        }

    }
}
