using ExcelDataReader;
using LFHSystems.ExcelReader.App.Model;
using System.Data;
using System.Text;

namespace LFHSystems.ExcelReader.App.Application
{
    public class LeituraExcel
    {
        public void LerBaseClientesConsolidada()
        {
            try
            {
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                //string filePath = @"C:\Users\luan.hora\Downloads\exemploLeitura.xlsx";
                string filePath = @"C:\Users\luan.hora\Downloads\Base_Clientes_Consolidada_2024 (1).xlsx";

                FileStream stream = File.Open(filePath, FileMode.Open, FileAccess.Read);
                IExcelDataReader excelReader;

                //1. Reading Excel file
                if (Path.GetExtension(filePath).ToUpper() == ".XLS")
                {
                    //1.1 Reading from a binary Excel file ('97-2003 format; *.xls)
                    excelReader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else
                {
                    //1.2 Reading from a OpenXml Excel file (2007 format; *.xlsx)
                    excelReader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }

                //2. DataSet - The result of each spreadsheet will be created in the result.Tables
                DataSet result = excelReader.AsDataSet(new ExcelDataSetConfiguration() { ConfigureDataTable = (data) => new ExcelDataTableConfiguration() { UseHeaderRow = true } });

                //3. DataSet - Create column names from first row
                //excelReader. = false;

                List<ClienteConsolidadoExcel> lstClientesConsolidados = new List<ClienteConsolidadoExcel>();
                foreach (DataRow item in result.Tables["2024.S1"].Rows)
                {
                    lstClientesConsolidados.Add(new ClienteConsolidadoExcel()
                    {
                        GestaoExterna = item["Gestão Externa"]?.ToString()
                    });
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }
    }
}
