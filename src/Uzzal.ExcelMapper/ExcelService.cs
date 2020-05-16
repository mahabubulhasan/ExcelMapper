using ExcelDataReader;
using System.Data;
using System.Linq;
using System.IO;
using System;

namespace Uzzal.ExcelMapper
{
    public class ExcelService
    {
        private DataSet Data { get; set; }

        public ExcelService(Stream stream)
        {
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            ReadStrem(stream);
        }

        public ExcelService(string filePath)
        {
            ReadFile(filePath);
        }

        private void ReadFile(string filePath)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                ReadStrem(stream);
        }

        private void ReadStrem(Stream stream)
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream))
            {
                Data = reader.AsDataSet();
            }
        }

        private DataTable ReadSheet(int sheetNo = 0)
        {
            return Data.Tables[sheetNo];            
        }

        private void TreadFirstRowAsColumnName(DataTable sheet)
        {
            var header = sheet.Rows[0].ItemArray;
            for (int i = 0; i < header.Length; i++)
            {
                sheet.Columns[i].ColumnName = header[i].ToString();
            }
        }

        public DataRowCollection GetRows(int sheetNo = 0, bool treadFirstRowAsColumn = true)
        {
            var sheet = ReadSheet(sheetNo);
            if (treadFirstRowAsColumn)
            {
                TreadFirstRowAsColumnName(sheet);
            }

            return sheet.Rows;
        }
        
        public bool HasColumns(string[] columnNames, int sheetNo = 0)
        {
            var header = ReadSheet(sheetNo).AsEnumerable().First().ItemArray;
            foreach (string item in columnNames)
            {
                if(!header.Any(h => h.ToString().Trim().Equals(item)))
                {
                    return false;
                }
            }
            return true;
        }
    }
}
