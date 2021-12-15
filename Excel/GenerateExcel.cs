using GemBox.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Reflection;

namespace WorkFiles.Excel
{
    public class GenerateExcel<T> where T : class
    {
        /// <summary>
        /// Cria e salva o aqruivo
        /// </summary>
        /// <param name="data">Lista de Objetos Genéricos para construção do conteúdo do arquivo</param>
        /// <param name="obj">Objeto modelo para definir as colunas</param>
        /// <param name="localPath">Local onde será salvo o arquivo</param>
        /// <param name="fileName">Nome do arquivo</param>
        /// <param name="maxTextSize">Número de caracteres(Description) do campo de modelo para definição do tamanho das colunas</param>
        public void Create(List<T> data, T obj, string localPath, string fileName, int maxTextSize = 26)
        {
            try
            {
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");

                ExcelFile ef = new ExcelFile();
                ExcelWorksheet ws = ef.Worksheets.Add("Sheet1");

                if (!Directory.Exists(localPath))
                    Directory.CreateDirectory(localPath);

                CreateHeader(ref ws, obj, maxTextSize);
                CreateBody(ref ws, data);

                ef.Save(localPath + "//" + fileName);
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private List<string> Header(T obj)
        {
            var headers = new List<string>();
            var prop = obj.GetType().GetProperties();

            foreach (var p in prop)
            {
                string name = GetDisplayNameProprety(p.CustomAttributes);
                headers.Add(name);
            }

            return headers;
        }

        private string GetDisplayNameProprety(IEnumerable<CustomAttributeData> attData)
        {
            try
            {
                string name = "";
                foreach (var item in attData)
                {                    
                    name = item.ConstructorArguments[0].Value.ToString();
                }

                return name;
            }
            catch (Exception)
            {
                throw new Exception("É necessário informar o 'Description' de todos os campo do objeto.");
            }
        }

        private object[,] GetRowsData(List<T> data)
        {
            int row = 0;
            var prop = typeof(T).GetProperties();
            object[,] objs = new object[data.Count + 1, prop.Length];

            foreach (var item in data)
            {
                int colum = 0;
                foreach (var p in prop)
                {
                    if (p.PropertyType.IsGenericType && !p.PropertyType.Equals(typeof(int?)) && !p.PropertyType.Equals(typeof(int)))
                    {
                        var list = (List<string>)p.GetValue(item);
                        objs[row, colum] = string.Join(",", list);
                    }
                    else if (p.PropertyType.Equals(typeof(DateTime)))
                    {
                        objs[row, colum] = ((DateTime)p.GetValue(item)).ToString("yyyy-MM-dd");
                    }
                    else
                    {
                        objs[row, colum] = p.GetValue(item);
                    }
                    colum++;
                }
                row++;
            }
            return objs;
        }
        
        private void CreateHeader(ref ExcelWorksheet ws, T obj, int maxTextSize)
        {
            int i = 0;
            foreach (var name in Header(obj))
            {
                CellStyle tmpStyle = new CellStyle();

                tmpStyle.Borders.SetBorders(MultipleBorders.All, Color.Black, LineStyle.Thin);
                tmpStyle.Font.Weight = ExcelFont.BoldWeight;
                tmpStyle.FillPattern.SetSolid(Color.LightGray);

                ws.Cells[0, i].Value = name;
                ws.Cells[0, i].Style = tmpStyle;
                ws.Columns[i].Width = maxTextSize * 256;

                i++;
            }
        }

        private void CreateBody(ref ExcelWorksheet ws, List<T> data)
        {
            var prop = typeof(T).GetProperties();

            for (int i = 0; i < data.Count; i++)
            {
                for (int j = 0; j < prop.Length; j++)
                {
                    if (prop[j].PropertyType.IsGenericType && !prop[j].PropertyType.Equals(typeof(int?)) && !prop[j].PropertyType.Equals(typeof(int)) && !prop[j].PropertyType.Equals(typeof(DateTime?)))
                    {
                        var list = (List<string>)prop[j].GetValue(data[i]);
                        ws.Cells[i + 1, j].Value = string.Join(",", list);

                    }
                    else if ((prop[j].PropertyType.Equals(typeof(DateTime)) || prop[j].PropertyType.Equals(typeof(DateTime?))) && prop[j].GetValue(data[i]) != null)
                    {
                        ws.Cells[i + 1, j].Value = (DateTime)prop[j].GetValue(data[i]);
                        ws.Cells[i + 1, j].Style.NumberFormat = "dd/mm/yyyy";
                    }
                    else
                    {
                        ws.Cells[i + 1, j].Value = prop[j].GetValue(data[i]);
                    }

                    ws.Cells[i + 1, j].Style.WrapText = false;
                }
            }
        }
    }
}
