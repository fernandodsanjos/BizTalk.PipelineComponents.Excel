using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Xml;
using System.Xml.Schema;


namespace BizTalk.PipelineComponents.Excel.Common.Encoder
{
    public class ExcelRowSchema
    {
       
        public string Name { get; set; }
        public string Namespace { get; set; }

        public int Occurrence { get; set; }
        public int Processed { get; set; }

        public ICellStyle DateStyle { get; set; }
        public ICellStyle DateTimeStyle { get; set; }

        /// <summary>
        /// Specify which Row to start processing from
        /// </summary>
        public int Index { get; set; }

        private Dictionary<string, ExcelCellSchema> cells;
        public Dictionary<string, ExcelCellSchema> Cells
        {
            get
            {
                if (cells == null)
                {
                    cells = new Dictionary<string, ExcelCellSchema>();
                }

                return cells;
            }
        }

        public void Process(XmlReader reader,ISheet sheet)
        {
            SetStyles(sheet);

            while (reader.Read())
            {
                
                int nextIndex = this.Index + this.Processed;

                if (this.Occurrence > -1 && this.Processed > this.Occurrence)
                    throw new XmlSchemaValidationException($"Rows processed {this.Processed } for {this.Name} exceedes allowed {this.Occurrence}");

                

                IRow row = sheet.GetRow(nextIndex);
                
                if (row == null)
                    row = sheet.CreateRow(nextIndex);

                for (int i = 0; i < reader.AttributeCount; i++)
                {
                    reader.MoveToAttribute(i);
                    string name = reader.LocalName;
                    string value = reader.GetAttribute(name);

                    //cell.SetCellType(CellType.FORMULA);
                    //cell.SetCellFormula(String.Format("$B$1*B{0}/$B$2*C{0}", i));

                    if (this.Cells.ContainsKey(name))
                    { 
                        ExcelCellSchema cellSchema = this.Cells[name];
                        ICell cell = row.GetCell(cellSchema.Index);

                        if (cell == null)
                            cell = row.CreateCell(cellSchema.Index);//cellType

                        cellSchema.SetCellValue(value, cell);
                    }


                }

                while (reader.Read())
                {

                    if (reader.IsStartElement())
                    {

                        if (this.Cells.ContainsKey(reader.LocalName))
                        {
                            ExcelCellSchema cellSchema = this.Cells[reader.LocalName];
                            ICell cell = row.GetCell(cellSchema.Index);

                            if (cell == null)
                                cell = row.CreateCell(cellSchema.Index);//cellType

                            cellSchema.SetCellValue(GetText(reader), cell);
                        }
                        
                    }
                }

                
            }

            this.Processed++;
        }

        private  string GetText(XmlReader reader)
        {
            if (reader.IsEmptyElement)
                return String.Empty;

            if (reader.HasValue)
                return reader.Value;

            while (reader.Read())
            {
                if (reader.NodeType == XmlNodeType.EndElement)
                    break;

                if (reader.NodeType == XmlNodeType.Text)
                {
                    return reader.Value;
                }
            }

            return String.Empty;
        }

        private void SetStyles(ISheet sheet)
        {
            DateStyle = sheet.Workbook.CreateCellStyle();
            DateStyle.DataFormat = sheet.Workbook.CreateDataFormat().GetFormat("yyyy-MM-dd");

            DateTimeStyle = sheet.Workbook.CreateCellStyle();
            DateTimeStyle.DataFormat = sheet.Workbook.CreateDataFormat().GetFormat("yyyy-MM-dd HH:mm:ss");
        }
    }
}
