using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Xml;


namespace BizTalk.PipelineComponents.Excel.Common.Decoder
{
    public class ExcelRowSchema
    {
        public string Name { get; set; }
        public string Namespace { get; set; }

        public int Occurrence { get; set; }

        /// <summary>
        /// Specify which Row to start processing from
        /// </summary>
        public int Index { get; set; }

        private Dictionary<int, ExcelCellSchema> cells;
        public Dictionary<int, ExcelCellSchema> Cells
        {
            get
            {
                if (cells == null)
                {
                    cells = new Dictionary<int, ExcelCellSchema>();
                }

                return cells;
            }
        }

        public void Process(XmlWriter wtr, IRow sheet)
        {
            if (String.IsNullOrEmpty(this.Namespace))
            {
                wtr.WriteStartElement(this.Name);
            }
            else
            {
                wtr.WriteStartElement("r", this.Name, this.Namespace);
            }

            foreach (KeyValuePair<int, ExcelCellSchema> cell in this.Cells)
            {
                ExcelCellSchema eSchema = cell.Value;

                ICell c = sheet.GetCell(cell.Key, MissingCellPolicy.CREATE_NULL_AS_BLANK);
                if (c == null)
                    continue;

                eSchema.Process(wtr,c );
            }

            wtr.WriteEndElement();
        }
    }
}
