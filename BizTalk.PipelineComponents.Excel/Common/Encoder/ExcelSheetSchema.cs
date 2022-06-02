using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;


namespace BizTalk.PipelineComponents.Excel.Common.Encoder
{
    public class ExcelSheetSchema
    {
        public string Name { get; set; }
        public string Namespace { get;  set; }
        public bool IsVirtual { get; set; }

        /// <summary>
        /// Specify which Sheet to process
        /// </summary>
        public int Index { get; set; }

        private Dictionary<string, ExcelRowSchema> rows;
        public Dictionary<string, ExcelRowSchema> Rows
        {
            get
            {
                if (rows == null)
                {
                    rows = new Dictionary<string, ExcelRowSchema>();
                }

                return rows;
            }
        }

        public void Process(XmlReader reader, ISheet sheet)
        {

            while (reader.Read())
            {
                if (reader.IsStartElement() && reader.Depth == 1)
                {
                    ExcelRowSchema rowSchema = this.Rows[reader.LocalName];
                   
                    rowSchema.Process(reader.ReadSubtree(), sheet);
                   
                }
            }
        }

    }
}
