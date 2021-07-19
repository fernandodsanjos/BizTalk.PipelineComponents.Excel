using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;


namespace BizTalk.PipelineComponents.Excel.Common.Decoder
{
    public class ExcelEnvelopeSchema
    {
        public string Name { get; set; }
        public string Namespace { get;  set; }

      
        /// <summary>
        /// Specify which Sheet to process
        /// </summary>
        public int Index { get; set; }

        private Dictionary<int, ExcelRowSchema> rows;
        public Dictionary<int, ExcelRowSchema> Rows
        {
            get
            {
                if (rows == null)
                {
                    rows = new Dictionary<int, ExcelRowSchema>();
                }

                return rows;
            }
        }

        public  void Process(XmlWriter wtr,ISheet sheet)
        {
            if(String.IsNullOrEmpty(this.Namespace))
            {
                wtr.WriteStartElement( this.Name);
            }
            else
            {
                wtr.WriteStartElement("s", this.Name, this.Namespace);
            }
            

            foreach (KeyValuePair<int, ExcelRowSchema> row in this.Rows)
            {
                

                ExcelRowSchema eSchema = row.Value;

                if(eSchema.Occurrence > -1)
                {
                    for (int i = eSchema.Index; i < eSchema.Index + eSchema.Occurrence; i++)
                    {
                        IRow r = sheet.GetRow(i);

                        if (r == null)
                            continue;

                        eSchema.Process(wtr, sheet.GetRow(i));
                    }
                }
                else
                {
                    int length = sheet.LastRowNum + 1;
                  
                    for (int i = eSchema.Index; i < length; i++)
                    {
                        IRow r = sheet.GetRow(i);

                        if (r == null)
                            continue;

                        eSchema.Process(wtr,r);
                    }
                }
              
            }

            wtr.WriteEndElement();
        }

    }
}
