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
    public class ExcelSheetSchema
    {
        public string Name { get; set; }
        public string Namespace { get;  set; }

        public bool IsEnvelope { get; set; }
        //In case of envelope, which row that is unbounded
        public int EnvelopeRow { get; set; }

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
            if(this.IsEnvelope)
            {
                ProcessEnvelope(wtr, sheet);
            }
            else
            {
                ProcessRegular(wtr, sheet);
            }
          
        }


        private void ProcessEnvelope(XmlWriter wtr, ISheet sheet)
        {
            //Row will act as group node
            
            int length = sheet.LastRowNum + 1;

            for (int x = EnvelopeRow; x < length; x++)
            {
                if (String.IsNullOrEmpty(this.Namespace))
                {
                    wtr.WriteStartElement(this.Name);
                }
                else
                {
                    wtr.WriteStartElement("s", this.Name, this.Namespace);
                }

                IRow er = sheet.GetRow(x);

                if (er == null)
                    continue;

                foreach (KeyValuePair<int, ExcelRowSchema> row in this.Rows)
                {
                    IRow r = null;

                    ExcelRowSchema eSchema = row.Value;

                    if (eSchema.Index == EnvelopeRow)
                    {
                        eSchema.Process(wtr, er);
                    }
                    else
                    {
                        if (eSchema.Occurrence > -1)
                        {
                            for (int i = eSchema.Index; i < eSchema.Index + eSchema.Occurrence; i++)
                            {
                                r = sheet.GetRow(i);

                                if (r == null)
                                    continue;

                                eSchema.Process(wtr, sheet.GetRow(i));
                            }
                        }
                        else
                        {
                            r = sheet.GetRow(eSchema.Index);

                            eSchema.Process(wtr, r);
                        }
                    }

                   

                }


                wtr.WriteEndElement();

            }
        }

        private void ProcessRegular(XmlWriter wtr, ISheet sheet)
        {
            if (String.IsNullOrEmpty(this.Namespace))
            {
                wtr.WriteStartElement(this.Name);
            }
            else
            {
                wtr.WriteStartElement("s", this.Name, this.Namespace);
            }


            foreach (KeyValuePair<int, ExcelRowSchema> row in this.Rows)
            {


                ExcelRowSchema eSchema = row.Value;

                if (eSchema.Occurrence > -1)
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

                        eSchema.Process(wtr, r);
                    }
                }

            }

            wtr.WriteEndElement();
        }

    }
}
