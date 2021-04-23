using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using System.Xml.Schema;
using System.Globalization;
namespace BizTalk.PipelineComponents.Excel.Common.Decoder
{
    public class ExcelCellSchema
    {
        private XmlTypeCode xmlxType = XmlTypeCode.String;

        // if (dec.BaseXmlSchemaType.Datatype.TypeCode == XmlTypeCode.AnyAtomicType)
        public XmlTypeCode XmlType
        {
            get
            {
                return xmlxType;
            }
            set
            {
               
               
            }
        }
       
        /// <summary>
        /// A = Attribute, E = Element
        /// </summary>
        public char NodeType { get; set; } = 'A';
        public int Index { get; set; }
        
        public string Name { get; set; }

        public IFormulaEvaluator FormulaEvaluator { get; set; }
        public void Process(XmlWriter wtr, ICell cell)
        {
           
            if (this.NodeType == 'E')
            {
                wtr.WriteElementString(this.Name, CellValue(cell));
            }
            else
            {
                wtr.WriteAttributeString(this.Name, CellValue(cell));
            }
            
        }

        private string CellValue(ICell cell)
        {

            string val = String.Empty;

            switch (cell.CellType)
            {
                case CellType.String:
                    val = cell.StringCellValue;
                    break;
                case CellType.Numeric:
                    //if (DateUtil.IsCellDateFormatted(cell))
                    var numericValue = cell.NumericCellValue;

                    if (this.XmlType == XmlTypeCode.Date)
                    {
                        DateTime date = cell.DateCellValue;
                        val = date.ToString("yyyy-MM-dd");

                    }
                    else if (this.XmlType == XmlTypeCode.DateTime)
                    {
                        DateTime date = cell.DateCellValue;
                        val = date.ToString("yyyy-MM-ddTHH:mm:ss");

                    }
                    else
                        val = numericValue.ToString(CultureInfo.InvariantCulture);



                    break;

                case CellType.Boolean:
                    val = cell.BooleanCellValue ? "true" : "false";
                    break;

                case CellType.Formula:
                    if (this.FormulaEvaluator != null)
                        val = this.FormulaEvaluator.EvaluateInCell(cell).ToString();
                    break;



            }

            return val;

        }

      
    }
}
