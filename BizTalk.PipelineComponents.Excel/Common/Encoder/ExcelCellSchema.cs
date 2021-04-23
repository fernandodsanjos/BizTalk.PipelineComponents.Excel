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

namespace BizTalk.PipelineComponents.Excel.Common.Encoder
{
    public class ExcelCellSchema
    {
          // if (dec.BaseXmlSchemaType.Datatype.TypeCode == XmlTypeCode.AnyAtomicType) (Primite type)
        public XmlTypeCode XmlType
        {
            get;set;
        }
       
        /// <summary>
        /// A = Attribute, E = Element
        /// </summary>
        public char NodeType { get; set; } = 'A';
        public int Index { get; set; }
        
        public string Name { get; set; }

        public IFormulaEvaluator FormulaEvaluator { get; set; }
       

        public void SetCellValue(string value,ICell cell)
        {
            //DateTime
            //double
            //string
            //bool
       
            switch (this.XmlType)
            {
                case XmlTypeCode.Boolean:
                    bool bVal = false;
                    if(Boolean.TryParse(value,out bVal))
                    {
                        cell.SetCellValue(bVal);
                    }
                    else
                        cell.SetCellValue(false);
                    break;
               
                case XmlTypeCode.Double:
                    Double dVal = 0;
                    if (Double.TryParse(value, out dVal))
                    {
                        cell.SetCellValue(dVal);
                    }
                    
                    break;

                case XmlTypeCode.DateTime:
                    DateTime dtVal;
                    if (DateTime.TryParse(value, out dtVal))
                    {
                        cell.SetCellValue(dtVal);
                    }
                   
                    break;
                default:
                    cell.SetCellValue(value);
                    break;
            }


       

            
        }

      
    }
}
