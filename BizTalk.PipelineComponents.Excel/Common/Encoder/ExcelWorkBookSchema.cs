using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Schema;
using System.Xml;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;


namespace BizTalk.PipelineComponents.Excel.Common.Encoder
{
    public class ExcelWorkBookSchema
    {
        
        private Dictionary<string, ExcelSheetSchema> sheets;

        public string Name { get; private set; }
        public string Namespace { get; private set; }


        public ExcelWorkBookSchema(XmlSchema schema, IFormulaEvaluator formulaEvaluator)
        {

            if (schema.Items.Count > 1)
                throw new ArgumentException("Source schema is only allowed to have one root node");

            XmlSchemaElement workbook = (XmlSchemaElement)schema.Items[0];

            this.Namespace = workbook.QualifiedName.Namespace;
            this.Name = workbook.QualifiedName.Name;

            if (!(workbook.SchemaType is XmlSchemaComplexType))
                throw new ArgumentException("Nodes bellow root node must be of record type");

            //first must be complex type containing sequences of Sheet type objects
            XmlSchemaComplexType tp = (XmlSchemaComplexType)workbook.SchemaType;

            XmlSchemaSequence seq = (XmlSchemaSequence)tp.Particle;

            for (int i = 0; i < seq.Items.Count; i++)
            {
                XmlSchemaElement sheet = (XmlSchemaElement)seq.Items[i];

                int sheetIndex = GetIndex(sheet.Annotation, sheet.Name);

                if (sheetIndex == -1)
                    sheetIndex = i;

                ExcelSheetSchema eSheet = new ExcelSheetSchema
                {
                    Name = sheet.Name,
                    Namespace = sheet.QualifiedName.Namespace,
                    Index = sheetIndex
                };

                this.Sheets.Add(sheet.Name, eSheet);
                
               

                XmlSchemaComplexType sheetType = (XmlSchemaComplexType)sheet.SchemaType;

                XmlSchemaSequence sheetSequence = (XmlSchemaSequence)sheetType.Particle;

                for (int x = 0; x < sheetSequence.Items.Count; x++)
                {
                    XmlSchemaElement row = (XmlSchemaElement)sheetSequence.Items[x];

                    int rowIndex = GetIndex(row.Annotation, row.Name);

                    if (rowIndex == -1)
                        rowIndex = x;

                    XmlSchemaComplexType rowType = (XmlSchemaComplexType)row.SchemaType;

                    ExcelRowSchema eRow = new ExcelRowSchema
                    {
                        Name = row.Name,
                        Namespace = row.QualifiedName.Namespace,
                        Index = rowIndex,
                        Occurrence = row.MaxOccursString == "unbounded"? -1:(int)row.MaxOccurs

                    };

                    if (rowType.Attributes.Count > 0)
                    {

                        foreach (XmlSchemaAttribute item in rowType.Attributes)
                        {
                            XmlSchemaSimpleType sp = item.AttributeSchemaType;
                            
                          //  XmlSchemaDatatype tp;
                          //  XmlTypeCode.DateTime
                            int cellIndex = GetIndex(item.Annotation,item.Name);

                            if (cellIndex > -1)
                            { 
                                eRow.Cells.Add(item.Name, new ExcelCellSchema
                                {
                                    XmlType = GetExcelType(sp.Datatype.TypeCode),
                                    Index = cellIndex,
                                    Name = item.Name,
                                    NodeType = 'A',
                                    FormulaEvaluator = formulaEvaluator
                                });
                            }

                              
                        }


                    }

                    if (rowType.Particle != null)
                    {
                        XmlSchemaSequence rowSequence = (XmlSchemaSequence)rowType.Particle;

                        foreach (XmlSchemaElement item in rowSequence.Items)
                        {


                            
                            int cellIndex = GetIndex(item.Annotation,item.Name);

                            if (cellIndex > -1)
                            { 
                                
                                  eRow.Cells.Add(item.Name, new ExcelCellSchema
                                  {
                                      XmlType = GetExcelType(item.ElementSchemaType.Datatype.TypeCode),
                                      Index = cellIndex,
                                      Name = item.Name,
                                      NodeType = 'E'
                                  });

                            }


                        }
                    }

                    eSheet.Rows.Add(row.Name, eRow);
                }



            }

          

        }

        public Dictionary<string, ExcelSheetSchema> Sheets
        {
            get
            {
                if (sheets == null)
                {
                    sheets = new Dictionary<string, ExcelSheetSchema>();
                }

                return sheets;
            }
        }


        private int GetIndex(XmlSchemaAnnotation annotations,string nodeName)
        {
            if (annotations == null)
                throw new XmlSchemaException($"Schema node {nodeName} is missing required numeric Notes value.\nAll xml schema nodes must have Notes specified except root node. ");

            foreach (XmlSchemaAppInfo annotation in annotations.Items)
            {
                XmlNode node = annotation.Markup[0];
                //Is null if it does not exist
                //0-1 for multi cell span,***later
                XmlAttribute att = node.Attributes["notes"];

                if (att != null)
                {

                    if(att.Value.Length > 1)
                    {
                        return int.Parse(att.Value.Substring(0,2).TrimEnd());
                    }
                    else
                        return int.Parse(att.Value);
                }

            }
            

            return -1;
        }

        private XmlTypeCode GetExcelType(XmlTypeCode typeCode)
        {
            
            //DateTime
            //double
            //string
            //bool
            if (typeCode == XmlTypeCode.DateTime || typeCode == XmlTypeCode.Boolean || typeCode == XmlTypeCode.String || typeCode == XmlTypeCode.Double)
            {
                return typeCode;
            }
            if (typeCode == XmlTypeCode.Float || typeCode == XmlTypeCode.Decimal)
            {
                return XmlTypeCode.Double;
            }
            else
            {
                return XmlTypeCode.String;
            }
        }

        public void ResetWorkBookRows()
        {
            foreach (var sh in Sheets)
            {
                ExcelSheetSchema sheet = sh.Value;

                foreach (var rw in sheet.Rows)
                {
                    ExcelRowSchema row = rw.Value;
                    row.Processed = 0;
                }
            }
        }
    }
}
