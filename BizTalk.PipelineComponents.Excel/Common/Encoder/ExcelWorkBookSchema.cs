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

        private IFormulaEvaluator FormulaEvaluator { get;  set; }

        private XmlSchema ResolvedSchema { get; set; }
        public ExcelWorkBookSchema(XmlSchema schema, IFormulaEvaluator formulaEvaluator)
        {
           
            FormulaEvaluator = formulaEvaluator;
            ResolvedSchema = schema; 
            //if (schema.Items.Count > 1)
            //   throw new ArgumentException("Source schema is only allowed to have one root node");

            XmlSchemaElement workbook = null;

            foreach (var item in schema.Items)
            {
                if (item is XmlSchemaElement)
                {
                    workbook = (XmlSchemaElement)item;
                    break;

                }

            }

            this.Namespace = workbook.QualifiedName.Namespace;
            this.Name = workbook.QualifiedName.Name;

            if (!(workbook.SchemaType is XmlSchemaComplexType))
                throw new ArgumentException("Nodes bellow root node must be of record type");

            //first must be complex type containing sequences of Sheet type objects
            XmlSchemaComplexType tp = (XmlSchemaComplexType)(workbook.SchemaType == null ? workbook.ElementSchemaType : workbook.SchemaType);

            XmlSchemaSequence seq = (XmlSchemaSequence)tp.Particle;

            XmlSchemaElement firstElement = (XmlSchemaElement)seq.Items[0];

            if (firstElement.MaxOccursString == "unbounded")
            {
                //Create virtual sheet
                ExcelSheetSchema eSheet = new ExcelSheetSchema
                {
                    Name = firstElement.QualifiedName.Name,
                    Namespace = firstElement.QualifiedName.Namespace,
                    Index = 0,
                    IsVirtual = true
                };

                this.Sheets.Add(eSheet.Name, eSheet);

                ParseRows(eSheet, seq);


            }
            else
            {

                for (int i = 0; i < seq.Items.Count; i++)
                {
                    XmlSchemaElement sheet = (XmlSchemaElement)seq.Items[i];

                    object sheetObject = (sheet.SchemaType == null ? sheet.ElementSchemaType : sheet.SchemaType);

                    if (!(sheetObject is XmlSchemaComplexType))
                        continue;

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

                    XmlSchemaComplexType sheetType = (XmlSchemaComplexType)sheetObject;



                    XmlSchemaSequence rowSequence = (XmlSchemaSequence)sheetType.Particle;


                    ParseRows(eSheet, rowSequence);


                }

            }

        }

        private void ParseRows(ExcelSheetSchema sheet, XmlSchemaSequence rowSequence)
        {
            for (int x = 0; x < rowSequence.Items.Count; x++)
            {
                XmlSchemaElement row = (XmlSchemaElement)rowSequence.Items[x];

                int occurrence = row.MaxOccursString == "unbounded" ? -1 : (int)row.MaxOccurs;

                if (row.Annotation == null)
                {
                    if(ResolvedSchema.Elements.Contains(row.QualifiedName))
                    {
                        row = (XmlSchemaElement)ResolvedSchema.Elements[row.QualifiedName];
                    }
                    
                }

                int rowIndex = GetIndex(row.Annotation, row.QualifiedName.Name);

                if (rowIndex == -1)
                    rowIndex = x;

                XmlSchemaComplexType rowType = (XmlSchemaComplexType)(row.SchemaType == null ? row.ElementSchemaType:row.SchemaType);

                ExcelRowSchema eRow = new ExcelRowSchema
                {
                    Name = row.QualifiedName.Name,
                    Namespace = row.QualifiedName.Namespace,
                    Index = rowIndex,
                    Occurrence = occurrence

                };

                if (rowType.Attributes.Count > 0)
                {

                    foreach (XmlSchemaAttribute item in rowType.Attributes)
                    {
                        XmlSchemaSimpleType sp = item.AttributeSchemaType;

                        //  XmlSchemaDatatype tp;
                        //  XmlTypeCode.DateTime
                        int cellIndex = GetIndex(item.Annotation, item.QualifiedName.Name);

                        if (cellIndex > -1)
                        {
                            eRow.Cells.Add(item.Name, new ExcelCellSchema
                            {
                                XmlType = GetExcelType(sp.Datatype.TypeCode),
                                Index = cellIndex,
                                Name = item.QualifiedName.Name,
                                NodeType = 'A',
                                FormulaEvaluator = FormulaEvaluator,
                                Parent = eRow

                            });
                        }


                    }


                }

                if (rowType.Particle != null)
                {
                    XmlSchemaSequence elementSequence = (XmlSchemaSequence)rowType.Particle;

                    foreach (XmlSchemaElement item in elementSequence.Items)
                    {



                        int cellIndex = GetIndex(item.Annotation, item.QualifiedName.Name);

                        if (cellIndex > -1)
                        {

                            eRow.Cells.Add(item.QualifiedName.Name, new ExcelCellSchema
                            {
                                XmlType = GetExcelType(item.ElementSchemaType.Datatype.TypeCode),
                                Index = cellIndex,
                                Name = item.QualifiedName.Name,
                                NodeType = 'E',
                                Parent = eRow
                            });

                        }


                    }
                }

                sheet.Rows.Add(row.QualifiedName.Name, eRow);
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
            if (typeCode == XmlTypeCode.DateTime || typeCode == XmlTypeCode.Date ||  typeCode == XmlTypeCode.Boolean || typeCode == XmlTypeCode.String || typeCode == XmlTypeCode.Double)
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
