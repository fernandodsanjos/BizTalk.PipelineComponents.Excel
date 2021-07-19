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


namespace BizTalk.PipelineComponents.Excel.Common.Decoder
{
    public class ExcelWorkBookSchema
    {
        
        private Dictionary<int, ExcelSheetSchema> sheets;

        public string Name { get; private set; }
        public string Namespace { get; private set; }
        public bool Envelope { get; private set; }

        public ExcelWorkBookSchema(XmlSchema schema, IFormulaEvaluator formulaEvaluator)
        {
            bool rootFound = false;

            XmlSchemaElement workbook = null;

            foreach (var item in schema.Items)
            {
                if(item is XmlSchemaAnnotation)
                {
                    //Check to see if its an envelope schema
                    //In that case all rows in a sheet is expected to grouped, only Sheet is the allowed to be unbounded 
                    Envelope = IsEnvelope((XmlSchemaAnnotation)item);
                }
                else
                {
                    workbook = (XmlSchemaElement)item;

                    if(rootFound)
                        throw new ArgumentException("Source schema is only allowed to have one root node");

                    rootFound = true;

                }
            }

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

                if(sheet.Name == null)
                {
                    sheet = (XmlSchemaElement)schema.Elements[sheet.RefName];
                }

                int sheetIndex = GetIndex(sheet.Annotation, sheet.Name);

                if (sheetIndex == -1)
                    sheetIndex = i;
               
                ExcelSheetSchema eSheet = new ExcelSheetSchema
                {
                    Name = sheet.Name,
                    Namespace = sheet.QualifiedName.Namespace,
                    Index = sheetIndex,
                    IsEnvelope = Envelope
                };

                this.Sheets.Add(i, eSheet);
                
               

                XmlSchemaComplexType sheetType = (XmlSchemaComplexType)sheet.SchemaType;

                XmlSchemaSequence sheetSequence = (XmlSchemaSequence)sheetType.Particle;

                for (int x = 0; x < sheetSequence.Items.Count; x++)
                {
                    XmlSchemaElement row = (XmlSchemaElement)sheetSequence.Items[x];

                    int rowIndex = GetIndex(row.Annotation, row.Name);

                    if (rowIndex == -1)
                        rowIndex = x;

                    XmlSchemaComplexType rowType = (XmlSchemaComplexType)row.SchemaType;

                    int occurrence = row.MaxOccursString == "unbounded" ? -1 : (int)row.MaxOccurs;//Only last Row is allowed to have unbounded

                    ExcelRowSchema eRow = new ExcelRowSchema
                    {
                        Name = row.Name,
                        Namespace = row.QualifiedName.Namespace,
                        Index = rowIndex,
                        Occurrence = occurrence
                    };

                    if(row.MinOccurs == 0 && row.MaxOccurs == 1)//Last row with this configuration is expected to be an repetative record
                    {
                        eSheet.EnvelopeRow = rowIndex;
                    }

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
                                eRow.Cells.Add(cellIndex, new ExcelCellSchema
                                {
                                    XmlType = sp.Datatype.TypeCode,
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
                                
                                  eRow.Cells.Add(cellIndex, new ExcelCellSchema
                                  {
                                      XmlType = item.ElementSchemaType.TypeCode,
                                      Index = cellIndex,
                                      Name = item.Name,
                                      NodeType = 'E'
                                  });

                            }


                        }
                    }

                    eSheet.Rows.Add(x, eRow);
                }



            }

          

        }

        public Dictionary<int, ExcelSheetSchema> Sheets
        {
            get
            {
                if (sheets == null)
                {
                    sheets = new Dictionary<int, ExcelSheetSchema>();
                }

                return sheets;
            }
        }

        private bool IsEnvelope(XmlSchemaAnnotation annotations)
        {

            if (annotations == null)
                return false;

            foreach (XmlSchemaAppInfo annotation in annotations.Items)
            {
                XmlNode node = annotation.Markup[0];
                //Is null if it does not exist
                //0-1 for multi cell span,***later
                XmlAttribute att = node.Attributes["is_envelope"];

                if (att != null)
                {

                    if (att.Value.Length > 1)
                    {
                        return (att.Value == "yes");
                    }
                 
                }

            }


            return false;
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

       
    }
}
