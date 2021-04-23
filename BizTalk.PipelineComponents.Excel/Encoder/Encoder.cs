using System;
using System.ComponentModel;
using System.ComponentModel.DataAnnotations;
using System.IO;
using System.Xml;
using System.Xml.XPath;
using Microsoft.BizTalk.Component.Interop;
using Microsoft.BizTalk.Message.Interop;
using Microsoft.BizTalk.Streaming;
using IComponent = Microsoft.BizTalk.Component.Interop.IComponent;
using BizTalkComponents.Utils;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.SS.Converter;
using BizTalk.PipelineComponents.Excel.Common;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Xml.Schema;
using System.Runtime.Remoting;
using System.Reflection;
using Microsoft.BizTalk.Component.Utilities;
using BizTalk.PipelineComponents.Excel.Common.Encoder;
namespace BizTalk.PipelineComponents.Excel
{

    [ComponentCategory(CategoryTypes.CATID_PipelineComponent)]
    [System.Runtime.InteropServices.Guid("b46b052c-c482-4434-9a6e-d49654852339")]
    [ComponentCategory(CategoryTypes.CATID_Encoder)]
    public partial class Encoder : IComponent, IBaseComponent, IComponentUI
    {
        object schemalock = new object();
        static ConcurrentDictionary<string, ExcelWorkBookSchema> CachedSchemas = new ConcurrentDictionary<string, ExcelWorkBookSchema>();

        private ExcelWorkBookSchema WorkBookSchema { get; set; }

        private IWorkbook WorkBook { get; set; }

        #region Name & Description
        //error is added so one does not forget to change Name and Description

        public string Name
        {

            get
            {
                return "Excel Encoder";

            }
        }

        public string Version { get { return "1.0"; } }


        public string Description
        {
            get
            {
                return "XML to Excel (xls(x))";

            }
        }
        #endregion

        #region Properties
        [DisplayName("XLS Output")]
        [Description("Output *.xls instead of *.xslx")]
        public bool XLSOutput { get; set; } = false;

        [DisplayName("Excel template file")]
        [Description("Path to *.xls or *.xslx file, to be used as base file")]
        [RequiredRuntime]
        public string ExcelTemplate { get; set; }

        [Description("Document xml schema")]
        public Schema DocumentSpecName
        {
            get;
            set;
        }

        #endregion
        public IBaseMessage Execute(IPipelineContext pContext, IBaseMessage pInMsg)
        {
            VirtualStream outStream = new VirtualStream();

            if (ExcelTemplate == null)
                throw new ArgumentException("Excel template file path must be specified", "Excel template file");

            if (DocumentSpecName == null)
                throw new ArgumentException("Document schema must be specified", "DocumentSpecName");

            using (FileStream input = new FileStream(ExcelTemplate, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                try
                {
                    this.WorkBook = WorkbookFactory.Create(input); 

                    IFormulaEvaluator formulaEvaluator = null;

                    if (this.WorkBook is XSSFWorkbook)
                        formulaEvaluator = new XSSFFormulaEvaluator(this.WorkBook);
                    else
                        formulaEvaluator = new HSSFFormulaEvaluator(this.WorkBook);


                    this.WorkBookSchema = GetWorkBookSchema(pContext, formulaEvaluator);

                    ProcessWorkbook(pInMsg.BodyPart.GetOriginalDataStream());

                    this.WorkBookSchema.ResetWorkBookRows();

                    this.WorkBook.Write(outStream);
                }
                finally
                {
                    if(this.WorkBook != null)
                        this.WorkBook.Close();
                }
                
            }

            outStream.Position = 0;

            pContext.ResourceTracker.AddResource(outStream);

            pInMsg.BodyPart.Data = outStream;



            return pInMsg;
        }

        private void ProcessWorkbook(Stream body)
        {

            using (XmlReader reader = XmlReader.Create(body))
            {

                while (reader.Read())
                {
                    if(reader.IsStartElement() && reader.Depth == 1)
                    {

                        var excelScheetSchema = this.WorkBookSchema.Sheets[reader.LocalName];
                        excelScheetSchema.Process(reader.ReadSubtree(), this.WorkBook.GetSheetAt(excelScheetSchema.Index));

                    }
                }

            }

            

           
        }

        private ExcelWorkBookSchema GetWorkBookSchema(IPipelineContext pContext, IFormulaEvaluator formulaEvaluator)
        {

            if (String.IsNullOrEmpty(DocumentSpecName.AssemblyName))
            {
                string assemblyName = pContext.PipelineName.Substring(pContext.PipelineName.IndexOf(",") + 1).TrimStart();

                DocumentSpecName = new Schema(Assembly.CreateQualifiedName(assemblyName, DocumentSpecName.DocSpecName));
            }

            if (CachedSchemas.ContainsKey(DocumentSpecName.SchemaName))
                return CachedSchemas[DocumentSpecName.SchemaName];


            lock (schemalock)
            {
                if (CachedSchemas.ContainsKey(DocumentSpecName.SchemaName))
                    return CachedSchemas[DocumentSpecName.SchemaName];


                ObjectHandle result = Activator.CreateInstance(DocumentSpecName.AssemblyName, DocumentSpecName.DocSpecName);

                Microsoft.XLANGs.BaseTypes.SchemaBase schemaBase = (Microsoft.XLANGs.BaseTypes.SchemaBase)result.Unwrap();

                XmlSchema schema = schemaBase.CreateResolvedSchema();


                CachedSchemas.TryAdd(DocumentSpecName.SchemaName, new ExcelWorkBookSchema(schema, formulaEvaluator));

            }

            return CachedSchemas[DocumentSpecName.SchemaName];


        }


    }
}
