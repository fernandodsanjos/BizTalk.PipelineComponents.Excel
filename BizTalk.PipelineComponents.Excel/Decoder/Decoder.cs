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
using Microsoft.BizTalk.Component.Utilities;
using System.Xml.Schema;
using System.Reflection;
using Microsoft.XLANGs.BaseTypes;
using System.Runtime.Remoting;
using System.Collections.Concurrent;
using System.Collections.Generic;
using BizTalk.PipelineComponents.Excel.Common.Decoder;

namespace BizTalk.PipelineComponents.Excel
{

    [ComponentCategory(CategoryTypes.CATID_PipelineComponent)]
    [System.Runtime.InteropServices.Guid("b46b052c-c482-4434-9a6e-d49654852340")]
    [ComponentCategory(CategoryTypes.CATID_Decoder)]
    public partial class Decoder : IComponent, IBaseComponent, IComponentUI
    {
        object schemalock = new object();
        static ConcurrentDictionary<string, ExcelWorkBookSchema> CachedSchemas = new ConcurrentDictionary<string, ExcelWorkBookSchema>();
        
        private ExcelWorkBookSchema WorkBookSchema { get; set; }

        private IWorkbook WorkBook { get; set; }

        #region Name & Description

        public string Name
        {
            get
            {
                return "Excel Decoder";

            }
        }

        public string Version { get { return "1.0"; } }


        public string Description
        {

            get
            {
                return "Excel (xls(x)) to XML";

            }
        }
        #endregion

        #region Properties
        [Description("Document xml schema")]
        public Schema DocumentSpecName
        {
            get;
            set;
        }
        #endregion

        
        public IBaseMessage Execute(IPipelineContext pContext, IBaseMessage pInMsg)
        {
            
            if (DocumentSpecName == null)
                throw new ArgumentException("DocSpec must be specified");
           
            this.WorkBook = WorkbookFactory.Create(pInMsg.BodyPart.Data,ImportOption.SheetContentOnly);

            IFormulaEvaluator formulaEvaluator = null;

            if (this.WorkBook is XSSFWorkbook)
                formulaEvaluator = new XSSFFormulaEvaluator(this.WorkBook);
            else
                formulaEvaluator = new HSSFFormulaEvaluator(this.WorkBook);


            this.WorkBookSchema  = GetWorkBookSchema(pContext, formulaEvaluator);

            VirtualStream outSstm = ProcessWorkbook();

            pInMsg.BodyPart.Data = outSstm;

            pInMsg.Context.Promote(new ContextProperty(SystemProperties.MessageType),$"{this.WorkBookSchema.Namespace}#{this.WorkBookSchema.Name}");
            pInMsg.Context.Write(new ContextProperty(SystemProperties.SchemaStrongName), DocumentSpecName.SchemaName);

            pContext.ResourceTracker.AddResource(outSstm);

            return pInMsg;
        }
    
        private VirtualStream ProcessWorkbook()
        {
            VirtualStream outStm = new VirtualStream();

            using (XmlWriter wtr = XmlWriter.Create(outStm, new XmlWriterSettings { Indent = false, CloseOutput = false }))
            {

                wtr.WriteStartElement("w", this.WorkBookSchema.Name, this.WorkBookSchema.Namespace);

                    foreach (KeyValuePair<int,ExcelSheetSchema> sheet in this.WorkBookSchema.Sheets)
                    {
                        ExcelSheetSchema eSchema = sheet.Value;
                        eSchema.Process(wtr, this.WorkBook.GetSheetAt(eSchema.Index));
                    }

                wtr.WriteEndElement();
                wtr.Flush();
            }

            outStm.Position = 0;

            return outStm;
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
