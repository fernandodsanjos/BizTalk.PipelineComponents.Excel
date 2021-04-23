using System;
using System.Collections;
using System.Linq;
using BizTalkComponents.Utils;
using Microsoft.BizTalk.Component.Interop;
using Microsoft.BizTalk.Message.Interop;
using Microsoft.BizTalk.Component.Utilities;

namespace BizTalk.PipelineComponents.Excel
{
    public partial class Decoder : IPersistPropertyBag
    {


        private string documentSpecName = null;

        public void GetClassID(out Guid classID)
        {
            classID = new Guid("b46b052c-c482-4434-9a6e-d49654852340");
        }

        public void InitNew()
        {

        }

        public IEnumerator Validate(object projectSystem)
        {
            return ValidationHelper.Validate(this, false).ToArray().GetEnumerator();
        }

        public bool Validate(out string errorMessage)
        {
            var errors = ValidationHelper.Validate(this, true).ToArray();

            if (errors.Any())
            {
                errorMessage = string.Join(",", errors);

                return false;
            }

            errorMessage = string.Empty;

            return true;
        }

        public IntPtr Icon { get { return IntPtr.Zero; } }

        //Load and Save are generic, the functions create properties based on the components "public" "read/write" properties.
        public void Load(IPropertyBag propertyBag, int errorLog)
        {

            documentSpecName = BizTalkComponents.Utils.PropertyBagHelper.ReadPropertyBag<string>(propertyBag, "DocumentSpecName", DocumentSpecName?.SchemaName);

            if (documentSpecName != null && documentSpecName.Length > 0)
            {

                DocumentSpecName = new Schema(documentSpecName);
                
            }

        }

        public void Save(IPropertyBag propertyBag, bool clearDirty, bool saveAllProperties)
        {
          
            if(DocumentSpecName != null)
                BizTalkComponents.Utils.PropertyBagHelper.WritePropertyBag(propertyBag, "DocumentSpecName", DocumentSpecName.SchemaName);

        }
    }
}
