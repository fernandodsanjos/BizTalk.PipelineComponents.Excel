namespace BizTalk.PipelineComponents.Excel.Schemas {
    using Microsoft.XLANGs.BaseTypes;
    
    
    [global::System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.BizTalk.Schema.Compiler", "3.0.1.0")]
    [global::System.Diagnostics.DebuggerNonUserCodeAttribute()]
    [global::System.Runtime.CompilerServices.CompilerGeneratedAttribute()]
    [SchemaType(SchemaTypeEnum.Document)]
    [Schema(@"http://BizTalk.PipelineComponents.Excel.Schemas.Order",@"Form")]
    [BodyXPath(@"/*[local-name()='Form' and namespace-uri()='http://BizTalk.PipelineComponents.Excel.Schemas.Order']/*[local-name()='Order' and namespace-uri()='']")]
    [System.SerializableAttribute()]
    [SchemaRoots(new string[] {@"Form"})]
    public sealed class Order : Microsoft.XLANGs.BaseTypes.SchemaBase {
        
        [System.NonSerializedAttribute()]
        private static object _rawSchema;
        
        [System.NonSerializedAttribute()]
        private const string _strSchema = @"<?xml version=""1.0"" encoding=""utf-16""?>
<xs:schema xmlns=""http://BizTalk.PipelineComponents.Excel.Schemas.Order"" xmlns:b=""http://schemas.microsoft.com/BizTalk/2003"" targetNamespace=""http://BizTalk.PipelineComponents.Excel.Schemas.Order"" xmlns:xs=""http://www.w3.org/2001/XMLSchema"">
  <xs:annotation>
    <xs:appinfo>
      <b:schemaInfo is_envelope=""yes"" xmlns:b=""http://schemas.microsoft.com/BizTalk/2003"" />
    </xs:appinfo>
  </xs:annotation>
  <xs:element name=""Form"">
    <xs:annotation>
      <xs:appinfo>
        <b:recordInfo body_xpath=""/*[local-name()='Form' and namespace-uri()='http://BizTalk.PipelineComponents.Excel.Schemas.Order']/*[local-name()='Order' and namespace-uri()='']"" />
      </xs:appinfo>
    </xs:annotation>
    <xs:complexType>
      <xs:sequence>
        <xs:element name=""Order"">
          <xs:annotation>
            <xs:appinfo>
              <b:recordInfo notes=""1"" xmlns:b=""http://schemas.microsoft.com/BizTalk/2003"" />
            </xs:appinfo>
          </xs:annotation>
          <xs:complexType>
            <xs:sequence>
              <xs:element maxOccurs=""unbounded"" name=""Group"">
                <xs:annotation>
                  <xs:appinfo>
                    <b:recordInfo notes=""6"" />
                  </xs:appinfo>
                </xs:annotation>
                <xs:complexType>
                  <xs:sequence>
                    <xs:element name=""OrderId"">
                      <xs:annotation>
                        <xs:appinfo>
                          <b:recordInfo notes=""2"" />
                        </xs:appinfo>
                      </xs:annotation>
                      <xs:complexType>
                        <xs:attribute name=""Id"" type=""xs:string"">
                          <xs:annotation>
                            <xs:appinfo>
                              <b:fieldInfo notes=""1"" />
                            </xs:appinfo>
                          </xs:annotation>
                        </xs:attribute>
                      </xs:complexType>
                    </xs:element>
                    <xs:element name=""OrderRow"">
                      <xs:annotation>
                        <xs:appinfo>
                          <b:recordInfo notes=""6"" />
                        </xs:appinfo>
                      </xs:annotation>
                      <xs:complexType>
                        <xs:sequence>
                          <xs:element name=""Article"" type=""xs:string"">
                            <xs:annotation>
                              <xs:appinfo>
                                <b:fieldInfo notes=""1"" />
                              </xs:appinfo>
                            </xs:annotation>
                          </xs:element>
                        </xs:sequence>
                        <xs:attribute name=""Environment"" type=""xs:boolean"">
                          <xs:annotation>
                            <xs:appinfo>
                              <b:fieldInfo notes=""3"" />
                            </xs:appinfo>
                          </xs:annotation>
                        </xs:attribute>
                        <xs:attribute name=""Qty"" type=""xs:double"">
                          <xs:annotation>
                            <xs:appinfo>
                              <b:fieldInfo notes=""4"" />
                            </xs:appinfo>
                          </xs:annotation>
                        </xs:attribute>
                        <xs:attribute name=""Price"" type=""xs:double"">
                          <xs:annotation>
                            <xs:appinfo>
                              <b:fieldInfo notes=""6"" />
                            </xs:appinfo>
                          </xs:annotation>
                        </xs:attribute>
                      </xs:complexType>
                    </xs:element>
                    <xs:element name=""OrderDate"">
                      <xs:annotation>
                        <xs:appinfo>
                          <b:recordInfo notes=""3"" />
                        </xs:appinfo>
                      </xs:annotation>
                      <xs:complexType>
                        <xs:sequence>
                          <xs:element name=""Date"" type=""xs:string"">
                            <xs:annotation>
                              <xs:appinfo>
                                <b:fieldInfo notes=""1"" />
                              </xs:appinfo>
                            </xs:annotation>
                          </xs:element>
                        </xs:sequence>
                      </xs:complexType>
                    </xs:element>
                    <xs:element name=""DeliveryDate"">
                      <xs:annotation>
                        <xs:appinfo>
                          <b:recordInfo notes=""4"" />
                        </xs:appinfo>
                      </xs:annotation>
                      <xs:complexType>
                        <xs:sequence>
                          <xs:element name=""Date"" type=""xs:string"">
                            <xs:annotation>
                              <xs:appinfo>
                                <b:fieldInfo notes=""1"" />
                              </xs:appinfo>
                            </xs:annotation>
                          </xs:element>
                        </xs:sequence>
                      </xs:complexType>
                    </xs:element>
                  </xs:sequence>
                </xs:complexType>
              </xs:element>
            </xs:sequence>
          </xs:complexType>
        </xs:element>
      </xs:sequence>
    </xs:complexType>
  </xs:element>
</xs:schema>";
        
        public Order() {
        }
        
        public override string XmlContent {
            get {
                return _strSchema;
            }
        }
        
        public override string[] RootNodes {
            get {
                string[] _RootElements = new string [1];
                _RootElements[0] = "Form";
                return _RootElements;
            }
        }
        
        protected override object RawSchema {
            get {
                return _rawSchema;
            }
            set {
                _rawSchema = value;
            }
        }
    }
}
