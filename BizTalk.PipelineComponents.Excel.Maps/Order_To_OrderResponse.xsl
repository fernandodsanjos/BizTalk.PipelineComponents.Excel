<?xml version="1.0" encoding="UTF-16"?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:msxsl="urn:schemas-microsoft-com:xslt" xmlns:var="http://schemas.microsoft.com/BizTalk/2003/var" exclude-result-prefixes="msxsl var" version="1.0" xmlns:ns0="http://BizTalk.PipelineComponents.Excel.Schemas.Order">
  <xsl:output omit-xml-declaration="yes" method="xml" version="1.0" />
  <xsl:template match="/">
    <xsl:apply-templates select="/ns0:Form" />
  </xsl:template>
  <xsl:template match="/ns0:Form">
    <ns0:FormResponse>
      <Order>
        <OrderId>
          <xsl:if test="Order/OrderId/@Id">
            <xsl:attribute name="Id">
              <xsl:value-of select="Order/OrderId/@Id" />
            </xsl:attribute>
          </xsl:if>
          <xsl:value-of select="Order/OrderId/text()" />
        </OrderId>
        <xsl:for-each select="Order/OrderRow">
          <OrderRow>
            <xsl:if test="@Qty">
              <xsl:attribute name="Qty">
                <xsl:value-of select="@Qty" />
              </xsl:attribute>
            </xsl:if>
            <xsl:if test="@Price">
              <xsl:attribute name="Price">
                <xsl:value-of select="@Price" />
              </xsl:attribute>
            </xsl:if>
            <Article>
              <xsl:value-of select="Article/text()" />
            </Article>
          </OrderRow>
        </xsl:for-each>
        <OrderDate>
          <Date>
            <xsl:value-of select="Order/OrderDate/Date/text()" />
          </Date>
        </OrderDate>
        <DeliveryDate>
          <Date>
            <xsl:value-of select="Order/DeliveryDate/Date/text()" />
          </Date>
        </DeliveryDate>
        <Contact>
          <Value>
            <xsl:text>Fernando Pires</xsl:text>
          </Value>
        </Contact>
        <Contact>
          <Value>
            <xsl:text>08570349996</xsl:text>
          </Value>
        </Contact>
        <Contact>
          <Value>
            <xsl:text>snabel@peab.se</xsl:text>
          </Value>
        </Contact>
      </Order>
    </ns0:FormResponse>
  </xsl:template>
</xsl:stylesheet>