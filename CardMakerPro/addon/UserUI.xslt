<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:frmwrk="Corel Framework Data">
 <xsl:output method="xml" encoding="UTF-8" indent="yes"/>
 <frmwrk:uiconfig>
  <frmwrk:applicationInfo userConfiguration="true"/>
  <frmwrk:compositeNode xPath="/uiConfig/frame"/>
 </frmwrk:uiconfig>
 <xsl:template match="node()|@*"><xsl:copy><xsl:apply-templates select="node()|@*"/></xsl:copy></xsl:template>
</xsl:stylesheet>
