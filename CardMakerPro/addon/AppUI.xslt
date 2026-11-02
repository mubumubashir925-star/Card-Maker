<?xml version="1.0" encoding="UTF-8"?>
<xsl:stylesheet version="1.0" xmlns:xsl="http://www.w3.org/1999/XSL/Transform" xmlns:frmwrk="Corel Framework Data">
 <xsl:output method="xml" encoding="UTF-8" indent="yes"/>
 <frmwrk:uiconfig><frmwrk:applicationInfo userConfiguration="true"/></frmwrk:uiconfig>
 <xsl:template match="node()|@*"><xsl:copy><xsl:apply-templates select="node()|@*"/></xsl:copy></xsl:template>
 <xsl:template match="uiConfig/items">
  <xsl:copy><xsl:apply-templates select="node()|@*"/>
   <itemData guid="b19217a7-81d3-498b-9c9a-0a9340e2f209"
    type="wpfhost" hostedType="Addons\CardMakerPro\CardMakerPro.UI.dll,CardMakerPro.UI.Workspace" enable="true"/>
  </xsl:copy>
 </xsl:template>
 <xsl:template match="uiConfig/dockers">
  <xsl:copy><xsl:apply-templates select="node()|@*"/>
   <dockerData guid="F46B2D41-831D-47C8-95C5-B4CC7C0A01A2" userCaption="Card Maker Pro" wantReturn="true" focusStyle="noThrow" width="390">
    <container><item dock="fill" margin="0,0,0,0" guidRef="b19217a7-81d3-498b-9c9a-0a9340e2f209"/></container>
   </dockerData>
  </xsl:copy>
 </xsl:template>
</xsl:stylesheet>
