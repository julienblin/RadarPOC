<?xml version="1.0" encoding="utf-8" standalone="yes"?>
<!-- Replace your namespace as needed-->
<xsl:stylesheet version="1.0"
 xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
 xmlns:w="http://schemas.microsoft.com/office/word/2003/wordml"
 xmlns:v="urn:schemas-microsoft-com:vml" 
 xmlns:w10="urn:schemas-microsoft-com:office:word" 
 xmlns:sl="http://schemas.microsoft.com/schemaLibrary/2003/core" 
 xmlns:aml="http://schemas.microsoft.com/aml/2001/core" 
 xmlns:wx="http://schemas.microsoft.com/office/word/2003/auxHint" 
 xmlns:o="urn:schemas-microsoft-com:office:office" 
 xmlns:dt="uuid:C2F41010-65B3-11d1-A29F-00AA00C14882">

 <xsl:template  match="/">
  <!-- Basic conversion of an Infopath rich text node to WordML
      Author: Stephane Bouillon - Microsoft Services Belgium
      Feb 2006 

      This is a work in progress that will work for most hand-typed rich text values, but will almost certainly fail with
      cut/pasted html content, especially with nested tables and divs, for which I do not yet have a good solution.
      
      Send suggestions and feedback to Stephane.Bouillon@microsoft.com
 -->
  <!-- optional input parameters for default paragraph and character formatting -->
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <xsl:choose>
   <xsl:when test="descendant::div | descendant::ul | descendant::ol">
    <xsl:apply-templates select="* | text()">
     <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
     <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
    </xsl:apply-templates>
   </xsl:when>
   <xsl:otherwise>
    <w:p>
     <w:pPr>
      <xsl:copy-of select="$pPr_Default"/>
     </w:pPr>
     <w:r>
      <w:rPr>
       <xsl:call-template name="apply-nested-character-formatting">
        <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
        <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
       </xsl:call-template>
      </w:rPr>
      <w:t>
       <xsl:apply-templates select="* | text()">
        <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
        <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
       </xsl:apply-templates>
      </w:t>
     </w:r>
    </w:p>
   </xsl:otherwise>
  </xsl:choose>
 </xsl:template>

 <!-- XHTML div -->
 <xsl:template match="div">
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <w:p>
   <w:pPr>
    <xsl:call-template name="apply-paragraph-formatting">
     <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
    </xsl:call-template>
   </w:pPr>
   <w:r>
    <w:rPr>
     <xsl:call-template name="apply-nested-character-formatting">
      <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
     </xsl:call-template>
    </w:rPr>
    <w:t>
     <xsl:apply-templates select="* | text()">
      <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
      <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
     </xsl:apply-templates>
    </w:t>
   </w:r>
  </w:p>
 </xsl:template>

 <!-- XHTML table -->
 <xsl:template match="table">
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <w:tbl>
   <w:tblPr>
    <w:tblBorders>
     <w:top w:val="single"/>
     <w:left w:val="single"/>
     <w:bottom w:val="single"/>
     <w:right w:val="single"/>
     <w:insideH w:val="single"/>
     <w:insideV w:val="single"/>
    </w:tblBorders>
   </w:tblPr>
   <xsl:apply-templates select="*">
    <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
    <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
   </xsl:apply-templates>
  </w:tbl>
 </xsl:template>

 <xsl:template match="tbody">
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <xsl:apply-templates select="*">
   <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
   <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
  </xsl:apply-templates>
 </xsl:template>

 <xsl:template match="tr">
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <w:tr>
   <xsl:apply-templates select="*">
    <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
    <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
   </xsl:apply-templates>
  </w:tr>
 </xsl:template>

 <xsl:template match="td">
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <w:tc>
   <xsl:choose>
    <xsl:when test="descendant::div | descendant::ul | descendant::ol">
     <xsl:apply-templates select="* | text()">
      <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
      <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
     </xsl:apply-templates>
    </xsl:when>
    <xsl:otherwise>
     <w:p>
      <w:r>
       <w:rPr>
        <xsl:call-template name="apply-nested-character-formatting">
         <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
        </xsl:call-template>
       </w:rPr>
       <w:t>
        <xsl:apply-templates select="* | text()">
         <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
         <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
        </xsl:apply-templates>
       </w:t>
      </w:r>
     </w:p>
    </xsl:otherwise>
   </xsl:choose>
  </w:tc>
 </xsl:template>

 <xsl:template match="colgroup">
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <w:tblGrid>
   <xsl:apply-templates select="*">
    <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
    <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
   </xsl:apply-templates>
  </w:tblGrid>
 </xsl:template>

 <xsl:template match="col">
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <xsl:if test="contains(@style,'WIDTH: ')">
   <w:gridCol>
    <xsl:attribute name="w:w">
     <xsl:value-of select="substring-before(substring-after(@style,'WIDTH: '),'px')"/>
   </xsl:attribute>
   </w:gridCol>
  </xsl:if>
 </xsl:template>

 <!-- XHTML UL (bulleted list) -->
 <xsl:template match="ul">
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <xsl:param name="seqnr" select="1+count(preceding-sibling::ul)+count(preceding-sibling::ol)"/>
  <w:cfChunk>
   <w:lists>
    <w:listDef w:listDefId="0">
     <w:plt w:val="HybridMultilevel"/>
     <w:tmpl w:val="CE70554C"/>
     <w:lvl w:ilvl="0" w:tplc="04090001">
      <w:start w:val="1"/>
       <xsl:choose>
         <xsl:when test="@type='circle'">
           <w:nfc w:val="23"/>
           <w:lvlText w:val="o"/>
         </xsl:when>
         <xsl:when test="@type='square'">
           <w:nfc w:val="23"/>
           <w:lvlText w:val="¡"/>
         </xsl:when>
         <xsl:otherwise>
           <w:nfc w:val="23"/>
           <w:lvlText w:val="·"/>
         </xsl:otherwise>
       </xsl:choose>
       <w:lvlJc w:val="left"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="720"/>
       </w:tabs>
       <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
        <xsl:choose>
          <xsl:when test="@type='square'">
            <w:rFonts w:ascii="Wingdings 2" w:h-ansi="Wingdings 2" w:hint="default"/>
          </xsl:when>
          <xsl:otherwise>
            <w:rFonts w:ascii="Symbol" w:h-ansi="Symbol" w:hint="default"/>
          </xsl:otherwise>
        </xsl:choose>
      </w:rPr>
     </w:lvl>
     <w:lvl w:ilvl="1" w:tplc="04090003" w:tentative="on">
      <w:start w:val="1"/>
      <w:nfc w:val="23"/>
      <w:lvlText w:val="·"/>
      <!--w:lvlText w:val="o"/-->
      <w:lvlJc w:val="left"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="1440"/>
       </w:tabs>
       <w:ind w:left="1440" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
       <w:rFonts w:ascii="Symbol" w:h-ansi="Symbol" w:hint="default"/>
       <!--w:rFonts w:ascii="Courier New" w:h-ansi="Courier New" w:cs="Courier New" w:hint="default"/-->
      </w:rPr>
     </w:lvl>
     <w:lvl w:ilvl="2" w:tplc="04090005" w:tentative="on">
      <w:start w:val="1"/>
      <w:nfc w:val="23"/>
      <w:lvlText w:val="·"/>
      <!--w:lvlText w:val="§"/-->
      <w:lvlJc w:val="left"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="2160"/>
       </w:tabs>
       <w:ind w:left="2160" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
       <w:rFonts w:ascii="Symbol" w:h-ansi="Symbol" w:hint="default"/>
       <!--w:rFonts w:ascii="Wingdings" w:h-ansi="Wingdings" w:hint="default"/-->
      </w:rPr>
     </w:lvl>
     <w:lvl w:ilvl="3" w:tplc="04090001" w:tentative="on">
      <w:start w:val="1"/>
      <w:nfc w:val="23"/>
      <w:lvlText w:val="·"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="2880"/>
       </w:tabs>
       <w:ind w:left="2880" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
       <w:rFonts w:ascii="Symbol" w:h-ansi="Symbol" w:hint="default"/>
      </w:rPr>
     </w:lvl>
     <w:lvl w:ilvl="4" w:tplc="04090003" w:tentative="on">
      <w:start w:val="1"/>
      <w:nfc w:val="23"/>
      <w:lvlText w:val="·"/>
      <!--w:lvlText w:val="o"/-->
      <w:lvlJc w:val="left"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="3600"/>
       </w:tabs>
       <w:ind w:left="3600" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
       <w:rFonts w:ascii="Symbol" w:h-ansi="Symbol" w:hint="default"/>
       <!--w:rFonts w:ascii="Courier New" w:h-ansi="Courier New" w:cs="Courier New" w:hint="default"/-->
      </w:rPr>
     </w:lvl>
     <w:lvl w:ilvl="5" w:tplc="04090005" w:tentative="on">
      <w:start w:val="1"/>
      <w:nfc w:val="23"/>
      <w:lvlText w:val="·"/>
      <!--w:lvlText w:val="§"/-->
      <w:lvlJc w:val="left"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="4320"/>
       </w:tabs>
       <w:ind w:left="4320" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
       <w:rFonts w:ascii="Symbol" w:h-ansi="Symbol" w:hint="default"/>
       <!--w:rFonts w:ascii="Wingdings" w:h-ansi="Wingdings" w:hint="default"/-->
      </w:rPr>
     </w:lvl>
     <w:lvl w:ilvl="6" w:tplc="04090001" w:tentative="on">
      <w:start w:val="1"/>
      <w:nfc w:val="23"/>
      <w:lvlText w:val="·"/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="5040"/>
       </w:tabs>
       <w:ind w:left="5040" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
       <w:rFonts w:ascii="Symbol" w:h-ansi="Symbol" w:hint="default"/>
      </w:rPr>
     </w:lvl>
     <w:lvl w:ilvl="7" w:tplc="04090003" w:tentative="on">
      <w:start w:val="1"/>
      <w:nfc w:val="23"/>
      <w:lvlText w:val="·"/>
      <!--w:lvlText w:val="o"/-->
      <w:lvlJc w:val="left"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="5760"/>
       </w:tabs>
       <w:ind w:left="5760" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
       <w:rFonts w:ascii="Symbol" w:h-ansi="Symbol" w:hint="default"/>
       <!--w:rFonts w:ascii="Courier New" w:h-ansi="Courier New" w:cs="Courier New" w:hint="default"/-->
      </w:rPr>
     </w:lvl>
     <w:lvl w:ilvl="8" w:tplc="04090005" w:tentative="on">
      <w:start w:val="1"/>
      <w:nfc w:val="23"/>
      <w:lvlText w:val="·"/>
      <!--w:lvlText w:val="§"/-->
      <w:lvlJc w:val="left"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="6480"/>
       </w:tabs>
       <w:ind w:left="6480" w:hanging="360"/>
      </w:pPr>
      <w:rPr>
       <w:rFonts w:ascii="Symbol" w:h-ansi="Symbol" w:hint="default"/>
       <!--w:rFonts w:ascii="Wingdings" w:h-ansi="Wingdings" w:hint="default"/-->
      </w:rPr>
     </w:lvl>
    </w:listDef>
    <w:list>
     <xsl:attribute name="w:ilfo"><xsl:value-of select="$seqnr"/></xsl:attribute>
     <w:ilst w:val="0"/>
    </w:list>
   </w:lists>
   <xsl:for-each select="*">
    <xsl:choose>
     <xsl:when test="local-name()='li'">
      <w:p>
       <w:pPr>
        <xsl:call-template name="apply-paragraph-formatting">
         <xsl:with-param name="pPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
        </xsl:call-template>
        <w:listPr>
         <w:ilvl>
          <xsl:attribute name="w:val"><xsl:value-of select="count(ancestor::ul|ancestor::ol)-1"/></xsl:attribute>
         </w:ilvl>
         <w:ilfo>
          <xsl:attribute name="w:val"><xsl:value-of select="$seqnr"/></xsl:attribute>
         </w:ilfo>
        </w:listPr>
       </w:pPr>
       <w:r>
        <w:rPr>
         <xsl:call-template name="apply-nested-character-formatting">
          <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
         </xsl:call-template>
        </w:rPr>
        <w:t>
         <xsl:apply-templates select="* | text()">
          <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
          <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
         </xsl:apply-templates>
        </w:t>
       </w:r>
      </w:p>
     </xsl:when>
     <xsl:otherwise>
      <xsl:apply-templates select=".">
       <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
       <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
      </xsl:apply-templates>
     </xsl:otherwise>
    </xsl:choose>
   </xsl:for-each>
  </w:cfChunk>
 </xsl:template>

 <!-- XHTML OL (numbered list) -->
 <xsl:template match="ol">
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <xsl:param name="seqnr" select="1+count(preceding-sibling::ul)+count(preceding-sibling::ol)"/>
   <w:cfChunk>
   <w:lists>
    <w:listDef w:listDefId="0">
     <w:plt w:val="HybridMultilevel" />
     <w:tmpl w:val="0986AB3E" />
     <w:lvl w:ilvl="0" w:tplc="0409000F">
      <w:start w:val="1"/>
       <xsl:choose>
         <xsl:when test="@type='I'">
           <w:nfc w:val="1"/>
           <w:lvlText w:val="%1."/>
         </xsl:when>
         <xsl:when test="@type='i'">
           <w:nfc w:val="2"/>
           <w:lvlText w:val="%1."/>
         </xsl:when>
         <xsl:when test="@type='A'">
           <w:nfc w:val="3"/>
           <w:lvlText w:val="%1."/>
         </xsl:when>
         <xsl:when test="@type='a'">
           <w:nfc w:val="4"/>
           <w:lvlText w:val="%1."/>
         </xsl:when>
         <xsl:otherwise>
           <w:nfc w:val="0"/>
           <w:lvlText w:val="%1."/>
         </xsl:otherwise>
       </xsl:choose>
      <w:lvlJc w:val="left"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="720"/>
       </w:tabs>
       <w:ind w:left="720" w:hanging="360"/>
      </w:pPr>
     </w:lvl>
     <w:lvl w:ilvl="1" w:tplc="04090019" w:tentative="on">
      <w:start w:val="1"/>
      <!-- w:nfc w:val="4"/ -->
      <w:lvlText w:val="%2."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="1440"/>
       </w:tabs>
       <w:ind w:left="1440" w:hanging="360"/>
      </w:pPr>
     </w:lvl>
     <w:lvl w:ilvl="2" w:tplc="0409001B" w:tentative="on">
      <w:start w:val="1"/>
      <!-- w:nfc w:val="2"/ -->
      <w:lvlText w:val="%3."/>
      <w:lvlJc w:val="right"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="2160"/>
       </w:tabs>
       <w:ind w:left="2160" w:hanging="180"/>
      </w:pPr>
     </w:lvl>
     <w:lvl w:ilvl="3" w:tplc="0409000F" w:tentative="on">
      <w:start w:val="1"/>
      <w:lvlText w:val="%4."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="2880"/>
       </w:tabs>
       <w:ind w:left="2880" w:hanging="360"/>
      </w:pPr>
     </w:lvl>
     <w:lvl w:ilvl="4" w:tplc="04090019" w:tentative="on">
      <w:start w:val="1"/>
      <!-- w:nfc w:val="4"/-->
      <w:lvlText w:val="%5."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="3600"/>
       </w:tabs>
       <w:ind w:left="3600" w:hanging="360"/>
      </w:pPr>
     </w:lvl>
     <w:lvl w:ilvl="5" w:tplc="0409001B" w:tentative="on">
      <w:start w:val="1"/>
      <!--w:nfc w:val="2"/-->
      <w:lvlText w:val="%6."/>
      <w:lvlJc w:val="right"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="4320"/>
       </w:tabs>
       <w:ind w:left="4320" w:hanging="180"/>
      </w:pPr>
     </w:lvl>
     <w:lvl w:ilvl="6" w:tplc="0409000F" w:tentative="on">
      <w:start w:val="1"/>
      <w:lvlText w:val="%7."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="5040"/>
       </w:tabs>
       <w:ind w:left="5040" w:hanging="360"/>
      </w:pPr>
     </w:lvl>
     <w:lvl w:ilvl="7" w:tplc="04090019" w:tentative="on">
      <w:start w:val="1"/>
      <!--w:nfc w:val="4"/-->
      <w:lvlText w:val="%8."/>
      <w:lvlJc w:val="left"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="5760"/>
       </w:tabs>
       <w:ind w:left="5760" w:hanging="360"/>
      </w:pPr>
     </w:lvl>
     <w:lvl w:ilvl="8" w:tplc="0409001B" w:tentative="on">
      <w:start w:val="1"/>
      <!--w:nfc w:val="2"/-->
      <w:lvlText w:val="%9."/>
      <w:lvlJc w:val="right"/>
      <w:pPr>
       <w:tabs>
        <w:tab w:val="list" w:pos="6480"/>
       </w:tabs>
       <w:ind w:left="6480" w:hanging="180"/>
      </w:pPr>
     </w:lvl>
    </w:listDef>
    <w:list>
     <xsl:attribute name="w:ilfo"><xsl:value-of select="$seqnr"/></xsl:attribute>
     <w:ilst w:val="0" />
    </w:list>
   </w:lists>
   <w:body>
    <xsl:for-each select="*">
     <xsl:choose>
      <xsl:when test="local-name()='li'">
       <w:p>
        <w:pPr>
         <xsl:call-template name="apply-paragraph-formatting">
          <xsl:with-param name="pPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
         </xsl:call-template>
         <w:listPr>
          <w:ilvl>
           <xsl:attribute name="w:val"><xsl:value-of select="count(ancestor::ol|ancestor::ul)-1"/></xsl:attribute>
          </w:ilvl>
          <w:ilfo>
           <xsl:attribute name="w:val"><xsl:value-of select="$seqnr"/></xsl:attribute>
          </w:ilfo>
         </w:listPr>
        </w:pPr>
        <w:r>
         <w:rPr>
          <xsl:call-template name="apply-nested-character-formatting">
           <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
          </xsl:call-template>
         </w:rPr>
         <w:t>
          <xsl:apply-templates select="* | text()">
           <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
           <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
          </xsl:apply-templates>
         </w:t>
        </w:r>
       </w:p>
      </xsl:when>
      <xsl:otherwise>
       <xsl:apply-templates select=".">
        <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
        <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
       </xsl:apply-templates>
      </xsl:otherwise>
     </xsl:choose>
    </xsl:for-each>
   </w:body>
  </w:cfChunk>
 </xsl:template>

 <!-- XHTML img -->
 <xsl:template match="img">
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <xsl:text disable-output-escaping="yes">&#60;/w:t&#62;&#60;/w:r&#62;</xsl:text>
  <w:r>
   <w:pict>
    <w:binData>
     <xsl:attribute name="w:name">wordml://<xsl:value-of select="@src"/></xsl:attribute>
    </w:binData>
    <v:shape id="_x000_i1025" type="#_x0000_t75">
     <xsl:attribute name="style"><xsl:value-of select="@style"/></xsl:attribute>
     <v:imagedata>
      <xsl:attribute name="src">wordml://<xsl:value-of select="@src"/></xsl:attribute>
     </v:imagedata>
    </v:shape>
   </w:pict>
  </w:r>
  <xsl:text disable-output-escaping="yes">&#60;w:r&#62;</xsl:text>
  <w:rPr>
   <xsl:call-template name="apply-nested-character-formatting">
    <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
   </xsl:call-template>
  </w:rPr>
  <xsl:text disable-output-escaping="yes">&#60;w:t&#62;</xsl:text>
 </xsl:template>

 <!-- XHTML hyperlink -->
 <xsl:template match="a">
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <xsl:text disable-output-escaping="yes">&#60;/w:t&#62;&#60;/w:r&#62;</xsl:text>
  <w:hlink>
   <xsl:attribute name="w:dest"><xsl:value-of select="@href"/></xsl:attribute>
   <w:r>
    <w:rPr>
     <xsl:call-template name="apply-nested-character-formatting">
      <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
     </xsl:call-template>
     <w:color w:val="0000FF" />
     <w:u w:val="single" />
    </w:rPr>
    <w:t>
     <xsl:apply-templates>
      <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
      <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
     </xsl:apply-templates>
    </w:t>
   </w:r>
  </w:hlink>
  <xsl:text disable-output-escaping="yes">&#60;w:r&#62;</xsl:text>
  <w:rPr>
   <xsl:call-template name="apply-nested-character-formatting">
    <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
   </xsl:call-template>
  </w:rPr>
  <xsl:text disable-output-escaping="yes">&#60;w:t&#62;</xsl:text>
 </xsl:template>

 <!-- XHTML bold -->
 <xsl:template match="b | strong">
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <xsl:text disable-output-escaping="yes">&#60;/w:t&#62;&#60;/w:r&#62;</xsl:text>
  <w:r>
   <w:rPr>
    <xsl:call-template name="apply-nested-character-formatting">
     <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
    </xsl:call-template>
    <w:b/>
   </w:rPr>
   <w:t>
    <xsl:apply-templates>
     <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
     <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
    </xsl:apply-templates>
   </w:t>
  </w:r>
  <xsl:text disable-output-escaping="yes">&#60;w:r&#62;</xsl:text>
  <w:rPr>
   <xsl:call-template name="apply-nested-character-formatting">
    <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
   </xsl:call-template>
  </w:rPr>
  <xsl:text disable-output-escaping="yes">&#60;w:t&#62;</xsl:text>
 </xsl:template>

 <!-- XHTML italic -->
 <xsl:template match="i | em">
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <xsl:text disable-output-escaping="yes">&#60;/w:t&#62;&#60;/w:r&#62;</xsl:text>
  <w:r>
   <w:rPr>
    <xsl:call-template name="apply-nested-character-formatting">
     <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
    </xsl:call-template>
    <w:i/>
   </w:rPr>
   <w:t>
    <xsl:apply-templates>
     <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
     <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
    </xsl:apply-templates>
   </w:t>
  </w:r>
  <xsl:text disable-output-escaping="yes">&#60;w:r&#62;</xsl:text>
  <w:rPr>
   <xsl:call-template name="apply-nested-character-formatting">
    <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
   </xsl:call-template>
  </w:rPr>
  <xsl:text disable-output-escaping="yes">&#60;w:t&#62;</xsl:text>
 </xsl:template>

 <!-- XHTML underline -->
 <xsl:template match="u">
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <xsl:text disable-output-escaping="yes">&#60;/w:t&#62;&#60;/w:r&#62;</xsl:text>
  <w:r>
   <w:rPr>
    <xsl:call-template name="apply-nested-character-formatting">
     <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
    </xsl:call-template>
    <w:u w:val="single"/>
   </w:rPr>
   <w:t>
    <xsl:apply-templates>
     <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
     <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
    </xsl:apply-templates>
   </w:t>
  </w:r>
  <xsl:text disable-output-escaping="yes">&#60;w:r&#62;</xsl:text>
  <w:rPr>
   <xsl:call-template name="apply-nested-character-formatting">
    <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
   </xsl:call-template>
  </w:rPr>
  <xsl:text disable-output-escaping="yes">&#60;w:t&#62;</xsl:text>
 </xsl:template>

 <!-- XHTML font -->
 <xsl:template match="font">
  <xsl:param name="pPr_Default"/>
  <xsl:param name="rPr_Default"/>
  <xsl:text disable-output-escaping="yes">&#60;/w:t&#62;&#60;/w:r&#62;</xsl:text>
  <w:r>
   <w:rPr>
    <xsl:call-template name="apply-nested-character-formatting">
     <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
    </xsl:call-template>
    <xsl:if test="@color">
     <w:color>
      <xsl:attribute name="w:val"><xsl:value-of select="@color"/></xsl:attribute>
     </w:color>
    </xsl:if>
    <xsl:if test="@face">
     <w:rFonts>
      <xsl:attribute name="w:ascii"><xsl:value-of select="@face"/></xsl:attribute>
      <xsl:attribute name="w:h-ansi"><xsl:value-of select="@face"/></xsl:attribute>
     </w:rFonts>
    </xsl:if>
    <xsl:if test="@size">
     <w:sz>
      <xsl:attribute name="w:val">
       <xsl:choose>
        <xsl:when test="@size=1">16</xsl:when>
        <xsl:when test="@size=2">20</xsl:when>
        <xsl:when test="@size=3">24</xsl:when>
        <xsl:when test="@size=4">28</xsl:when>
        <xsl:when test="@size=5">36</xsl:when>
        <xsl:when test="@size=6">48</xsl:when>
        <xsl:when test="@size=7">72</xsl:when>
       </xsl:choose>
      </xsl:attribute>
     </w:sz>
    </xsl:if>
    <xsl:if test="not(@size)">
     <w:sz w:val="20"/>
    </xsl:if>
    <xsl:if test="@style">
     <xsl:call-template name="apply-font-style-definition">
      <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
      <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
      <xsl:with-param name="instyle"><xsl:value-of select="@style"/></xsl:with-param>
     </xsl:call-template>
    </xsl:if>
   </w:rPr>
   <w:t>
    <xsl:apply-templates>
     <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
     <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
    </xsl:apply-templates>
   </w:t>
  </w:r>
  <xsl:text disable-output-escaping="yes">&#60;w:r&#62;</xsl:text>
  <w:rPr>
   <xsl:call-template name="apply-nested-character-formatting">
    <xsl:with-param name="pPr_Default"><xsl:copy-of select="$pPr_Default"/></xsl:with-param>
    <xsl:with-param name="rPr_Default"><xsl:copy-of select="$rPr_Default"/></xsl:with-param>
   </xsl:call-template>
  </w:rPr>
  <xsl:text disable-output-escaping="yes">&#60;w:t&#62;</xsl:text>
 </xsl:template>

 <!-- XHTML nested character formatting -->
 <xsl:template name="apply-nested-character-formatting">
  <xsl:param name="rPr_Default"/>
  <!-- default font (when font face is not specified) is Verdana -->
  <xsl:if test="not(ancestor::font/@face)">
   <w:rFonts w:ascii="Verdana" w:h-ansi="Verdana"/>
  </xsl:if>
  <!-- default font size (when size is not specified) is 20 -->
  <xsl:if test="not(ancestor::font/@size)">
   <w:sz w:val="20"/>
  </xsl:if>
  <!-- apply default character formatting first -->
  <xsl:copy-of select="$rPr_Default"/>
  <xsl:if test="ancestor::i or ancestor::em">
   <w:i/>
  </xsl:if>
  <xsl:if test="ancestor::b or ancestor::strong">
   <w:b/>
  </xsl:if>
  <xsl:if test="ancestor::u">
   <w:u w:val="single"/>
  </xsl:if>
  <xsl:if test="ancestor::strike">
   <w:strike/>
  </xsl:if>
  <xsl:if test="ancestor::sup">
   <w:vertAlign w:val="superscript"/>
  </xsl:if>
  <xsl:if test="ancestor::sub">
   <w:vertAlign w:val="subscript"/>
  </xsl:if>
  <xsl:if test="ancestor::font/@color">
   <w:color>
    <xsl:attribute name="w:val"><xsl:value-of select="ancestor::font/@color"/></xsl:attribute>
   </w:color>
  </xsl:if>
  <xsl:if test="ancestor::font/@face">
   <w:rFonts>
    <xsl:attribute name="w:ascii"><xsl:value-of select="ancestor::font/@face"/></xsl:attribute>
    <xsl:attribute name="w:h-ansi"><xsl:value-of select="ancestor::font/@face"/></xsl:attribute>
   </w:rFonts>
  </xsl:if>
  <xsl:if test="ancestor::font/@size">
   <w:sz>
    <xsl:attribute name="w:val">
       <xsl:choose>
        <xsl:when test="ancestor::font/@size=1">16</xsl:when>
        <xsl:when test="ancestor::font/@size=2">20</xsl:when>
        <xsl:when test="ancestor::font/@size=3">24</xsl:when>
        <xsl:when test="ancestor::font/@size=4">28</xsl:when>
        <xsl:when test="ancestor::font/@size=5">36</xsl:when>
        <xsl:when test="ancestor::font/@size=6">48</xsl:when>
        <xsl:when test="ancestor::font/@size=7">72</xsl:when>
       </xsl:choose>
      </xsl:attribute>
   </w:sz>
  </xsl:if>
  <xsl:if test="ancestor::font/@style">
   <xsl:call-template name="apply-font-style-definition">
    <xsl:with-param name="instyle"><xsl:value-of select="ancestor::font/@style"/></xsl:with-param>
   </xsl:call-template>
  </xsl:if>
 </xsl:template>

 <!-- XHTML style formatting -->
 <xsl:template name="apply-font-style-definition">
  <xsl:param name="instyle" select="''"/>
  <xsl:if test="contains($instyle,'BACKGROUND-COLOR: #')">
   <w:shd w:val="clear" w:color="auto">
    <xsl:attribute name="w:fill">
     <xsl:choose>
      <xsl:when test="contains($instyle,';')">
       <xsl:value-of select="substring-before(substring-after($instyle,'BACKGROUND-COLOR: #'),';')"/>
      </xsl:when>
      <xsl:otherwise>
       <xsl:value-of select="substring-after($instyle,'BACKGROUND-COLOR: #')"/>
      </xsl:otherwise>
     </xsl:choose>
    </xsl:attribute>
   </w:shd>
  </xsl:if>
 </xsl:template>

 <!-- XHTML paragraph formatting -->
 <xsl:template name="apply-paragraph-formatting">
  <xsl:param name="pPr_Default"/>
  <!-- apply default paragraph formatting first -->
  <xsl:copy-of select="$pPr_Default"/>
  <xsl:if test="@align">
   <xsl:choose>
    <xsl:when test="@align='justify'">
     <w:jc w:val="both"/>
    </xsl:when>
    <xsl:otherwise>
     <w:jc>
      <xsl:attribute name="w:val"><xsl:value-of select="@align"/></xsl:attribute>
     </w:jc>
    </xsl:otherwise>
   </xsl:choose>
  </xsl:if>
 </xsl:template>
</xsl:stylesheet>