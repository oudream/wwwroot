<?xml version='1.0'?>
<xsl:stylesheet xmlns:xsl="http://www.w3.org/TR/WD-xsl">
<xsl:template match="/">
<html>
<body>
<xsl:for-each select="PERSONNEL/PERSON">
<xsl:choose>
<xsl:when match=".[FG='boy']">
<font color="green"><li><xsl:value-of select="NAME"/></li></font>
<br/>
</xsl:when>
<xsl:otherwise match=".[FG='girl']">
<font color="red"><li><xsl:value-of select="NAME"/></li></font>
<br/>
</xsl:otherwise>
<xsl:otherwise>
<font color="blue"><hr><xsl:value-of select="NAME"/></hr></font>
</xsl:otherwise>
</xsl:choose>

</xsl:for-each>
</body>
</html>
</xsl:template>
</xsl:stylesheet>