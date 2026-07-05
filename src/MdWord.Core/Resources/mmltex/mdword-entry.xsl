<?xml version="1.0" encoding="UTF-8"?>
<!--
  MD-Word's own entry point into the vendored xsltml 2.1.2 library
  (mmltex.xsl + its includes: tokens/glayout/scripts/tables/entities/cmarkup).
  This file is NOT part of the xsltml distribution: it is a thin wrapper
  we wrote that imports mmltex.xsl unmodified (exactly the usage the
  library's own README documents: "In your stylesheet import ... the main
  stylesheet, mmltex.xsl").

  Why this file exists: mmltex.xsl's own two top-level `m:math` templates
  wrap their output in delimiters we don't want: `$ ... $` for inline and
  `\n\[\n\t...\n\]` for display (confirmed empirically before wiring this
  up). MD-Word's own Markdown convention is `$...$` / `$$...$$`, decided in
  C# by the caller (which already knows whether it's converting an inline
  m:oMath or a block m:oMathPara), not by this stylesheet. Because this
  file uses xsl:import (not xsl:include), its templates have strictly
  higher import precedence than every template in mmltex.xsl and its own
  includes, so this override always wins regardless of the pattern
  specificity/priority of mmltex.xsl's own `m:math` rules: no need to
  touch the vendored files themselves.
-->
<xsl:stylesheet xmlns:xsl="http://www.w3.org/1999/XSL/Transform"
                xmlns:m="http://www.w3.org/1998/Math/MathML"
                version="1.0">

  <xsl:import href="mmltex.xsl"/>

  <xsl:output method="text" indent="no" encoding="UTF-8"/>

  <xsl:template match="m:math">
    <xsl:apply-templates/>
  </xsl:template>

</xsl:stylesheet>
