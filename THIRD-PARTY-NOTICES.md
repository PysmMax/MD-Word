# Third-party notices

MD-Word is licensed under the [MIT License](LICENSE). It distributes or
embeds the following third-party components. Version numbers are the ones
shipped with this release; see each project for current information.

At runtime MD-Word also uses `MML2OMML.XSL` / `OMML2MML.XSL` **from the end
user's own Microsoft Office installation** to convert between MathML and
OMML. These files are not part of MD-Word and are never redistributed; if
they are absent, formulas degrade to plain text.

---

## Markdig (NuGet `Markdig` 1.3.2)

BSD 2-Clause License — Copyright (c) Alexandre Mutel. All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:

1. Redistributions of source code must retain the above copyright notice,
   this list of conditions and the following disclaimer.
2. Redistributions in binary form must reproduce the above copyright notice,
   this list of conditions and the following disclaimer in the documentation
   and/or other materials provided with the distribution.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
POSSIBILITY OF SUCH DAMAGE.

## Jint (NuGet `Jint` 4.10.1)

BSD 2-Clause License — Copyright (c) Sebastien Ros. All rights reserved.

(Same BSD 2-Clause terms and disclaimer as reproduced above for Markdig.)

## Open XML SDK (NuGet `DocumentFormat.OpenXml` 3.5.1)

MIT License — Copyright (c) Microsoft Corporation. All rights reserved.

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in
all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
IN THE SOFTWARE.

## KaTeX 0.16.47 (embedded `katex.min.js`)

MIT License — Copyright (c) 2013-2020 Khan Academy and other contributors.

(Same MIT terms and disclaimer as reproduced above for the Open XML SDK.)

## xsltml 2.1.2 (embedded `mmltex` XSLT stylesheets)

Copyright (C) 2001-2003 Vasil Yaroshevich. Permissive MIT-style license;
the original distribution's notice is preserved verbatim in
`src/MdWord.Core/Resources/mmltex/README-xsltml.txt`, which is embedded and
shipped unchanged alongside the stylesheets.

## Microsoft.Office.Interop.Word (NuGet 15.0.4797.1004)

Interop **type definitions are embedded** into `MdWord.AddIn.dll` at compile
time (`EmbedInteropTypes`); the primary interop assembly itself is not
redistributed.
