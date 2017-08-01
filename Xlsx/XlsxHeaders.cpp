/*

Copyright (c) 2012-2013, Pavel Akimov
All rights reserved.

Redistribution and use in source and binary forms, with or without
modification, are permitted provided that the following conditions are met:
    * Redistributions of source code must retain the above copyright
      notice, this list of conditions and the following disclaimer.
    * Redistributions in binary form must reproduce the above copyright
      notice, this list of conditions and the following disclaimer in the
      documentation and/or other materials provided with the distribution.
    * Neither the name of the <organization> nor the
      names of its contributors may be used to endorse or promote products
      derived from this software without specific prior written permission.

THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND
ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED
WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
DISCLAIMED. IN NO EVENT SHALL <COPYRIGHT HOLDER> BE LIABLE FOR ANY
DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES
(INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND
ON ANY THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
(INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.

*/

#include "XlsxHeaders.h"

namespace SimpleXlsx
{
    const char * ns_content_types   = "http://schemas.openxmlformats.org/package/2006/content-types";
    const char * content_printer 	= "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings";
    const char * content_rels 		= "application/vnd.openxmlformats-package.relationships+xml";
    const char * content_xml 		= "application/xml";
    const char * content_book 		= "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    const char * content_sheet 		= "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    const char * content_theme 		= "application/vnd.openxmlformats-officedocument.theme+xml";
    const char * content_styles     = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml";
    const char * content_sharedStr 	= "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml";
    const char * content_comment    = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml";
    const char * content_vml        = "application/vnd.openxmlformats-officedocument.vmlDrawing";
    const char * content_core 		= "application/vnd.openxmlformats-package.core-properties+xml";
    const char * content_app 		= "application/vnd.openxmlformats-officedocument.extended-properties+xml";
    const char * content_chart      = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml";
    const char * content_drawing    = "application/vnd.openxmlformats-officedocument.drawing+xml";
    const char * content_chartsheet = "application/vnd.openxmlformats-officedocument.spreadsheetml.chartsheet+xml";
    const char * content_chain      = "application/vnd.openxmlformats-officedocument.spreadsheetml.calcChain+xml";

    const char * ns_markup_compatibility    = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const char * ns_c14                     = "http://schemas.microsoft.com/office/drawing/2007/8/2/chart";
    const char * ns_relationships_chart     = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    const char * ns_relationships	= "http://schemas.openxmlformats.org/package/2006/relationships";
    const char * type_book			= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";
    const char * type_core			= "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties";
    const char * type_app			= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";

    const char * ns_cp				= "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
    const char * ns_dc				= "http://purl.org/dc/elements/1.1/";
    const char * ns_dcterms			= "http://purl.org/dc/terms/";
    const char * ns_dcmitype        = "http://purl.org/dc/dcmitype/";
    const char * ns_xsi				= "http://www.w3.org/2001/XMLSchema-instance";
    const char * xsi_type			= "dcterms:W3CDTF";

    const char * ns_doc_prop		= "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties";
    const char * ns_vt				= "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes";

    const char * ns_book			= "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    const char * ns_book_r			= "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

    const char * ns_mc				= "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const char * ns_x14ac			= "http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac";

    const char * ns_x14				= "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";

    const char * type_comments		= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments";
    const char * type_vml			= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing";
    const char * type_sheet			= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet";
    const char * type_style			= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles";
    const char * type_sharedStr		= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings";
    const char * type_theme			= "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme";
    const char * type_chain         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/calcChain";
    const char * type_chartsheet    = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chartsheet";
    const char * type_chart         = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart";
    const char * type_drawing       = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing";

    const char * ns_a				= "http://schemas.openxmlformats.org/drawingml/2006/main";
    const char * ns_c               = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    const char * ns_xdr             = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
}
