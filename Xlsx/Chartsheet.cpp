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

#include "Chartsheet.h"

#include "Chart.h"
#include "Drawing.h"
#include "SimpleXlsxDef.h"
#include "Worksheet.h"
#include "XlsxHeaders.h"

#include "../PathManager.hpp"
#include "../XMLWriter.hpp"

namespace SimpleXlsx
{
    // ****************************************************************************
    /// @brief  The class constructor
    /// @param  index index of a sheet to be created (for example, chart1.xml)
    /// @param	chart is a reference for correspond chart
    /// @param	drawing is a references for correspond drawing
    /// @param  Parent is a pointer to parent CWorkbook
    /// @return no
    // ****************************************************************************
    CChartsheet::CChartsheet( size_t index, CChart & chart, CDrawing & drawing, PathManager & pathmanager ) :
        m_index( index ), m_Chart( chart ), m_Drawing( drawing ), m_pathManager( pathmanager )
    {
    }

    // ****************************************************************************
    /// @brief  The class destructor (virtual)
    /// @return no
    // ****************************************************************************
    CChartsheet::~CChartsheet()
    {
    }

    // ****************************************************************************
    /// @brief  For each chart sheet it creates and saves 2 files:
    ///         sheetXX.xml, sheetXX.xml.rels,
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool CChartsheet::Save()
    {
        {
            // [- /xl/chartsheets/_rels/sheetX.xml.rels
            _tstringstream FileName;
            FileName << _T( "/xl/chartsheets/_rels/sheet" ) << m_index << _T( ".xml.rels" );

            _tstringstream Target;
            Target << "../drawings/drawing" << m_Drawing.GetIndex() << ".xml";

            XMLWriter xmlw( m_pathManager.RegisterXML( FileName.str() ) );
            xmlw.Tag( "Relationships" ).Attr( "xmlns", ns_relationships );
            xmlw.TagL( "Relationship" ).Attr( "Id", "rId1" ).Attr( "Type", type_drawing ).Attr( "Target", Target.str() ).EndL();

            xmlw.End( "Relationships" );
            // /xl/chartsheets/_rels/sheetX.xml.rels -]
        }

        {
            // [- /xl/chartsheets/sheetX.xml
            _tstringstream FileName;
            FileName << _T( "/xl/chartsheets/sheet" ) << m_index << _T( ".xml" );

            XMLWriter xmlw( m_pathManager.RegisterXML( FileName.str() ) );
            xmlw.Tag( "chartsheet" ).Attr( "xmlns", ns_book ).Attr( "xmlns:r", ns_book_r ).TagL( "sheetPr" ).EndL();

            xmlw.Tag( "sheetViews" );
            xmlw.TagL( "sheetView" ).Attr( "zoomScale", 85 ).Attr( "workbookViewId", 0 ).Attr( "zoomToFit", 1 ).EndL();
            xmlw.End( "sheetViews" );

            xmlw.TagL( "pageMargins" ).Attr( "left", 0.7 ).Attr( "right", 0.7 ).Attr( "top", 0.75 );
            xmlw.Attr( "bottom", 0.75 ).Attr( "header", 0.3 ).Attr( "footer", 0.3 ).EndL();
            xmlw.TagL( "drawing" ).Attr( "r:id", "rId1" ).EndL();

            xmlw.End( "chartsheet" );
            // /xl/chartsheets/sheetX.xml -]
        }
        return true;
    }

}	// namespace SimpleXlsx
