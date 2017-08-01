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

#include <stdlib.h>
#include <limits>
#include <iomanip>

#include "Worksheet.h"
#include "XlsxHeaders.h"
#include "Drawing.h"

#include "../PathManager.hpp"

namespace SimpleXlsx
{
    // ****************************************************************************
    /// @brief      The class constructor
    /// @param      index index of a sheet to be created (for example, sheet1.xml)
    /// @return     no
    /// @note	    The constructor creates an instance with specified sheetX.xml filename
    ///             and without any frozen panes
    // ****************************************************************************
    CWorksheet::CWorksheet( size_t index, CDrawing & drawing, PathManager & pathmanager ) : m_pathManager( pathmanager ), m_Drawing( drawing )
    {
        std::vector<ColumnWidth> colWidths;
        Init( index, 0, 0, colWidths );
    }

    // ****************************************************************************
    /// @brief      The class constructor to create sheet with frozen pane
    /// @param      index index of a sheet to be created (for example, sheet1.xml)
    /// @param      width frozen pane with
    /// @param      height frozen pane height
    /// @return     no
    // ****************************************************************************
    CWorksheet::CWorksheet( size_t index, uint32_t width, uint32_t height,
                            CDrawing & drawing, PathManager & pathmanager ) : m_pathManager( pathmanager ), m_Drawing( drawing )
    {
        std::vector<ColumnWidth> colWidths;
        Init( index, width, height, colWidths );
    }

    // ****************************************************************************
    /// @brief      The class constructor
    /// @param      index index of a sheet to be created (for example, sheet1.xml)
    /// @param		colWidths list of pairs colNumber:Width
    /// @return     no
    /// @note	    The constructor creates an instance with specified sheetX.xml filename
    ///             and without any frozen panes
    // ****************************************************************************
    CWorksheet::CWorksheet( size_t index, const std::vector<ColumnWidth> & colWidths,
                            CDrawing & drawing, PathManager & pathmanager ) : m_pathManager( pathmanager ), m_Drawing( drawing )
    {
        Init( index, 0, 0, colWidths );
    }

    // ****************************************************************************
    /// @brief      The class constructor to create sheet with frozen pane
    /// @param      index index of a sheet to be created (for example, sheet1.xml)
    /// @param      width frozen pane with
    /// @param      height frozen pane height
    /// @param		colWidths list of pairs colNumber:Width
    /// @return     no
    // ****************************************************************************
    CWorksheet::CWorksheet( size_t index, uint32_t width, uint32_t height, const std::vector<ColumnWidth> & colWidths,
                            CDrawing & drawing, PathManager & pathmanager ) : m_pathManager( pathmanager ), m_Drawing( drawing )
    {
        Init( index, width, height, colWidths );
    }

    // ****************************************************************************
    /// @brief  The class destructor (virtual)
    /// @return no
    // ****************************************************************************
    CWorksheet::~CWorksheet()
    {
        if( m_XMLWriter != NULL ) delete m_XMLWriter;
    }

    // ****************************************************************************
    /// @brief  Initializes internal variables and creates a header of a sheet xml tree
    /// @param  index index of a sheet to be created (for example, sheet1.xml)
    /// @param  frozenWidth frozen pane with
    /// @param  frozenHeight frozen pane height
    /// @param	colWidths list of pairs colNumber:Width
    /// @return no
    // ****************************************************************************
    void CWorksheet::Init( size_t index, uint32_t frozenWidth, uint32_t frozenHeight, const std::vector<ColumnWidth> & colWidths )
    {
        m_isOk = true;
        m_row_opened = false;
        m_current_column = 0;
        m_offset_column = 0;
        m_index = index;
        m_title = _T( "Sheet 1" );
        m_withFormula = false;
        m_withComments = false;
        m_calcChain.clear();
        m_sharedStrings = NULL;
        m_comments = NULL;
        m_mergedCells.clear();
        m_row_index = 0;
        m_page_orientation = PAGE_PORTRAIT;

        _tstringstream FileName;
        FileName << _T( "/xl/worksheets/sheet" ) << m_index << _T( ".xml" );
        m_XMLWriter = new XMLWriter( m_pathManager.RegisterXML( FileName.str() ) );
        if( ( m_XMLWriter == NULL ) || ! m_XMLWriter->IsOk() )
        {
            m_isOk = false;
            return;
        }

        m_XMLWriter->Tag( "worksheet" ).Attr( "xmlns", ns_book ).Attr( "xmlns:r", ns_book_r ).Attr( "xmlns:mc", ns_mc ).Attr( "mc:Ignorable", "x14ac" ).Attr( "xmlns:x14ac", ns_x14ac );
        m_XMLWriter->TagL( "dimension" ).Attr( "ref", "A1" ).EndL();
        m_XMLWriter->Tag( "sheetViews" ).Tag( "sheetView" ).Attr( "tabSelected", 0 ).Attr( "workbookViewId", 0 );
        if( frozenWidth != 0 || frozenHeight != 0 )
            AddFrozenPane( frozenWidth, frozenHeight );
        m_XMLWriter->End( "sheetView" ).End( "sheetViews" );

        m_XMLWriter->TagL( "sheetFormatPr" ).Attr( "defaultRowHeight", 15 ).Attr( "x14ac:dyDescent", 0.25 ).EndL();
        if( ! colWidths.empty() )
        {
            m_XMLWriter->Tag( "cols" );
            for( std::vector<ColumnWidth>::const_iterator it = colWidths.begin(); it != colWidths.end(); it++ )
                m_XMLWriter->TagL( "col" ).Attr( "min", it->colFrom + 1 ).Attr( "max", it->colTo + 1 ).Attr( "width", it->width ).EndL();
            m_XMLWriter->End( "cols" );
        }
        m_XMLWriter->Tag( "sheetData" );    // open sheetData tag
    }

    // ****************************************************************************
    /// @brief	Add string-formatted cell with specified style
    /// @param	data reference to data
    /// @return	no
    // ****************************************************************************
    void CWorksheet::AddCell( const std::string & value, size_t style_id )
    {
        if( value != "" )
        {
            const std::string & szCoord = CellCoord( m_row_index, m_offset_column + m_current_column ).ToString();
            m_XMLWriter->Tag( "c" ).Attr( "r", szCoord );

            if( style_id != 0 ) m_XMLWriter->Attr( "s", style_id );  // default style is not necessary to sign explisitly

            if( value[0] == '=' )
            {
                m_XMLWriter->TagOnlyContent( "f", value.substr( 1, value.length() - 1 ) );

                m_withFormula = true;
                m_calcChain.push_back( szCoord );
            }
            else
            {
                if( m_sharedStrings != NULL )
                {
                    uint64_t str_index = 0;
                    std::map<std::string, uint64_t>::iterator it = m_sharedStrings->find( value );
                    if( it == m_sharedStrings->end() )
                    {
                        str_index = m_sharedStrings->size();
                        ( *m_sharedStrings )[ value ] = str_index;
                    }
                    else
                    {
                        str_index = it->second;
                    }
                    m_XMLWriter->Attr( "t", "s" ).TagOnlyContent( "v", str_index );
                }
                else
                {
                    m_XMLWriter->TagOnlyContent( "v", value );
                }
            }
            m_XMLWriter->End( "c" );
        }

        m_current_column++;
    }

    // ****************************************************************************
    /// @brief	Add time-formatted cell with specified style
    /// @param	data reference to data
    /// @return	no
    // ****************************************************************************
    void CWorksheet::AddCell( time_t value, size_t style_id )
    {
        const int64_t secondsFrom1900to1970 = 2208988800u;
        const double excelOneSecond = 0.0000115740740740741;

        struct tm * t = localtime( & value );

        time_t timeSinceEpoch = t->tm_sec + t->tm_min * 60 + t->tm_hour * 3600 + t->tm_yday * 86400 +
                                ( t->tm_year - 70 ) * 31536000 + ( ( t->tm_year - 69 ) / 4 ) * 86400 -
                                ( ( t->tm_year - 1 ) / 100 ) * 86400 + ( ( t->tm_year + 299 ) / 400 ) * 86400;

        double CalcedValue = excelOneSecond * ( secondsFrom1900to1970 + timeSinceEpoch ) + 2;

        std::streamsize OldPrecision = m_XMLWriter->SetFloatPrecision( std::numeric_limits<double>::digits10 );
        AddCellRoutine( CalcedValue, style_id );
        m_XMLWriter->SetFloatPrecision( OldPrecision );
    }

    // ****************************************************************************
    /// @brief  Internal initializatino method adds frozen pane`s information into sheet
    /// @param  width frozen pane width (in number of cells)
    /// @param  height frozen pane height (in number of cells)
    /// @return no
    // ****************************************************************************
    void CWorksheet::AddFrozenPane( uint32_t width, uint32_t height )
    {
        m_XMLWriter->Tag( "pane" );
        if( width != 0 ) m_XMLWriter->Attr( "xSplit", width );
        if( height != 0 ) m_XMLWriter->Attr( "ySplit", height );

        const std::string & szCoord = CellCoord( height + 1, width ).ToString();
        m_XMLWriter->Attr( "topLeftCell", szCoord );

        const char * ActivePane = "activePane";
        if( ( width != 0 ) && ( height != 0 ) ) m_XMLWriter->Attr( ActivePane, "bottomRight" );
        else if( ( width == 0 ) && ( height != 0 ) ) m_XMLWriter->Attr( ActivePane, "bottomLeft" );
        else if( ( width != 0 ) && ( height == 0 ) ) m_XMLWriter->Attr( ActivePane, "topRight" );

        m_XMLWriter->Attr( "state", "frozen" ).End( "pane" );

        if( ( width > 0 ) && ( height > 0 ) )
        {
            const std::string & szCoordBL = CellCoord( height + 1, width - 1 ).ToString();
            const std::string & szCoordTR = CellCoord( height, width ).ToString();
            m_XMLWriter->TagL( "selection" ).Attr( "pane", "topRight" ).Attr( "activeCell", szCoordTR ).Attr( "sqref", szCoordTR ).EndL();
            m_XMLWriter->TagL( "selection" ).Attr( "pane", "bottomLeft" ).Attr( "activeCell", szCoordBL ).Attr( "sqref", szCoordBL ).EndL();
            m_XMLWriter->TagL( "selection" ).Attr( "pane", "bottomRight" ).Attr( "activeCell", szCoord ).Attr( "sqref", szCoord ).EndL();
        }
        else if( width > 0 )
        {
            m_XMLWriter->TagL( "selection" ).Attr( "pane", "topRight" ).Attr( "activeCell", szCoord ).Attr( "sqref", szCoord ).EndL();
        }
        else if( height > 0 )
            m_XMLWriter->TagL( "selection" ).Attr( "pane", "bottomLeft" ).Attr( "activeCell", szCoord ).Attr( "sqref", szCoord ).EndL();
    }

    // ****************************************************************************
    /// @brief  Appends merged cells range into the sheet
    /// @param  cellFrom (row value from 1, col value from 0)
    /// @param  cellTo (row value from 1, col value from 0)
    /// @return no
    // ****************************************************************************
    void CWorksheet::MergeCells( CellCoord cellFrom, CellCoord cellTo )
    {
        if( ( cellFrom.row == 0 ) || ( cellTo.row == 0 ) ) return;
        m_mergedCells.push_back( cellFrom.ToString() + ':' + cellTo.ToString() );
    }

    // ****************************************************************************
    ///	@brief	Receives next to write cell`s coordinates
    /// @param	currCell (row value from 1, col value from 0)
    /// @return	no
    // ****************************************************************************
    void CWorksheet::GetCurrentCellCoord( CellCoord & currCell ) const
    {
        currCell.row = m_row_index;
        currCell.col = m_current_column;
    }

    // ****************************************************************************
    /// @brief  Saves current xml document into a file with preset name
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool CWorksheet::Save()
    {
        m_XMLWriter->End( "sheetData" );    // close sheetData tag

        if( ! m_mergedCells.empty() )
        {
            m_XMLWriter->Tag( "mergeCells" ).Attr( "count", m_mergedCells.size() );
            for( std::list<std::string>::const_iterator it = m_mergedCells.begin(); it != m_mergedCells.end(); it++ )
                m_XMLWriter->TagL( "mergeCell" ).Attr( "ref", * it ).EndL();
            m_XMLWriter->End( "mergeCells" );
        }
        std::string sOrient;
        if( m_page_orientation == PAGE_PORTRAIT ) sOrient = "portrait";
        else if( m_page_orientation == PAGE_LANDSCAPE ) sOrient = "landscape";

        m_XMLWriter->TagL( "pageMargins" ).Attr( "left", 0.7 ).Attr( "right", 0.7 ).Attr( "top", 0.75 );
        /*                  */m_XMLWriter->Attr( "bottom", 0.75 ).Attr( "header", 0.3 ).Attr( "footer", 0.3 ).EndL();
        m_XMLWriter->TagL( "pageSetup" ).Attr( "paperSize", 9 ).Attr( "orientation", sOrient ).EndL();  // A4 paper size

        size_t rId = 1;
        if( m_withComments )
        {
            m_XMLWriter->TagL( "legacyDrawing" ).Attr( "r:id", "rId1" ).EndL();
            rId += 2;
        }

        if( ! m_Drawing.IsEmpty() )
        {
            _tstringstream rIdStream;
            rIdStream << "rId" << m_Drawing.GetIndex();
            m_XMLWriter->TagL( "drawing" ).Attr( "r:id", rIdStream.str() ).EndL();
            rId++;
        }

        m_XMLWriter->End( "worksheet" );

        // by deleting the stream the end of file writes and closes the stream
        delete m_XMLWriter;
        m_XMLWriter = NULL;

        if( ( rId != 1 ) && ! SaveSheetRels() ) return false;

        m_isOk = false;
        return true;
    }

    // ****************************************************************************
    /// @brief  Saves current sheet relations file
    /// @return no
    // ****************************************************************************
    bool CWorksheet::SaveSheetRels()
    {
        // [- zip/xl/worksheets/_rels/sheetN.xml.rels
        _tstringstream FileName;
        FileName << _T( "/xl/worksheets/_rels/sheet" ) << m_index << _T( ".xml.rels" );

        XMLWriter xmlw( m_pathManager.RegisterXML( FileName.str() ) );
        xmlw.Tag( "Relationships" ).Attr( "xmlns", ns_relationships );
        size_t rId = 1;
        if( m_withComments )
        {
            _tstringstream Vml, Comments;
            Vml << "../drawings/vmlDrawing" << m_index << ".vml";
            Comments << "../comments" << m_index << ".xml";
            xmlw.TagL( "Relationship" ).Attr( "Id", "rId1" ).Attr( "Type", type_vml ).Attr( "Target", Vml.str() ).EndL();
            xmlw.TagL( "Relationship" ).Attr( "Id", "rId2" ).Attr( "Type", type_comments ).Attr( "Target", Comments.str() ).EndL();
            rId += 2;
        }
        if( ! m_Drawing.IsEmpty() )
        {
            _tstringstream rIdStream, Drawing;
            rIdStream << "rId" << m_Drawing.GetIndex();
            Drawing << "../drawings/drawing" << m_index << ".xml";
            xmlw.TagL( "Relationship" ).Attr( "Id", rIdStream.str() ).Attr( "Type", type_drawing ).Attr( "Target", Drawing.str() ).EndL();
        }

        xmlw.End( "Relationships" );

        // zip/xl/worksheets/_rels/sheetN.xml.rels -]
        return true;
    }

}	// namespace SimpleXlsx
