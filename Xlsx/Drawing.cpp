#include "Drawing.h"

#include "Chart.h"
#include "XlsxHeaders.h"

#include "../PathManager.hpp"
#include "../XMLWriter.hpp"


namespace SimpleXlsx
{
    CDrawing::CDrawing( size_t index, PathManager & pathmanager ) : m_index( index ), m_pathManager( pathmanager )
    {
    }

    CDrawing::~CDrawing()
    {
    }

    bool CDrawing::Save()
    {
        if( ! IsEmpty() )
        {
            SaveDrawingRels();
            SaveDrawing();
        }
        return true;
    }

    // [- /xl/drawings/_rels/drawingX.xml.rels
    void CDrawing::SaveDrawingRels()
    {
        _tstringstream FileName;
        FileName << _T( "/xl/drawings/_rels/drawing" ) << m_index << _T( ".xml.rels" );

        XMLWriter xmlw( m_pathManager.RegisterXML( FileName.str() ) );
        xmlw.Tag( "Relationships" ).Attr( "xmlns", ns_relationships );
        int rId = 1;
        for( std::vector<ChartInfo>::const_iterator it = m_charts.begin(); it != m_charts.end(); it++, rId++ )
        {
            _tstringstream Target, rIdStream;
            Target << "../charts/chart" << ( *it ).Chart->GetIndex() << ".xml";
            rIdStream << "rId" << rId;

            xmlw.TagL( "Relationship" ).Attr( "Id", rIdStream.str() ).Attr( "Type", type_chart ).Attr( "Target", Target.str() ).EndL();
        }
        xmlw.End( "Relationships" );
    }

    // [- /xl/drawings/drawingX.xml
    void CDrawing::SaveDrawing()
    {
        _tstringstream FileName;
        FileName << _T( "/xl/drawings/drawing" ) << m_index << _T( ".xml" );

        XMLWriter xmlw( m_pathManager.RegisterXML( FileName.str() ) );
        xmlw.Tag( "xdr:wsDr" ).Attr( "xmlns:xdr", ns_xdr ).Attr( "xmlns:a", ns_a );

        int rId = 1;
        for( std::vector<ChartInfo>::const_iterator it = m_charts.begin(); it != m_charts.end(); it++, rId++ )
        {
            switch( ( * it ).AType )
            {
                case ChartInfo::absoluteAnchor:
                {
                    xmlw.Tag( "xdr:absoluteAnchor" );
                    xmlw.TagL( "xdr:pos" ).Attr( "x", 0 ).Attr( "y", 0 ).EndL();
                    xmlw.TagL( "xdr:ext" ).Attr( "cx", 9312088 ).Attr( "cy", 6084794 ).EndL();
                    SaveChartSection( xmlw, ( * it ).Chart, rId );
                    xmlw.End( "xdr:absoluteAnchor" );
                    break;
                }
                case ChartInfo::twoCellAnchor:
                {
                    xmlw.Tag( "xdr:twoCellAnchor" );
                    SaveChartPoint( xmlw, "xdr:from", ( * it ).TopLeft );
                    SaveChartPoint( xmlw, "xdr:to", ( * it ).BottomRight );
                    SaveChartSection( xmlw, ( * it ).Chart, rId );
                    xmlw.End( "xdr:twoCellAnchor" );
                    break;
                }
            }
        }
        xmlw.End( "xdr:wsDr" );
    }

    void CDrawing::SaveChartSection( XMLWriter & xmlw, CChart * chart, int rId )
    {
        _tstringstream rIdStream;
        rIdStream << "rId" << rId;

        xmlw.Tag( "xdr:graphicFrame" ).Attr( "macro", "" ).Tag( "xdr:nvGraphicFramePr" );
        xmlw.TagL( "xdr:cNvPr" ).Attr( "id", rId ).Attr( "name", chart->GetTitle() ).EndL();
        xmlw.Tag( "xdr:cNvGraphicFramePr" ).TagL( "a:graphicFrameLocks" ).Attr( "noGrp", 1 ).EndL().End( "xdr:cNvGraphicFramePr" );
        xmlw.End( "xdr:nvGraphicFramePr" );

        xmlw.Tag( "xdr:xfrm" );
        xmlw.TagL( "a:off" ).Attr( "x", 0 ).Attr( "y", 0 ).EndL();
        xmlw.TagL( "a:ext" ).Attr( "cx", 0 ).Attr( "cy", 0 ).EndL();
        xmlw.End( "xdr:xfrm" );

        xmlw.Tag( "a:graphic" ).Tag( "a:graphicData" ).Attr( "uri", ns_c );
        xmlw.TagL( "c:chart" ).Attr( "xmlns:c", ns_c ).Attr( "xmlns:r", ns_book_r ).Attr( "r:id", rIdStream.str() ).EndL();
        xmlw.End( "a:graphicData" ).End( "a:graphic" );

        xmlw.End( "xdr:graphicFrame" );
        xmlw.TagL( "xdr:clientData" ).EndL();
    }

    void CDrawing::SaveChartPoint( XMLWriter & xmlw, const char * Tag, const ChartPoint & Point )
    {
        xmlw.Tag( Tag );
        xmlw.TagOnlyContent( "xdr:col", Point.col );
        xmlw.TagOnlyContent( "xdr:colOff", Point.colOff );
        xmlw.TagOnlyContent( "xdr:row", Point.row );
        xmlw.TagOnlyContent( "xdr:rowOff", Point.rowOff );
        xmlw.End( Tag );
    }
}
