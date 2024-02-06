/*
  SimpleXlsxWriter
  Copyright (C) 2012-2021 Pavel Akimov <oxod.pavel@gmail.com>, Alexandr Belyak <programmeralex@bk.ru>

  This software is provided 'as-is', without any express or implied
  warranty. In no event will the authors be held liable for any damages
  arising from the use of this software.

  Permission is granted to anyone to use this software for any purpose,
  including commercial applications, and to alter it and redistribute it
  freely, subject to the following restrictions:

  1. The origin of this software must not be misrepresented; you must not
     claim that you wrote the original software. If you use this software
     in a product, an acknowledgment in the product documentation would be
     appreciated but is not required.
  2. Altered source versions must be plainly marked as such, and must not be
     misrepresented as being the original software.
  3. This notice may not be removed or altered from any source distribution.
*/

#include <cmath>

#include "Chart.h"

#include "Worksheet.h"
#include "XlsxHeaders.h"

#include "../PathManager.hpp"
#include "../XMLWriter.hpp"

namespace SimpleXlsx
{
// ****************************************************************************
/// @brief  The class constructor
/// @param  index index of a sheet to be created (for example, chart1.xml)
/// @param  type chart type (BAR or LINEAR)
/// @param	Parent parent CWorkbook
/// @return no
// ****************************************************************************
CChart::CChart( size_t index, EChartTypes type, PathManager & pathmanager ) :  m_index( index ), m_pathManager( pathmanager )
{
    m_title = "Diagramm 1";
    m_seriesSet.clear();

    m_diagramm.typeMain = type;

    m_xAxis.id = 100;
    m_yAxis.id = 101;
    m_x2Axis.id = 102;
    m_y2Axis.id = 103;

    m_xAxis.pos = POS_BOTTOM;
    m_yAxis.pos = POS_LEFT;
    m_x2Axis.pos = POS_TOP;
    m_x2Axis.cross = CROSS_MAX;
    m_y2Axis.pos = POS_RIGHT;
    m_y2Axis.cross = CROSS_MAX;
}

// ****************************************************************************
/// @brief  The class destructor (virtual)
/// @return no
// ****************************************************************************
CChart::~CChart()
{
}

// ****************************************************************************
/// @brief  For each chart it creates and saves 5 files:
///         chartXX.xml,
///         sheetXX.xml, sheetXX.xml.rels,
///         drawingXX.xml, drawingXX.xml.rels
/// @return Boolean result of the operation
// ****************************************************************************
bool CChart::Save()
{
    // [- /xl/charts/chartX.xml
    if( m_seriesSet.empty() ) return false;

    std::stringstream FileName;
    FileName << "/xl/charts/chart" << m_index << ".xml";
    XMLWriter xmlw( m_pathManager.RegisterXML( FileName.str() ) );

    xmlw.Tag( "c:chartSpace" ).Attr( "xmlns:c", ns_c ).Attr( "xmlns:a", ns_a ).Attr( "xmlns:r", ns_relationships_chart );
    xmlw.TagL( "c:date1904" ).Attr( "val", 0 ).EndL();
    xmlw.TagL( "c:lang" ).Attr( "val", "en-US" ).EndL();
    xmlw.TagL( "c:roundedCorners" ).Attr( "val", 0 ).EndL();

    xmlw.Tag( "mc:AlternateContent" ).Attr( "xmlns:mc", ns_markup_compatibility );
    xmlw.Tag( "mc:Choice" ).Attr( "Requires", "c14" ).Attr( "xmlns:c14", ns_c14 );
    xmlw.TagL( "c14:style" ).Attr( "val", 102 ).EndL().End( "mc:Choice" );
    xmlw.Tag( "mc:Fallback" ).TagL( "c:style" ).Attr( "val", 2 ).EndL().End( "mc:Fallback" );
    xmlw.End( "mc:AlternateContent" );

    xmlw.Tag( "c:chart" );
    if( ! m_diagramm.name.empty() )
        AddTitle( xmlw, m_diagramm.name, m_diagramm.nameSize, false );

    xmlw.TagL( "c:autoTitleDeleted" ).Attr( "val", 0 ).EndL();
    xmlw.Tag( "c:plotArea" ).TagL( "c:layout" ).EndL();


    switch( m_diagramm.typeMain )
    {
        case CHART_LINEAR   :
        {
            AddLineChart( xmlw, m_xAxis, m_yAxis.id, m_diagramm, m_seriesSet, 0 );
            break;
        }
        case CHART_BAR      :
        {
            AddBarChart( xmlw, m_xAxis, m_yAxis.id, m_diagramm, m_seriesSet, 0, m_diagramm.barDir, m_diagramm.barGroup );
            break;
        }
        case CHART_SCATTER  :
        {
            AddScatterChart( xmlw, m_xAxis.id, m_yAxis.id, m_diagramm, m_seriesSet, 0, m_diagramm.scatterStyle );
            break;
        }
        case CHART_PIE      :
        {
            AddPieChart( xmlw, m_diagramm, m_seriesSet, 0 );
            break;
        }
        case CHART_NONE     : return false;
            /*default:
                return false;*/
    }

    if( ! m_seriesSetAdd.empty() )
    {
        switch( m_diagramm.typeAdditional )
        {
            case CHART_LINEAR   :
            {
                AddLineChart( xmlw, m_x2Axis, m_y2Axis.id, m_diagramm, m_seriesSetAdd, m_seriesSet.size() );
                break;
            }
            case CHART_BAR      :
            {
                AddBarChart( xmlw, m_x2Axis, m_y2Axis.id, m_diagramm, m_seriesSetAdd, m_seriesSet.size(), m_diagramm.barDir, m_diagramm.barGroup );
                break;
            }
            case CHART_SCATTER  :
            {
                AddScatterChart( xmlw, m_x2Axis.id, m_y2Axis.id, m_diagramm, m_seriesSetAdd, m_seriesSet.size(), m_diagramm.scatterStyle );
                break;
            }
            case CHART_PIE      :
            {
                AddPieChart( xmlw, m_diagramm, m_seriesSetAdd, m_seriesSet.size() );
                break;
            }
            case CHART_NONE     : break;
                /*default:
                    break;*/
        }
    }

    if( m_diagramm.typeMain != CHART_PIE )
    {
        // Main axes
        AddXAxis( xmlw, m_xAxis, m_yAxis.id );
        AddYAxis( xmlw, m_yAxis, m_xAxis.id );
        // Additional axes
        if( ! m_seriesSetAdd.empty() && ( m_diagramm.typeAdditional != CHART_NONE ) )
        {
            AddXAxis( xmlw, m_x2Axis, m_y2Axis.id );
            AddYAxis( xmlw, m_y2Axis, m_x2Axis.id );
        }
    }
    AddTableData( xmlw, m_diagramm.tableData );

    AddAreaFill( xmlw, m_diagramm.plotAreaFill );

    xmlw.End( "c:plotArea" );
    AddLegend( xmlw, m_diagramm.legend_pos );
    xmlw.TagL( "c:plotVisOnly" ).Attr( "val", m_diagramm.showDataFromHiddenCells ? 0 : 1 ).EndL();
    switch( m_diagramm.emptyCellsDisplayMethod )
    {
        case EMPTY_CELLS_DISP_GAPS:
            xmlw.TagL( "c:dispBlanksAs" ).Attr( "val", "gap" ).EndL();
            break;
        case EMPTY_CELLS_DISP_ZERO:
            xmlw.TagL( "c:dispBlanksAs" ).Attr( "val", "zero" ).EndL();
            break;
        case EMPTY_CELLS_DISP_CONNECT:
            xmlw.TagL( "c:dispBlanksAs" ).Attr( "val", "span" ).EndL();
            break;
    }
    xmlw.TagL( "c:showDLblsOverMax" ).Attr( "val", 0 ).EndL();
    xmlw.End( "c:chart" );
    AddAreaFill( xmlw, m_diagramm.chartAreaFill );
    xmlw.End( "c:chartSpace" );

    // /xl/charts/chartX.xml -]
    return true;
}

// ****************************************************************************
/// @brief  Adds specified series into chartsheet
/// @param  series reference to series object to be added
/// @param  mainChart indicates whether series must be added into main chart or additional (main by default)
/// @return Boolean result of the oepration
// ****************************************************************************
bool CChart::AddSeries( const CChart::Series & series, bool mainChart )
{
    if( ( series.valSheet == NULL ) || ( ( series.valAxisTo.row == 0 ) && ( series.valAxisTo.col == 0 ) ) )
    {
        return false;
    }

    if( mainChart )
    {
        if( ( m_diagramm.typeMain == CHART_SCATTER ) && ( ( series.catSheet == NULL ) ||
                ( ( series.catAxisTo.row == 0 ) && ( series.catAxisTo.col == 0 ) ) ) )
        {
            return false;
        }

        m_seriesSet.push_back( series );
    }
    else
    {
        if( ( m_diagramm.typeAdditional == CHART_SCATTER ) && ( ( series.catSheet == NULL ) ||
                ( ( series.catAxisTo.row == 0 ) && ( series.catAxisTo.col == 0 ) ) ) )
        {
            return false;
        }

        m_seriesSetAdd.push_back( series );
    }

    return true;
}

// ****************************************************************************
/// @brief  This internal method adds title section into xml object
/// @param  name    Title value
/// @param  size    Font size
/// @param  vertPos indicates whether the title is allocated on X axis (false) or vertical (true)
/// @return no
// ****************************************************************************
void CChart::AddTitle( XMLWriter & xmlw, const UniString & name, uint32_t size, bool vertPos )
{
    xmlw.Tag( "c:title" ).Tag( "c:tx" ).Tag( "c:rich" ).Tag( "a:bodyPr" );
    if( vertPos )
    {
        xmlw.Attr( "rot", -5400000 ).Attr( "vert", "horz" );

    } // else
    // xmlw.Attr( "rot", 0 ).Attr( "vert", "horz" ) ;
    xmlw.End( "a:bodyPr" ).TagL( "a:lstStyle" ).EndL();

    xmlw.Tag( "a:p" ).Tag( "a:pPr" );
    xmlw.TagL( "a:defRPr" ).Attr( "lang", "en-US" ).Attr( "sz", size * 100 ).EndL().End( "a:pPr" );

    xmlw.Tag( "a:r" );
    xmlw.TagL( "a:rPr" ).Attr( "lang", "en-US" ).Attr( "sz", size * 100 ).EndL();
    xmlw.TagOnlyContent( "a:t", name ).End( "a:r" );

    xmlw.TagL( "a:endParaRPr" ).Attr( "lang", "en-US" ).EndL();
    xmlw.End( "a:p" ).End( "c:rich" ).End( "c:tx" );

    xmlw.TagL( "c:layout" ).EndL();
    xmlw.TagL( "c:overlay" ).Attr( "val", 0 ).EndL();
    xmlw.End( "c:title" );
}

// ****************************************************************************
/// @brief  Adds table data into chart if it is necessary
/// @return no
// ****************************************************************************
void CChart::AddTableData( XMLWriter & xmlw, ETableData tableData )
{
    bool showKeys = false;

    switch( tableData )
    {
        case TBL_DATA_NONE      :   return;
        case TBL_DATA           :   break;
        case TBL_DATA_N_KEYS    :   showKeys = true; break;
            //                default:                break;
    }
    xmlw.Tag( "c:dTable" );
    xmlw.TagL( "c:showHorzBorder" ).Attr( "val", 1 ).EndL();
    xmlw.TagL( "c:showVertBorder" ).Attr( "val", 1 ).EndL();
    xmlw.TagL( "c:showOutline" ).Attr( "val", 1 ).EndL();
    if( showKeys ) xmlw.TagL( "c:showKeys" ).Attr( "val", 1 ).EndL();
    xmlw.End( "c:dTable" );
}

// ****************************************************************************
/// @brief  Adds legend into chart if it is necessary
/// @return no
// ****************************************************************************
void CChart::AddLegend( XMLWriter & xmlw, EPosition legend_pos )
{
    int overlay = 1;
    char pos = '\0';
    switch( legend_pos )
    {
        case POS_NONE:  return;

        case POS_LEFT_ASIDE     :   overlay = 0; pos = 'l'; break;
        case POS_LEFT           :   pos = 'l'; break;

        case POS_RIGHT_ASIDE    :   overlay = 0; pos = 'r'; break;
        case POS_RIGHT          :   pos = 'r'; break;

        case POS_TOP_ASIDE      :   overlay = 0; pos = 't'; break;
        case POS_TOP            :   pos = 't'; break;

        case POS_BOTTOM_ASIDE   :   overlay = 0; pos = 'b'; break;
        case POS_BOTTOM         :   pos = 'b'; break;
    }

    xmlw.Tag( "c:legend" );
    xmlw.TagL( "c:legendPos" ).Attr( "val", pos ).EndL();
    xmlw.TagL( "c:layout" ).EndL();
    xmlw.TagL( "c:overlay" ).Attr( "val", overlay ).EndL();
    xmlw.End( "c:legend" );
}

static const char * AxisTypeToTag( CChart::EAxisType Type )
{
    switch( Type )
    {
        case CChart::AXIS_VALUE     : return "c:valAx";
        case CChart::AXIS_DATE      : return "c:dateAx";
        case CChart::AXIS_CATEGORY  : return "c:catAx";
        default                     : return "c:valAx";
    }
}

static const char * TickMarkToTag( CChart::ETickMark Mark )
{
    switch( Mark )
    {
        case CChart::MARK_NONE      : return "none";
        case CChart::MARK_CROSS     : return "cross";
        case CChart::MARK_INSIDE    : return "in";
        case CChart::MARK_OUTSIDE   : return "out";
        default                     : return "none";
    }
}

static const char * TimeUnitToTag( CChart::ETimeUnit Unit )
{
    switch( Unit )
    {
        case CChart::TIME_UNIT_DAYS     : return "days";
        case CChart::TIME_UNIT_MONTHS   : return "months";
        case CChart::TIME_UNIT_YEARS    : return "years";
        default                         : return "days";
    }
}

static void AddTickMarks( XMLWriter & xmlw, const CChart::TickMark & m, bool asMajor, CChart::EAxisType AxisType )
{
    xmlw.TagL( asMajor ? "c:majorTickMark" : "c:minorTickMark" ).Attr( "val", TickMarkToTag( m.type ) ).EndL();
    if( m.unit <= 0 )   // auto
        return;
    xmlw.TagL( asMajor ? "c:majorUnit" : "c:minorUnit" ).Attr( "val", m.unit ).EndL();
    if( AxisType == CChart::AXIS_DATE )
        xmlw.TagL( asMajor ? "c:majorTimeUnit" : "c:minorTimeUnit" ).Attr( "val", TimeUnitToTag( m.timeUnit ) ).EndL();
}

void CChart::AddAxisCommon( XMLWriter & xmlw, const Axis & x, uint32_t crossAxisId, bool AsX )
{
    xmlw.Tag( AxisTypeToTag( x.axisType ) ).TagL( "c:axId" ).Attr( "val", x.id ).EndL();

    xmlw.Tag( "c:scaling" ).TagL( "c:orientation" ).Attr( "val", "minMax" ).EndL();
    if( x.minValue != "" ) xmlw.TagL( "c:min" ).Attr( "val", x.minValue ).EndL();
    if( x.maxValue != "" ) xmlw.TagL( "c:max" ).Attr( "val", x.maxValue ).EndL();
    xmlw.End( "c:scaling" );

    xmlw.TagL( "c:delete" ).Attr( "val", x.pos == POS_NONE ? 1 : 0 ).EndL();
    xmlw.TagL( "c:axPos" ).Attr( "val", GetCharForPos( x.pos, AsX ? 'b' : 'l' ) ).EndL();

    AddTickMarks( xmlw, x.majorTickMarks, true, x.axisType );
    AddTickMarks( xmlw, x.minorTickMarks, false, x.axisType );

    if( x.gridLines == GRID_MAJOR ) xmlw.TagL( "c:majorGridlines" ).EndL();
    else if( x.gridLines == GRID_MINOR ) xmlw.TagL( "c:minorGridlines" ).EndL();
    else if( x.gridLines == GRID_MAJOR_N_MINOR ) xmlw.TagL( "c:majorGridlines" ).EndL().TagL( "c:minorGridlines" ).EndL();

    if( ! x.name.empty() )
        AddTitle( xmlw, x.name, x.nameSize, ! AsX );

    xmlw.TagL( "c:tickLblPos" ).Attr( "val", "nextTo" ).EndL();

    if( crossAxisId != 0 )
    {
        xmlw.TagL( "c:crossAx" ).Attr( "val", crossAxisId ).EndL().TagL( "c:crosses" );

        if( x.cross == CROSS_AUTO_ZERO ) xmlw.Attr( "val", "autoZero" ).EndL();
        else if( x.cross == CROSS_MIN ) xmlw.Attr( "val", "min" ).EndL();
        else if( x.cross == CROSS_MAX ) xmlw.Attr( "val", "max" ).EndL();
        else  xmlw.EndL();

        if( ! AsX )
            xmlw.TagL( "c:crossBetween" ).Attr( "val", "between" ).EndL();
    }

    xmlw.TagL( "c:numFmt" ).Attr( "formatCode", x.formatCode ).Attr( "sourceLinked", x.sourceLinked ? 1 : 0 ).EndL();
}


// ****************************************************************************
/// @brief  Internal method adds X axis block into chart
/// @param  x reference to an axis object
/// @param  crossAxisId id of axis that is used in pair with setting
/// @return no
// ****************************************************************************

void CChart::AddXAxis( XMLWriter & xmlw, const Axis & x, uint32_t crossAxisId )
{
    AddAxisCommon( xmlw, x, crossAxisId, true );

    if( x.lblAngle != -1 )
    {
        xmlw.Tag( "c:txPr" ).TagL( "a:bodyPr" ).Attr( "rot", x.lblAngle * 60000 ).EndL().TagL( "a:lstStyle" ).EndL();
        xmlw.Tag( "a:p" ).Tag( "a:pPr" ).TagL( "a:defRPr" ).EndL().End( "a:pPr" );
        xmlw.TagL( "a:endParaRPr" ).Attr( "lang", "en-US" ).EndL().End( "a:p" ).End( "c:txPr" );
    }

    xmlw.TagL( "c:auto" ).Attr( "val", 1 ).EndL();
    xmlw.TagL( "c:lblAlgn" ).Attr( "val", "ctr" ).EndL();
    xmlw.TagL( "c:lblOffset" ).Attr( "val", 100 ).EndL();

    if( x.lblSkipInterval != -1 ) xmlw.TagL( "c:tickLblSkip" ).Attr( "val", x.lblSkipInterval ).EndL();
    if( x.markSkipInterval != -1 ) xmlw.TagL( "c:tickMarkSkip" ).Attr( "val", x.markSkipInterval ).EndL();
    xmlw.TagL( "c:noMultiLvlLbl" ).Attr( "val", 0 ).EndL();

    xmlw.End( AxisTypeToTag( x.axisType ) );
}
// ****************************************************************************
/// @brief  Internal method adds Y axis block into chart
/// @param  y reference to an axis object
/// @param  crossAxisId id of axis that is used in pair with setting
/// @return no
// ****************************************************************************
void CChart::AddYAxis( XMLWriter & xmlw, const Axis & y, uint32_t crossAxisId )
{
    AddAxisCommon( xmlw, y, crossAxisId, false );
    xmlw.End( AxisTypeToTag( y.axisType ) );
}

// ****************************************************************************
/// @brief  Adds line chart xml section
/// @param  xAxis           reference to axis object
/// @param  yAxisId         id of using y axis
/// @param  diagramm        reference to diagram parameters
/// @param  series          reference to the vector of line sets
/// @param  firstSeriesId   is used for synchronization between different line charts
/// @return no
// ****************************************************************************
void CChart::AddLineChart( XMLWriter & xmlw, Axis & xAxis, uint32_t yAxisId, const Diagramm & diagramm, const std::vector<Series> & series, size_t firstSeriesId )
{
    xmlw.Tag( "c:lineChart" );
    xmlw.TagL( "c:grouping" ).Attr( "val", "standard" ).EndL();
    xmlw.TagL( "c:varyColors" ).Attr( "val", 0 ).EndL();

    for( std::vector<Series>::const_iterator it = series.begin(); it != series.end(); it++ )
    {
        xmlw.Tag( "c:ser" );
        xmlw.TagL( "c:idx" ).Attr( "val", firstSeriesId ).EndL();
        xmlw.TagL( "c:order" ).Attr( "val", firstSeriesId ).EndL();

        const char * markerID = "none";
        //   if(style == SCATTER_POINT)
        switch( it->Marker.Type )
        {
            case Series::symDiamond:  markerID = "diamond";
                break;
            case Series::symCircle:   markerID = "circle";
                break;
            case Series::symSquare:   markerID = "square";
                break;
            case Series::symTriangle: markerID = "triangle";
                break;
            default:;
        }
        AddMarker( xmlw, * it, markerID );

        if( ! it->title.empty() )
        {
            xmlw.Tag( "c:tx" ).Tag( "c:strRef" ).Tag( "c:strCache" );
            xmlw.TagL( "c:ptCount" ).Attr( "val", 1 ).EndL();
            xmlw.Tag( "c:pt" ).Attr( "idx", 0 ).TagOnlyContent( "c:v", it->title ).End( "c:pt" );
            xmlw.End( "c:strCache" ).End( "c:strRef" ).End( "c:tx" );
        }
        if( it->JoinType == Series::joinNone )
        {
            xmlw.Tag( "c:spPr" ).Tag( "a:ln" ).Attr( "w", 28000 ).TagL( "a:noFill" ).EndL().End( "a:ln" ).End( "c:spPr" );
        }
        else
        {
            const char * dashID = "solid";
            switch( it->DashType )
            {
                case Series::dashDot:       dashID = "sysDot"; break;
                case Series::dashShortDash: dashID = "sysDash"; break;
                case Series::dashDash:      dashID = "dash"; break;
                default:;
            }
            xmlw.Tag( "c:spPr" ).Tag( "a:ln" ).Attr( "w", std::floor( it->LineWidth * 12700 ) );
            if( it->LineColor.size() == 6 )
            {
                xmlw.Tag( "a:solidFill" ).TagL( "a:srgbClr" ).Attr( "val", it->LineColor ).EndL().End( "a:solidFill" );
            }
            xmlw.TagL( "a:prstDash" ).Attr( "val", dashID ).EndL().End( "a:ln" ).End( "c:spPr" );

        }

        if( ( it->catSheet != NULL ) && ( ( it->catAxisTo.row != 0 ) || ( it->catAxisTo.col != 0 ) ) )
        {
            std::string cfRange = CellRangeString( it->catSheet->GetTitle(),
                                                   CellCoord( it->catAxisFrom.row, it->catAxisFrom.col ),
                                                   CellCoord( it->catAxisTo.row, it->catAxisTo.col ) );
            xmlw.Tag( "c:cat" ).Tag( "c:numRef" ).TagOnlyContent( "c:f", cfRange ).End( "c:numRef" ).End( "c:cat" );
        }

        std::string cfRange = CellRangeString( it->valSheet->GetTitle(),
                                               CellCoord( it->valAxisFrom.row, it->valAxisFrom.col ),
                                               CellCoord( it->valAxisTo.row, it->valAxisTo.col ) );
        xmlw.Tag( "c:val" ).Tag( "c:numRef" ).TagOnlyContent( "c:f", cfRange ).End( "c:numRef" ).End( "c:val" );
        xmlw.TagL( "c:smooth" ).Attr( "val", it->JoinType == Series::joinSmooth ? 1 : 0 ).EndL();
        xmlw.End( "c:ser" );

        firstSeriesId++;
    }

    AddDataLabels( xmlw, diagramm.dataLabels );
    xmlw.TagL( "c:marker" ).Attr( "val", 1 ).EndL();
    xmlw.TagL( "c:smooth" ).Attr( "val", 0 ).EndL();
    xmlw.TagL( "c:axId" ).Attr( "val", xAxis.id ).EndL();
    xmlw.TagL( "c:axId" ).Attr( "val", yAxisId ).EndL();

    xmlw.End( "c:lineChart" );
}

// ****************************************************************************
/// @brief  Adds bar chart xml section
/// @param  xAxis           reference to axis object
/// @param  yAxisId         id of using y axis
/// @param  diagramm        reference to diagram parameters
/// @param  series          reference to the vector of line sets
/// @param  firstSeriesId   is used for synchronization between different line charts
/// @param	barDir			bars` direction (horizontal or vertical)
/// @param	barGroup		bars` relative position
/// @see	EBarDirection
/// @see	EBarGrouping
/// @return no
// ****************************************************************************
void CChart::AddBarChart( XMLWriter & xmlw, Axis & xAxis, uint32_t yAxisId, const Diagramm & diagramm, const std::vector<Series> & series, size_t firstSeriesId, EBarDirection barDir, EBarGrouping barGroup )
{
    xmlw.Tag( "c:barChart" );
    if( barDir == BAR_DIR_VERTICAL ) xmlw.TagL( "c:barDir" ).Attr( "val", "col" ).EndL();
    else if( barDir == BAR_DIR_HORIZONTAL ) xmlw.TagL( "c:barDir" ).Attr( "val", "bar" ).EndL();

    if( barGroup == BAR_GROUP_CLUSTERED ) xmlw.TagL( "c:grouping" ).Attr( "val", "clustered" ).EndL();
    else if( barGroup == BAR_GROUP_STACKED ) xmlw.TagL( "c:grouping" ).Attr( "val", "stacked" ).EndL();
    else if( barGroup == BAR_GROUP_PERCENT_STACKED ) xmlw.TagL( "c:grouping" ).Attr( "val", "percentStacked" ).EndL();

    xmlw.TagL( "c:varyColors" ).Attr( "val", 1 ).EndL();

    for( std::vector<Series>::const_iterator it = series.begin(); it != series.end(); it++ )
    {
        xmlw.Tag( "c:ser" );
        xmlw.TagL( "c:idx" ).Attr( "val", firstSeriesId ).EndL();
        xmlw.TagL( "c:order" ).Attr( "val", firstSeriesId ).EndL();

        if( ! it->title.empty() ) xmlw.Tag( "c:tx" ).TagOnlyContent( "c:v", it->title ).End( "c:tx" );
        if( ( it->catSheet != NULL ) && ( ( it->catAxisTo.row != 0 ) || ( it->catAxisTo.col != 0 ) ) )
        {
            std::string cfRange = CellRangeString( it->catSheet->GetTitle(),
                                                   CellCoord( it->catAxisFrom.row, it->catAxisFrom.col ),
                                                   CellCoord( it->catAxisTo.row, it->catAxisTo.col ) );
            xmlw.Tag( "c:cat" ).Tag( "c:numRef" ).TagOnlyContent( "c:f", cfRange ).End( "c:numRef" ).End( "c:cat" );
        }

        switch( it->barFillStyle )
        {
            case Series::BAR_FILL_NONE:
                xmlw.Tag( "c:spPr" ).Tag( "a:noFill" ).End( "a:noFill" ).End( "c:spPr" );
                break;
            case Series::BAR_FILL_SOLID:
                xmlw.Tag( "c:spPr" ).Tag( "a:solidFill" ).Tag( "a:srgbClr" ).Attr( "val", it->LineColor );
                xmlw.End( "a:srgbClr" ).End( "a:solidFill" ).End( "c:spPr" );
                break;
            case Series::BAR_FILL_AUTOMATIC:
                break;
        }
        xmlw.TagL( "c:invertIfNegative" ).Attr( "val", it->barInvertIfNegative ? 1 : 0 ).EndL();

        std::string cfRange = CellRangeString( it->valSheet->GetTitle(),
                                               CellCoord( it->valAxisFrom.row, it->valAxisFrom.col ),
                                               CellCoord( it->valAxisTo.row, it->valAxisTo.col ) );
        xmlw.Tag( "c:val" ).Tag( "c:numRef" ).TagOnlyContent( "c:f", cfRange ).End( "c:numRef" ).End( "c:val" );

        if( it->barInvertIfNegative )
        {
            switch( it->barFillStyle )
            {
                case Series::BAR_FILL_NONE:
                    break;
                case Series::BAR_FILL_SOLID:
                    xmlw.Tag( "c:extLst" ).Tag( "c:ext" ).Attr( "uri", "{6F2FDCE9-48DA-4B69-8628-5D25D57E5C99}" ).Attr( "xmlns:c14", ns_c14 );
                    xmlw.Tag( "c14:invertSolidFillFmt" ).Tag( "c14:spPr" ).Attr( "xmlns:c14", ns_c14 );
                    xmlw.Tag( "a:solidFill" ).Tag( "a:srgbClr" ).Attr( "val", it->barInvertedColor ).End( "a:srgbClr" ).End( "a:solidFill" );
                    xmlw.End( "c14:spPr" ).End( "c14:invertSolidFillFmt" );
                    xmlw.End( "c:ext" ).End( "c:extLst" );
                    break;
                case Series::BAR_FILL_AUTOMATIC:
                    break;
            }
        }

        xmlw.TagL( "c:smooth" ).Attr( "val", it->JoinType == Series::joinSmooth ? 1 : 0 ).EndL();

        xmlw.End( "c:ser" );

        firstSeriesId++;
    }
    AddDataLabels( xmlw, diagramm.dataLabels );
    xmlw.TagL( "c:marker" ).Attr( "val", 1 ).EndL();
    xmlw.TagL( "c:smooth" ).Attr( "val", 0 ).EndL();
    xmlw.TagL( "c:axId" ).Attr( "val", xAxis.id ).EndL();
    xmlw.TagL( "c:axId" ).Attr( "val", yAxisId ).EndL();

    xmlw.End( "c:barChart" );
}

// ****************************************************************************
/// @brief  Adds scatter chart xml section
/// @param  xAxisId         id of using x axis
/// @param  yAxisId         id of using y axis
/// @param  diagramm        reference to diagram parameters
/// @param  series          reference to the vector of line sets
/// @param  firstSeriesId   is used for synchronization between different line charts
/// @param	style			determines diagramm style
/// @see	EScatterStyle
/// @return no
// ****************************************************************************
void CChart::AddScatterChart( XMLWriter & xmlw, uint32_t xAxisId, uint32_t yAxisId, const Diagramm & diagramm, const std::vector<Series> & series, size_t firstSeriesId, EScatterStyle style )
{
    xmlw.Tag( "c:scatterChart" );

    if( style == SCATTER_FILL ) xmlw.TagL( "c:scatterStyle" ).Attr( "val", "smoothMarker" ).EndL();
    else if( style == SCATTER_POINT ) xmlw.TagL( "c:scatterStyle" ).Attr( "val", "lineMarker" ).EndL();
    xmlw.TagL( "c:varyColors" ).Attr( "val", 0 ).EndL();
    for( std::vector<Series>::const_iterator it = series.begin(); it != series.end(); it++ )
    {
        xmlw.Tag( "c:ser" );
        xmlw.TagL( "c:idx" ).Attr( "val", firstSeriesId ).EndL();
        xmlw.TagL( "c:order" ).Attr( "val", firstSeriesId ).EndL();
        const char * markerID = "none";
        if( style == SCATTER_POINT )
            switch( it->Marker.Type )
            {
                case
                        Series::symDiamond:  markerID = "diamond";
                    break;
                case
                        Series::symCircle:   markerID = "circle";
                    break;
                case
                        Series::symSquare:   markerID = "square";
                    break;
                case
                        Series::symTriangle: markerID = "triangle";
                    break;
                default:;
            }
        AddMarker( xmlw, * it, markerID );

        if( ! it->title.empty() )
            xmlw.Tag( "c:tx" ).TagOnlyContent( "c:v", it->title ).End( "c:tx" );
        //  if( (style == SCATTER_POINT) && (it->lineWidth>0) )
        if( it->JoinType == Series::joinNone )
        {
            xmlw.Tag( "c:spPr" ).Tag( "a:ln" ).Attr( "w", 28000 ).TagL( "a:noFill" ).EndL().End( "a:ln" ).End( "c:spPr" );
        }
        else
        {
            const char * dashID = "solid";
            switch( it->DashType )
            {
                case
                        Series::dashDot:       dashID = "sysDot";
                    break;
                case
                        Series::dashShortDash: dashID = "sysDash";
                    break;
                case
                        Series::dashDash:      dashID = "dash";
                    break;

                default:;
            }
            xmlw.Tag( "c:spPr" ).Tag( "a:ln" ).Attr( "w", std::floor( it->LineWidth * 12700 ) );
            if( it->LineColor.size() == 6 )
            {
                xmlw.Tag( "a:solidFill" ).TagL( "a:srgbClr" ).Attr( "val", it->LineColor ).EndL().End( "a:solidFill" );
            }
            xmlw.TagL( "a:prstDash" ).Attr( "val", dashID ).EndL().End( "a:ln" ).End( "c:spPr" );

        }
        std::string cfRange = CellRangeString( it->catSheet->GetTitle(),
                                               CellCoord( it->catAxisFrom.row, it->catAxisFrom.col ),
                                               CellCoord( it->catAxisTo.row, it->catAxisTo.col ) );
        xmlw.Tag( "c:xVal" ).Tag( "c:numRef" ).TagOnlyContent( "c:f", cfRange ).End( "c:numRef" ).End( "c:xVal" );

        cfRange = CellRangeString( it->valSheet->GetTitle(),
                                   CellCoord( it->valAxisFrom.row, it->valAxisFrom.col ),
                                   CellCoord( it->valAxisTo.row, it->valAxisTo.col ) );
        xmlw.Tag( "c:yVal" ).Tag( "c:numRef" ).TagOnlyContent( "c:f", cfRange ).End( "c:numRef" ).End( "c:yVal" );
        xmlw.TagL( "c:smooth" ).Attr( "val", it->JoinType == Series::joinSmooth ? 1 : 0 ).EndL();

        xmlw.End( "c:ser" );

        firstSeriesId++;
    }
    AddDataLabels( xmlw, diagramm.dataLabels );
    xmlw.TagL( "c:marker" ).Attr( "val", 1 ).EndL();
    xmlw.TagL( "c:smooth" ).Attr( "val", 0 ).EndL();
    xmlw.TagL( "c:axId" ).Attr( "val", xAxisId ).EndL();
    xmlw.TagL( "c:axId" ).Attr( "val", yAxisId ).EndL();

    xmlw.End( "c:scatterChart" );
}

// ****************************************************************************
/// @brief  Adds pie chart xml section
/// @param  xAxis           reference to axis object
/// @param  diagramm        reference to diagram parameters
/// @param  series          reference to the vector of line sets
/// @param  firstSeriesId   is used for synchronization between different line charts
/// @return no
// ****************************************************************************
void CChart::AddPieChart( XMLWriter & xmlw, const Diagramm & diagramm, const std::vector<CChart::Series> & series, size_t firstSeriesId )
{
    xmlw.Tag( "c:pieChart" );
    xmlw.TagL( "c:varyColors" ).Attr( "val", 1 ).EndL();

    for( std::vector<Series>::const_iterator it = series.begin(); it != series.end(); it++ )
    {
        xmlw.Tag( "c:ser" );
        xmlw.TagL( "c:idx" ).Attr( "val", firstSeriesId ).EndL();
        xmlw.TagL( "c:order" ).Attr( "val", firstSeriesId ).EndL();

        if( ! it->title.empty() )
            xmlw.Tag( "c:tx" ).TagOnlyContent( "c:v", it->title ).End( "c:tx" );
        AddDataLabels( xmlw, it->dataLabels, true );
        if( ( it->catSheet != NULL ) && ( ( it->catAxisTo.row != 0 ) || ( it->catAxisTo.col != 0 ) ) )
        {
            std::string cfRange = CellRangeString( it->catSheet->GetTitle(),
                                                   CellCoord( it->catAxisFrom.row, it->catAxisFrom.col ),
                                                   CellCoord( it->catAxisTo.row, it->catAxisTo.col ) );
            xmlw.Tag( "c:cat" ).Tag( "c:numRef" ).TagOnlyContent( "c:f", cfRange ).End( "c:numRef" ).End( "c:cat" );
        }

        std::string cfRange = CellRangeString( it->valSheet->GetTitle(),
                                               CellCoord( it->valAxisFrom.row, it->valAxisFrom.col ),
                                               CellCoord( it->valAxisTo.row, it->valAxisTo.col ) );
        xmlw.Tag( "c:val" ).Tag( "c:numRef" ).TagOnlyContent( "c:f", cfRange ).End( "c:numRef" ).End( "c:val" );
        xmlw.End( "c:ser" );

        firstSeriesId++;
    }
    AddDataLabels( xmlw, diagramm.dataLabels, true );
    xmlw.TagL( "c:firstSliceAng" ).Attr( "val", diagramm.firstSliceAng ).EndL();

    xmlw.End( "c:pieChart" );
}

void CChart::AddMarker( XMLWriter & xmlw, const CChart::Series & ser, const char * markerID )
{
    xmlw.Tag( "c:marker" ).TagL( "c:symbol" ).Attr( "val", markerID ).EndL().TagL( "c:size" ).Attr( "val", ser.Marker.Size ).EndL();
    const bool IsFillColor = ser.Marker.FillColor.size() == 6;
    const bool IsLineColor = ser.Marker.LineColor.size() == 6;
    if( IsFillColor || IsLineColor ) // check formal RGB record format
    {
        xmlw.Tag( "c:spPr" );
        if( IsFillColor )
            xmlw.Tag( "a:solidFill" ).TagL( "a:srgbClr" ).Attr( "val", ser.Marker.FillColor ).EndL().End( "a:solidFill" ); // marker fill
        if( IsLineColor )
            xmlw.Tag( "a:ln" ).Attr( "w", std::floor( ser.Marker.LineWidth * 12700 ) ).Tag( "a:solidFill" ).TagL( "a:srgbClr" ).Attr( "val", ser.Marker.LineColor ).EndL().End( "a:solidFill" ).End( "a:ln" ); // marker line
        xmlw.End( "c:spPr" );
    }
    xmlw.End( "c:marker" );
}

static inline void AddFillPath( XMLWriter & xmlw, const char * PathName, CChart::EGradientFillDirection Dir )
{
    const char * FTag = "a:fillToRect", * TTag = "a:tileRect", * PTag = "a:path";
    xmlw.Tag( "a:path" ).Attr( "path", PathName );
    switch( Dir )
    {
        case CChart::FROM_BOTTOM_RIGHT_CORNER:
        {
            xmlw.TagL( FTag ).Attr( "l", 100000 ).Attr( "t", 100000 ).EndL().End( PTag );
            xmlw.TagL( TTag ).Attr( "r", -100000 ).Attr( "b", -100000 ).EndL();
            break;
        }
        case CChart::FROM_BOTTOM_LEFT_CORNER:
        {
            xmlw.TagL( FTag ).Attr( "t", 100000 ).Attr( "r", 100000 ).EndL().End( PTag );
            xmlw.TagL( TTag ).Attr( "l", -100000 ).Attr( "b", -100000 ).EndL();
            break;
        }
        case CChart::FROM_CENTER:
        {
            xmlw.TagL( FTag ).Attr( "l", 50000 ).Attr( "t", 50000 ).Attr( "r", 50000 ).Attr( "b", 50000 ).EndL().End( PTag );
            xmlw.TagL( TTag ).EndL();
            break;
        }
        case CChart::FROM_TOP_RIGHT_CORNER:
        {
            xmlw.TagL( FTag ).Attr( "l", 100000 ).Attr( "b", 100000 ).EndL().End( PTag );
            xmlw.TagL( TTag ).Attr( "t", -100000 ).Attr( "r", -100000 ).EndL();
            break;
        }
        case CChart::FROM_TOP_LEFT_CORNER:
        {
            xmlw.TagL( FTag ).Attr( "r", 100000 ).Attr( "b", 100000 ).EndL().End( PTag );
            xmlw.TagL( TTag ).Attr( "l", -100000 ).Attr( "t", -100000 ).EndL();
            break;
        }
    }
}

static inline const char * PatternPresetCode( CChart::EPatternFillStyle Style )
{
    switch( Style )
    {
        case CChart::PERCENT_5  :   return "pct5";
        case CChart::PERCENT_10 :   return "pct10";
        case CChart::PERCENT_20 :   return "pct20";
        case CChart::PERCENT_25 :   return "pct25";
        case CChart::PERCENT_30 :   return "pct30";
        case CChart::PERCENT_40 :   return "pct40";
        case CChart::PERCENT_50 :   return "pct50";
        case CChart::PERCENT_60 :   return "pct60";
        case CChart::PERCENT_70 :   return "pct70";
        case CChart::PERCENT_75 :   return "pct75";
        case CChart::PERCENT_80 :   return "pct80";
        case CChart::PERCENT_90 :   return "pct90";

        case CChart::LIGHT_DOWNWARD_DIAGONAL:   return "ltDnDiag";
        case CChart::DARK_DOWNWARD_DIAGONAL :   return "dkDnDiag";
        case CChart::WIDE_DOWNWARD_DIAGONAL :   return "wdDnDiag";
        case CChart::LIGHT_UPWARD_DIAGONAL  :   return "ltUpDiag";
        case CChart::DARK_UPWARD_DIAGONAL   :   return "dkUpDiag";
        case CChart::WIDE_UPWARD_DIAGONAL   :   return "wdUpDiag ";

        case CChart::LIGHT_VERTICAL     :   return "ltVert";
        case CChart::NARROW_VERTICAL    :   return "narVert";
        case CChart::DARK_VERTICAL      :   return "dkVert";
        case CChart::LIGHT_HORIZONTAL   :   return "ltHorz";
        case CChart::NARROW_HORIZONTAL  :   return "narHorz";
        case CChart::DARK_HORIZONTAL    :   return "dkHorz";

        case CChart::DASHED_DOWNWARD_DIAGONAL   :   return "dashDnDiag";
        case CChart::DASHED_UPWARD_DIAGONAL     :   return "dashUpDiag";
        case CChart::DASHED_HORIZONTAL          :   return "dashHorz";
        case CChart::DASHED_VERTICAL            :   return "dashVert";
        case CChart::SMALL_CONFETTI             :   return "smConfetti";
        case CChart::LARGE_CONFETTI             :   return "lgConfetti";

        case CChart::ZIG_ZAG            :   return "zigZag";
        case CChart::WAVE               :   return "wave";
        case CChart::DIAGONAL_BRICK     :   return "diagBrick";
        case CChart::HORIZONTAL_BRICK   :   return "horzBrick";
        case CChart::WEAVE              :   return "weave";
        case CChart::PLAID              :   return "plaid";

        case CChart::DIVOT          :   return "divot";
        case CChart::DOTTED_GRID    :   return "dotGrid";
        case CChart::DOTTED_DIAMOND :   return "dotDmnd";
        case CChart::SHINGLE        :   return "shingle";
        case CChart::TRELLIS        :   return "trellis";
        case CChart::SPHERE         :   return "sphere";

        case CChart::SMALL_GRID         :   return "smGrid";
        case CChart::LARGE_GRID         :   return "lgGrid";
        case CChart::SMALL_CHECKER_BOARD:   return "smCheck";
        case CChart::LARGE_CHECKER_BOARD:   return "lgCheck";
        case CChart::OUTLINED_DIAMOND   :   return "openDmnd";
        case CChart::SOLID_DIAMOND      :   return "solidDmnd";
    }
    return "";
}

void CChart::AddAreaFill( XMLWriter & xmlw, const CChart::AreaFill & areaFill )
{
    switch( areaFill.Style )
    {
        case PLOT_AREA_FILL_NONE    : break;
        case PLOT_AREA_FILL_SOLID   :
        {
            if( areaFill.SolidColor.size() == 6 )
                xmlw.Tag( "c:spPr" ).Tag( "a:solidFill" ).TagL( "a:srgbClr" ).Attr( "val", areaFill.SolidColor ).EndL().End( "a:solidFill" ).End( "c:spPr" );
            break;
        }
        case PLOT_AREA_FILL_GRADIENT:
        {
            xmlw.Tag( "c:spPr" ).Tag( "a:gradFill" ).Attr( "flip", "none" ).Attr( "rotWithShape", 1 ).Tag( "a:gsLst" );
            const GradientFill & GF = areaFill.Gradient;
            GradientStops::const_iterator it = GF.ColorPoints.begin();
            for( ; it != GF.ColorPoints.end(); it++ )
                xmlw.Tag( "a:gs" ).Attr( "pos", it->first * 1000 ).Tag( "a:srgbClr" ).Attr( "val", it->second ).End( "a:srgbClr" ).End( "a:gs" );
            xmlw.End( "a:gsLst" );
            switch( GF.FillType )
            {
                case GradientFill::linear:
                {
                    if( GF.LinearAngle != 0.0 )
                        xmlw.Tag( "a:lin" ).Attr( "ang", int( GF.LinearAngle * 60000 ) ).Attr( "scaled", GF.LinearScaleAngle ? 1 : 0 ).End( "a:lin" );
                    break;
                }
                case GradientFill::radial       :  AddFillPath( xmlw, "circle", GF.FillDirection ); break;
                case GradientFill::rectangular  :  AddFillPath( xmlw, "rect", GF.FillDirection ); break;
                case GradientFill::path         :  AddFillPath( xmlw, "shape", CChart::FROM_CENTER ); break;
            }
            xmlw.End( "a:gradFill" ).End( "c:spPr" );
            break;
        }
        case PLOT_AREA_FILL_PATTERN:
        {
            xmlw.Tag( "c:spPr" ).Tag( "a:pattFill" ).Attr( "prst", PatternPresetCode( areaFill.Pattern ) );
            xmlw.Tag( "a:fgClr" ).TagL( "a:srgbClr" ).Attr( "val", areaFill.PatternFgColor ).EndL().End( "a:fgClr" );
            xmlw.Tag( "a:bgClr" ).TagL( "a:srgbClr" ).Attr( "val", areaFill.PatternBgColor ).EndL().End( "a:bgClr" );
            xmlw.End( "a:pattFill" ).End( "c:spPr" );
            break;
        }
    }
}

void CChart::AddDataLabels( XMLWriter & xmlw, const CChart::DataLabels & dataLabels, bool UseLeaderLines )
{
    xmlw.Tag( "c:dLbls" );
    xmlw.TagL( "c:showLegendKey" ).Attr( "val", int( dataLabels.showLegendKey ) ).EndL();
    xmlw.TagL( "c:showVal" ).Attr( "val", int( dataLabels.showVal ) ).EndL();
    xmlw.TagL( "c:showCatName" ).Attr( "val", int( dataLabels.showCategoryName ) ).EndL();
    xmlw.TagL( "c:showSerName" ).Attr( "val", int( dataLabels.showSeriesName ) ).EndL();
    xmlw.TagL( "c:showPercent" ).Attr( "val", int( dataLabels.showPercent ) ).EndL();
    xmlw.TagL( "c:showBubbleSize" ).Attr( "val", int( dataLabels.showBubbleSize ) ).EndL();
    if( UseLeaderLines )
        xmlw.TagL( "c:showLeaderLines" ).Attr( "val", int( dataLabels.showLeaderLines ) ).EndL();
    xmlw.End( "c:dLbls" );
}

std::string CChart::CellRangeString( const std::string & Title, const CellCoord & CellFrom, const CellCoord & szCellTo )
{
    CellCoord::TConvBuf Buffer;
    std::stringstream RangeStream;
    RangeStream << '\'' << Title << "\'!$" << CellFrom.ToString( Buffer );  // Don't join these
    RangeStream << ":$" << szCellTo.ToString( Buffer );                     // two rows -> undefined behavior
    return RangeStream.str();
}

}
