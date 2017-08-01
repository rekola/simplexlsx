#ifndef XLSX_DRAWING_H
#define XLSX_DRAWING_H

#include "SimpleXlsxDef.h"

namespace SimpleXlsx
{
    class CChart;

    class PathManager;
    class XMLWriter;

    class CDrawing
    {
        public:
            //Index of Drawing
            inline size_t GetIndex() const
            {
                return m_index;
            }

            //Is no registered Charts
            inline bool IsEmpty() const
            {
                return m_charts.empty();
            }

        protected:
            CDrawing( size_t index, PathManager & pathmanager );
            virtual ~CDrawing();

        private:
            //Disable copy and assignment
            CDrawing( const CDrawing & that );
            CDrawing & operator=( const CDrawing & );

            size_t              m_index;            ///< Drawing ID number
            PathManager    &    m_pathManager;      ///< reference to XML PathManager

            struct ChartInfo
            {
                enum AnchorType
                {
                    absoluteAnchor,         //For CChartSheet
                    twoCellAnchor,          //For chart on CWorkSheet
                };

                CChart   *  Chart;
                AnchorType  AType;
                ChartPoint  TopLeft;        //For chart on CWorkSheet
                ChartPoint  BottomRight;    //For chart on CWorkSheet
            };

            std::vector<ChartInfo> m_charts;


            //Append Chart from Chartsheet
            inline void AppendChart( CChart * Chart )
            {
                ChartInfo CInfo = { Chart, ChartInfo::absoluteAnchor, ChartPoint(), ChartPoint() };
                m_charts.push_back( CInfo );
            }

            //Append Chart from Worksheet
            inline void AppendChart( CChart * Chart, const ChartPoint & TopLeft, const ChartPoint & BottomRight )
            {
                ChartInfo CInfo = { Chart, ChartInfo::twoCellAnchor, TopLeft, BottomRight };
                m_charts.push_back( CInfo );
            }

            bool Save();

            void SaveDrawingRels();
            void SaveDrawing();
            void SaveChartSection( XMLWriter & xmlw, CChart * chart, int rId );
            void SaveChartPoint( XMLWriter & xmlw, const char * Tag, const ChartPoint & Point );

            friend class CWorkbook;
    };

}

#endif // XLSX_DRAWING_H
