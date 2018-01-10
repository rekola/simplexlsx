#ifndef XLSX_CHARTSHEET_H
#define XLSX_CHARTSHEET_H

#include <stdlib.h>

// ****************************************************************************
/// @brief  The following namespace contains declarations of a number of classes
///         which allow writing .xlsx files with formulae and charts
/// @note   All classes inside the namepace are used together, so a separate
///         using will not guarantee stability and reliability
// ****************************************************************************
namespace SimpleXlsx
{
    class CChart;
    class CDrawing;
    class PathManager;

    // ****************************************************************************
    /// @brief	The class CChartsheet is used for creation sheet with chart.
    /// @see    EChartTypes supplying types of charts
    /// @note   All created chartsheets inside workbook will be allocated on its` own sheet
    // ****************************************************************************
    class CChartsheet
    {
        public:
            //Index of Chartsheet
            inline size_t GetIndex() const
            {
                return m_index;
            }

            //Reference to the Chart
            inline CChart & Chart()
            {
                return m_Chart;
            }

        protected:
            CChartsheet( size_t index, CChart & chart, CDrawing & drawing, PathManager & pathmanager );
            virtual ~CChartsheet();

        private:
            //Disable copy and assignment
            CChartsheet( const CChartsheet & that );
            CChartsheet & operator=( const CChartsheet & );

            size_t              m_index;            ///< chart ID number
            CChart       &      m_Chart;            ///< Reference to chart
            CDrawing      &     m_Drawing;          ///< Reference to drawing object
            PathManager    &    m_pathManager;      ///< reference to XML PathManager

            bool Save();

            friend class CWorkbook;
    };

}	// namespace SimpleXlsx

#endif	// XLSX_CHARTSHEET_H
