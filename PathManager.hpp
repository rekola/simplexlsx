#ifndef XLSX_PATHMANAGER_HPP
#define XLSX_PATHMANAGER_HPP

#include "Xlsx/SimpleXlsxDef.h"

namespace SimpleXlsx
{

    class PathManager
    {
        public:
            inline PathManager( const _tstring & temp_path ) : m_temp_path( temp_path ) {}

            inline ~PathManager()
            {
                ClearTemp();
            }

            //Register XML and creating all necessary subdirectories
            inline const _tstring RegisterXML( const _tstring & PathToFile )
            {
                _tstring Result = m_temp_path + PathToFile;
                MakeDirectory( Result );
                m_contentFiles.push_back( PathToFile );
                return Result;
            }

            //Deletes all temporary files and directories which have been created
            void ClearTemp();

            inline const std::list<_tstring> & ContentFiles() const
            {
                return m_contentFiles;
            }

            // *INDENT-OFF*   For AStyle tool

            //Encode File Path for the Operation System
#if ! defined( _WIN32 ) && defined( _UNICODE )      //Linux with Unicode
            static std::string PathEncode( const wchar_t * Path );
            static inline std::string PathEncode( const std::wstring & Path )   { return PathEncode( Path.c_str() ); }
#else
            static inline _tstring PathEncode( const TCHAR * Path )             { return Path; }
            static inline _tstring PathEncode( const _tstring & Path )          { return Path; }
#endif
            // *INDENT-ON*   For AStyle tool

        private:
            //Disable copy and assignment
            PathManager( const PathManager & that );
            PathManager & operator=( const PathManager & );

            const _tstring       &      m_temp_path;		///< path to the temporary directory (unique for a book)
            std::vector<_tstring>       m_temp_dirs;        ///< a series of temporary subdirectories
            std::list<_tstring>         m_contentFiles;		///< a series of relative file pathes to be saved inside xlsx archive

            // ****************************************************************************
            /// @brief  Function to create nested directories` tree
            /// @param  dirName directories` tree to be created
            /// @return Boolean result of the operation
            // ****************************************************************************
            bool MakeDirectory( const _tstring & dirName );
    };

}
#endif // XLSX_PATHMANAGER_HPP
