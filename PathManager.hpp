/*
  SimpleXlsxWriter
  Copyright (C) 2012-2018 Pavel Akimov <oxod.pavel@gmail.com>, Alexandr Belyak <programmeralex@bk.ru>

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
                return RegisterFile( PathToFile );
            }

            //Creating all necessary subdirectories and copy image file
            bool RegisterImage( const _tstring & LocalPath, const _tstring & XLSX_Path );

            //Deletes all temporary files and directories which have been created
            void ClearTemp();

            inline const std::list<_tstring> & ContentFiles() const
            {
                return m_contentFiles;
            }

            // *INDENT-OFF*   For AStyle tool

            //Encode File Path for the Operation System
#if defined( _UNICODE )

#if ! defined( _WIN32 ) //Linux with Unicode
            static std::string PathEncode( const wchar_t * Path );
            static inline std::string PathEncode( const std::wstring & Path )   { return PathEncode( Path.c_str() ); }
#else   //Windows with Unicode
            static std::string PathEncode( const wchar_t * Path );
            static inline std::string PathEncode( const std::wstring & Path )   { return PathEncode( Path.c_str() ); }
#endif

#else
            static inline std::string PathEncode( const char * Path )           { return Path; }
            static inline std::string PathEncode( const std::string & Path )    { return Path; }
#endif  // _UNICODE
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

            //Register file for XLSX and creating all necessary subdirectories
            inline const _tstring RegisterFile( const _tstring & PathToFile )
            {
                _tstring Result = m_temp_path + PathToFile;
                MakeDirectory( Result );
                m_contentFiles.push_back( PathToFile );
                return Result;
            }
    };

}
#endif // XLSX_PATHMANAGER_HPP
