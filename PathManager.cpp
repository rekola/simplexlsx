#include <algorithm>
#include <cerrno>
#include <cstring>

#include "PathManager.hpp"

#ifdef _WIN32
#include <direct.h>
#else
#include <unistd.h>
#include <sys/stat.h>
#include <sys/types.h>
#include <sys/syscall.h>
#endif

namespace SimpleXlsx
{

    //Deletes all temporary files and directories which have been created
    void PathManager::ClearTemp()
    {
        for( std::list<_tstring>::const_iterator it = m_contentFiles.begin(); it != m_contentFiles.end(); it++ )
            _tremove( PathEncode( m_temp_path + ( * it ) ).c_str() );
        m_contentFiles.clear();
        for( std::vector<_tstring>::const_reverse_iterator it = m_temp_dirs.rbegin(); it != m_temp_dirs.rend(); it++ )
            _trmdir( PathEncode( ( * it ) ).c_str() );
        m_temp_dirs.clear();
    }


    // ****************************************************************************
    /// @brief  Function to create nested directories` tree
    /// @param  dirName directories` tree to be created
    /// @return Boolean result of the operation
    // ****************************************************************************
    bool PathManager::MakeDirectory( const _tstring & dirName )
    {
        _tstring part = _T( "" );
        size_t oldPointer = 0;
        size_t currPointer = 0;
        int32_t res = -1;

        for( size_t i = 0; i < dirName.length(); i++ )
        {
            if( ( dirName.at( i ) == _T( '\\' ) ) || ( dirName.at( i ) == _T( '/' ) ) )
            {
                part += dirName.substr( oldPointer, currPointer - oldPointer ) + _T( "/" );
                oldPointer = currPointer + 1;
#ifdef _WIN32
                std::replace( part.begin(), part.end(), '/', '\\' );
                res = _tmkdir( part.c_str() );
#else
                res = mkdir( PathEncode( part ).c_str(), 0777 );
#endif
                if( res == 0 ) m_temp_dirs.push_back( part );  //Remember the created subdirectories
                if( ( res == -1 ) && ( errno == ENOENT ) ) return false;
            }
            currPointer++;
        }
        return true;
    }

#if ! defined( _WIN32 ) && defined( _UNICODE )      //Linux with Unicode
    std::string PathManager::PathEncode( const wchar_t * Path )
    {
        mbstate_t MbState;
        const wchar_t * Ptr = Path;
        //mbrlen( NULL, 0, & MbState );
        std::memset( & MbState, 0, sizeof( MbState ) );
        std::size_t BufSize = wcsrtombs( NULL, & Ptr, 0, & MbState );
        if( ( BufSize == static_cast<std::size_t>( -1 ) ) || ( BufSize == 0 ) ) return "";
        std::vector<char> MbVector( BufSize + 1, 0 );
        wcsrtombs( MbVector.data(), & Ptr, BufSize, & MbState );
        MbVector[ BufSize ] = '\0';
        return MbVector.data();
    }
#endif

}
