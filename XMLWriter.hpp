/*

Copyright (c) 2017, Alexandr Belyak ( ProgrammerAlex@bk.ru )
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

#ifndef XMLWRITER_H
#define XMLWRITER_H

#include <cassert>
#include <iostream>
#include <ostream>
#include <stack>
#include <string>

#include "PathManager.hpp"
#include "UTF8Encoder.hpp"

#include "Xlsx/SimpleXlsxDef.h"


namespace SimpleXlsx
{

    class XMLWriter
    {
        public:
            inline XMLWriter( const _tstring & FileName ) : _TagOpen( false ), _SelfClosed( true )
            {
                Init( PathManager::PathEncode( FileName ) );
            }

            inline ~XMLWriter()
            {
                EndAll();
                DebugCheckIsLightTagOpened();
            }

            inline bool IsOk() const
            {
                return _OStream.is_open();
            }


            //RAW output to the Stream
            /*template<class _T>
            inline XMLWriter & operator <<( const _T & Value )
            {
                _OStream << Value;
                return * this;
            }*/

            //Returns the current precision of floating point
            std::streamsize GetFloatPrecision()
            {
                return _OStream.precision();
            }

            //Set the current precision for floating point.
            //Returns the precision before the call this function.
            std::streamsize SetFloatPrecision( std::streamsize NewPrecision )
            {
                return _OStream.precision( NewPrecision );
            }

            //Light version without using stack of Tag Names.
            //No internal elements/tags or content string. Only attributes accepted.
            //Must be used with EndL.
            inline XMLWriter & TagL( const char * TagName )
            {
                assert( TagName != NULL );
                DebugCheckAndIncLightTag();
                CloseOpenedTag();
                _OStream << '<' << TagName;
                _TagOpen = true;
                _SelfClosed = false;
                return * this;
            }

            //Light version without using stack of Tag Names.
            //No internal elements/tags or content string. Only attributes accepted.
            //Must be used with TagL.
            inline XMLWriter & EndL()
            {
                DebugCheckAndDecLightTag();
                _OStream << "/>";
                _TagOpen = false;
                return * this;
            }

            inline XMLWriter & Tag( const char * TagName )
            {
                assert( TagName != NULL );
                CloseOpenedTag();
                DebugCheckIsLightTagOpened();
                _OStream << '<' << TagName;
                _TagOpen = true;
                _SelfClosed = true;
                _Tags.push( TagName );
                return * this;
            }

            inline XMLWriter & End( const char * TagName = NULL )
            {
                assert( ! _Tags.empty() );
                DebugCheckIsLightTagOpened();
                if( _SelfClosed ) _OStream << "/>";
                else _OStream << "</" << _Tags.top() << '>';
#ifndef NDEBUG
                if( TagName != NULL )
                {
                    if( _Tags.top() != TagName )
                        std::cerr << "Wrong TagName for End: " << TagName << ". Wanted: " << _Tags.top() << std::endl;
                    assert( _Tags.top() == TagName );
                }
#else
                ( void )TagName;
#endif
                _Tags.pop();
                _TagOpen = false;
                _SelfClosed = false;
                return * this;
            }

            inline XMLWriter & EndAll()
            {
                while( ! _Tags.empty() )
                    this->End();
                return * this;
            }

            inline XMLWriter & TagOnlyContent( const char * TagName, const char * ContentString )
            {
                assert( TagName != NULL );
                CloseOpenedTag();
                DebugCheckIsLightTagOpened();
                _OStream << '<' << TagName << '>';
                WriteStringEscape( ContentString );
                _OStream << "</" << TagName << '>';
                _SelfClosed = false;
                return * this;
            }

            inline XMLWriter & TagOnlyContent( const char * TagName, const std::string  & ContentString )
            {
                return TagOnlyContent( TagName, ContentString.c_str() );
            }

            //TagOnlyContent() template for all streamable types
            template <typename _T>
            inline XMLWriter & TagOnlyContent( const char * TagName, _T Value )
            {
                CloseOpenedTag();
                DebugCheckIsLightTagOpened();
                _OStream << '<' << TagName << '>' << Value << "</" << TagName << '>';
                _SelfClosed = false;
                return *this;
            }

            template <typename _T>
            inline XMLWriter & TagOnlyContent( const char * TagName, _T * Value );

            //Fast writing an attribute for the current Tag
            inline XMLWriter & Attr( const char * AttrName, const char * String )
            {
                assert( AttrName != NULL );
                assert( _TagOpen );
                _OStream << ' ' << AttrName << "=\"" << String << '"';
                return * this;
            }

            //Write an attribute for the current Tag
            inline XMLWriter & Attr( const char * AttrName, char * String )
            {
                return AttrInt( AttrName, String );
            }

            //Attr() overload for std::string type
            inline XMLWriter & Attr( const char * AttrName, const std::string & String )
            {
                return AttrInt( AttrName, String.c_str() );
            }

            //Attr() function template for all streamable types
            template <typename _T>
            inline XMLWriter & Attr( const char * AttrName, _T Value )
            {
                assert( AttrName != NULL );
                assert( _TagOpen );
                _OStream.put( ' ' );
                _OStream << AttrName << "=\"" << Value;
                _OStream.put( '"' );
                return *this;
            }

            template <typename _T>
            inline XMLWriter & Attr( const char * AttrName, _T * Value );

            //Write an content for the current Tag
            inline XMLWriter & Cont( const char * String )
            {
                CloseOpenedTag();
                DebugCheckIsLightTagOpened();
                WriteStringEscape( String );
                _SelfClosed = false;
                return * this;
            }

            //Write an content for the current Tag
            inline XMLWriter & Cont( char * String )
            {
                return Cont( static_cast<const char *>( String ) );
            }

            //Content() overload for std::string type
            inline XMLWriter & Cont( const std::string & String )
            {
                return Cont( String.c_str() );
            }

            //Content() template for all streamable types
            template <typename _T>
            inline XMLWriter & Cont( _T Value )
            {
                CloseOpenedTag();
                DebugCheckIsLightTagOpened();
                _OStream << Value;
                _SelfClosed = false;
                return *this;
            }

            template <typename _T>
            inline XMLWriter & Cont( _T * Value );

#ifdef _UNICODE
            //TagOnlyContent() overload for std::wstring type
            inline XMLWriter & TagOnlyContent( const char * TagName, const std::wstring  & ContentString )
            {
                return TagOnlyContent( TagName, UTF8Encoder::From_wstring( ContentString ).c_str() );
            }
            //Attr() overload for std::wstring type
            inline XMLWriter & Attr( const char * AttrName, const std::wstring & String )
            {
                return AttrInt( AttrName, UTF8Encoder::From_wstring( String ).c_str() );
            }
            //Cont() overload for std::wstring type
            inline XMLWriter & Cont( const std::wstring & String )
            {
                return Cont( UTF8Encoder::From_wstring( String ).c_str() );
            }
#endif

        private:
            bool _TagOpen, _SelfClosed;
            std::ofstream       _OStream;
            std::stack<std::string> _Tags;

            inline void Init( const std::string & FileName )
            {
                assert( ! FileName.empty() );
#ifndef NDEBUG
                LightTagCounter = 0;
#endif
                _OStream.open( FileName.c_str(), std::ios_base::out );
                _OStream.imbue( std::locale( "C" ) );
                _OStream << "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\n";
            }

            inline void CloseOpenedTag()
            {
                if( ! _TagOpen ) return;
                _OStream.put( '>' );
                _TagOpen = false;
            }

            inline void WriteStringEscape( const char * String )
            {
                for( ; * String; String++ )
                    switch( * String )
                    {
                        case '&'    :   _OStream << "&amp;";        break;
                        case '<'    :   _OStream << "&lt;";         break;
                        case '>'    :   _OStream << "&gt;";         break;
                        case '\''   :   _OStream << "&apos;";       break;
                        case '"'    :   _OStream << "&quot;";       break;
                        default     :   _OStream.put( * String );   break;
                    }
            }

            //Write an attribute for the current Tag
            inline XMLWriter & AttrInt( const char * AttrName, const char * String )
            {
                assert( AttrName != NULL );
                assert( _TagOpen );
                _OStream << ' ' << AttrName << "=\"";
                WriteStringEscape( String );
                _OStream.put( '"' );
                return * this;
            }

            //Debug version for checking TagL and EndL
#ifdef NDEBUG
            inline void DebugCheckAndIncLightTag()      {}
            inline void DebugCheckAndDecLightTag()      {}
            inline void DebugCheckIsLightTagOpened()    {}
#else
            intptr_t    LightTagCounter;

            inline void DebugCheckAndIncLightTag()
            {
                assert( LightTagCounter == 0 );
                LightTagCounter++;
            }

            inline void DebugCheckAndDecLightTag()
            {
                assert( LightTagCounter == 1 );
                LightTagCounter--;
            }
            inline void DebugCheckIsLightTagOpened()
            {
                assert( LightTagCounter == 0 );
            }
#endif
    };
}

#endif // XMLWRITER_H
