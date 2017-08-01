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

#ifndef UTF8ENCODER_H
#define UTF8ENCODER_H

#include <assert.h>
#include <stdint.h>
#include <string>

class UTF8Encoder
{
    public:
        static inline std::string From_wchar_t( int32_t Code )
        {
            assert( ( Code >= 0 ) && "Code can not be negative." );
            if( Code <= 0x7F ) return std::string( 1, static_cast<char>( Code ) );
            else
            {
                unsigned char FirstByte = 0xC0;          //First byte with mask
                unsigned char Buffer[ 7 ] = { 0 };       //Buffer for UTF-8 bytes, null-ponter string
                char * Ptr = reinterpret_cast<char *>( & Buffer[ 5 ] ); //Ptr to Buffer, started from end
                unsigned char LastValue = 0x1F;          //Max value for implement to the first byte
                while( true )
                {
                    * Ptr = static_cast<char>( ( Code & 0x3F ) | 0x80 );    //Making new value with format 10xxxxxx
                    Ptr--;                                  //Move Ptr to new position
                    Code >>= 6;                             //Shifting Code
                    if( Code <= LastValue ) break;          //if Code can set to FirstByte => break;
                    LastValue >>= 1;                        //Make less max value
                    FirstByte = ( FirstByte >> 1 ) | 0x80;  //Modify first byte
                }
                * Ptr = static_cast<char>( FirstByte | Code );  //Making first byte of UTF-8 sequence
                return std::string( Ptr );
            }
        }

#ifdef _UNICODE
        static inline std::string From_wstring( const std::wstring & Source )
        {
            std::string Result;
            for( std::wstring::const_iterator iter = Source.begin(); iter != Source.end(); iter++ )
                Result.append( From_wchar_t( * iter ) );
            return Result;
        }
#endif
};


#endif // UTF8ENCODER_H
