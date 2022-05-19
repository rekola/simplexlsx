/*
  Addon to SimpleXlsxWriter for the convenience of working with colors,
  Copyright (C) 2017-2022 E.Naumovich

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

#ifndef CLSRGBCOLORRECORD_H
#define CLSRGBCOLORRECORD_H
#include <string>
namespace SimpleXlsx
{
class clsRGBColorRecord
{
    public:
        clsRGBColorRecord();
        clsRGBColorRecord( const unsigned char gs) { Set(gs);};
        clsRGBColorRecord( const unsigned char r, const unsigned char g, const unsigned char b) { Set(r,g,b);};
        clsRGBColorRecord( const double gs){ Set(gs);};
        clsRGBColorRecord( const char (&sharpn)[7]){ Set(sharpn);};
        virtual ~clsRGBColorRecord();
        clsRGBColorRecord(const clsRGBColorRecord& other);
        clsRGBColorRecord& operator=(const clsRGBColorRecord& other);
        const char * Get() const {return colstr;};
        void Set(const unsigned char r, const unsigned char g, const unsigned char b);
        void Set(const unsigned char chargs);
        void Set(const double gs);
        void Set(const char (&sharpn)[7]);
    protected:
        char colstr[7];
        void Copy(const clsRGBColorRecord& other);
    private:
};
};
#endif // CLSRGBCOLORRECORD_H
