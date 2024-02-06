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

#ifndef XLSXCOLORLIB_H_INCLUDED
#define XLSXCOLORLIB_H_INCLUDED
#include <map>
#include "clsRGBColorRecord.h"
namespace SimpleXlsx
{
 class
   XLSXColorLib{
   std::map<std::string, clsRGBColorRecord> lib;
   public:
     XLSXColorLib(){};
     virtual ~XLSXColorLib();
     void AddColor(const char * id, const clsRGBColorRecord & cl){ lib[id]=cl;  };
     const char * GetColor(const char * id) { return lib[id].Get(); };
     void Clear() { lib.clear(); };
   };
 extern void make_grayscale10(XLSXColorLib & xlib);
 extern void make_excell_like_named_colors(XLSXColorLib & xlib);
}
#endif // XLSXCOLORLIB_H_INCLUDED
