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

#include <stdio.h>
#include <stdlib.h>
#include <inttypes.h>
#include <algorithm>
#include "clsRGBColorRecord.h"
#include <string.h>
namespace SimpleXlsx
{


clsRGBColorRecord::clsRGBColorRecord()
{
    //ctor
    snprintf(colstr,sizeof(colstr),"%02X%02X%02X",0,0,0);

}

clsRGBColorRecord::~clsRGBColorRecord()
{
    //dtor
}

clsRGBColorRecord::clsRGBColorRecord(const clsRGBColorRecord& other)
{
    //copy ctor
    Copy(other);
}

clsRGBColorRecord& clsRGBColorRecord::operator=(const clsRGBColorRecord& rhs)
{
    if (this == &rhs) return *this; // handle self assignment
    Copy(rhs);
    return *this;
}
void clsRGBColorRecord::Copy(const clsRGBColorRecord& p){
   //std::move(p.colstr,p.colstr+sizeof(p.colstr),colstr);
   memcpy(colstr,p.colstr,7*sizeof(char));
 };

 void clsRGBColorRecord::Set(const char (&sharpn)[7]){
   //std::move(sharpn,sharpn+sizeof(sharpn),colstr);
   memcpy(colstr,sharpn,7*sizeof(char));
   };

 void clsRGBColorRecord::Set(const unsigned char r, const unsigned char g, const unsigned char b){
  snprintf(colstr,sizeof(colstr),"%02X%02X%02X",r,g,b);
  };
  void clsRGBColorRecord::Set(const unsigned char gs){
   snprintf(colstr,sizeof(colstr),"%02X%02X%02X",gs,gs,gs);
   };
  void clsRGBColorRecord::Set(const double dgs){
   unsigned char gs=(255*dgs)*0.01;
   Set(gs);
  };
};
