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

#include "XLSXColorLib.h"
namespace SimpleXlsx
{
XLSXColorLib::~XLSXColorLib(){};

 void make_grayscale10(XLSXColorLib & xlib){
   xlib.Clear();

   xlib.AddColor("Black",clsRGBColorRecord((unsigned char)0));
   xlib.AddColor("Gray10%",clsRGBColorRecord(10.));
   xlib.AddColor("Gray20%",clsRGBColorRecord(20.));
   xlib.AddColor("Gray30%",clsRGBColorRecord(30.));
   xlib.AddColor("Gray40%",clsRGBColorRecord(40.));
   xlib.AddColor("Gray50%",clsRGBColorRecord(50.));
   xlib.AddColor("Gray60%",clsRGBColorRecord(60.));
   xlib.AddColor("Gray70%",clsRGBColorRecord(70.));
   xlib.AddColor("Gray80%",clsRGBColorRecord(80.));
   xlib.AddColor("Gray90%",clsRGBColorRecord(90.));
   xlib.AddColor("White",clsRGBColorRecord((unsigned char)255));
  }
void make_excell_like_named_colors(XLSXColorLib & xlib){
  xlib.Clear();
  xlib.AddColor("Black",clsRGBColorRecord((unsigned char)0));
  xlib.AddColor("Brown",clsRGBColorRecord("993300"));
  xlib.AddColor("Olive Green",clsRGBColorRecord("333300"));
  xlib.AddColor("Dark Green",clsRGBColorRecord("003300"));
  xlib.AddColor("Dark Teal",clsRGBColorRecord("003366"));
  xlib.AddColor("Dark Blue",clsRGBColorRecord("000080"));
  xlib.AddColor("Indigo",clsRGBColorRecord("333399"));
  xlib.AddColor("Dark Red",clsRGBColorRecord("800000"));
  xlib.AddColor("Orange",clsRGBColorRecord("FF6600"));
  xlib.AddColor("Dark Yellow",clsRGBColorRecord("808000"));
  xlib.AddColor("Green",clsRGBColorRecord("008000"));
  xlib.AddColor("Teal",clsRGBColorRecord("008080"));
  xlib.AddColor("Blue",clsRGBColorRecord("0000FF"));
  xlib.AddColor("Blue-Gray",clsRGBColorRecord("666699"));
  xlib.AddColor("Red",clsRGBColorRecord("FF0000"));
  xlib.AddColor("Light Orange",clsRGBColorRecord("FF9900"));
  xlib.AddColor("Lime",clsRGBColorRecord("99CC00"));
  xlib.AddColor("Sea Green",clsRGBColorRecord("339966"));
  xlib.AddColor("Aqua",clsRGBColorRecord("33CCCC"));
  xlib.AddColor("Light Blue",clsRGBColorRecord("3366FF"));
  xlib.AddColor("Violet",clsRGBColorRecord("800080"));
  xlib.AddColor("Pink",clsRGBColorRecord("FF00FF"));
  xlib.AddColor("Gold",clsRGBColorRecord("FFCC00"));
  xlib.AddColor("Yellow",clsRGBColorRecord("FFFF00"));
  xlib.AddColor("Bright Green",clsRGBColorRecord("00FF00"));
  xlib.AddColor("Turquoise",clsRGBColorRecord("00FFFF"));
  xlib.AddColor("Sky Blue",clsRGBColorRecord("00CCFF"));
  xlib.AddColor("Plum",clsRGBColorRecord("993366"));
  xlib.AddColor("Rose",clsRGBColorRecord("FF99CC"));
  xlib.AddColor("Tan",clsRGBColorRecord("FFCC99"));
  xlib.AddColor("Light Yellow",clsRGBColorRecord("FFFF99"));
  xlib.AddColor("Light Green",clsRGBColorRecord("CCFFCC"));
  xlib.AddColor("Light Turquoise",clsRGBColorRecord("CCFFFF"));
  xlib.AddColor("Pale Blue",clsRGBColorRecord("99CCFF"));
  xlib.AddColor("Lavender",clsRGBColorRecord("CC99FF"));
  xlib.AddColor("Periwinkle",clsRGBColorRecord("9999FF"));
  xlib.AddColor("Dark Purple",clsRGBColorRecord("660066"));
  xlib.AddColor("Coral",clsRGBColorRecord("FF8080"));
  xlib.AddColor("Ocean Blue",clsRGBColorRecord("0066CC"));
  xlib.AddColor("Ice Blue",clsRGBColorRecord("CCCCFF"));
  xlib.AddColor("White",clsRGBColorRecord((unsigned char)255));
  xlib.AddColor("Gray",clsRGBColorRecord((unsigned char)127));
 };
};
