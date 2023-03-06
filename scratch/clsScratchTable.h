/*
  Addon to SimpleXlsxWriter like a scratch sheet with matrix-type access to
  cells (writing only),
  Copyright (C) 2022 E.Naumovich

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

#ifndef HEADER_CLSSCRATCHTABLE
#define HEADER_CLSSCRATCHTABLE
#include <string.h>
#include "../Xlsx/Workbook.h"


namespace ScratchTable{
struct clsCell{
 enum valtype {INTV=0,UINTV=1,DBLV=2,FLOV=3,WTV=4,UTV=5,LINTV=6,ULINTV=7,STRV=20,WSTRV=21,EMPTYV=30, ERRORV=40} vtype;
 enum { maxRowsN=1048576,maxColsN=16384}; //  limits from the internet.
 size_t Style,stError,row,col;

 int64_t priority; // if> 0, cell will be deleted; if <1,  cell with lower value is a last version, all other at that coordinates will deleted. Deletion will be done in Sort().
 union {
         int iv;
         unsigned int uiv;
         double dbv;
         float flv;
         int64_t liv;
         uint64_t uliv;
         time_t  utv;
#ifdef _WIN32
         SYSTEMTIME  wtv;
#endif
           } val;
 std::string strv;
 std::wstring wstrv;
 void InsertTo(SimpleXlsx::CWorksheet & sheet) const;

//----------------------------

 clsCell &  SetLong( const size_t r, const size_t c,const int64_t &  _v, const size_t st=0){
   row=(r<maxRowsN)?r:maxRowsN-1;
   col=(c<maxColsN)?c:maxColsN-1;
   Style=st;
   vtype=LINTV;
   val.liv=_v;
   return *this;
  }
 //----------------------------
 clsCell &  SetLong( const size_t r, const size_t c,const uint64_t &  _v, const size_t st=0){
   row=(r<maxRowsN)?r:maxRowsN-1;
   col=(c<maxColsN)?c:maxColsN-1;
   Style=st;
   vtype=ULINTV;
   val.uliv=_v;
   return *this;
  }
//----------------------------
 template <class T>
  clsCell &  Set( const size_t r, const size_t c,const T & _v, const size_t st=0){
   row=(r<maxRowsN)?r:maxRowsN-1; //
   col=(c<maxColsN)?c:maxColsN-1;  //
   Style=st;
   return Set(_v);
  }
//----------------------------

 clsCell & Set(const int &          _v){ vtype=INTV;  val.iv=_v;  return *this;};
 clsCell & Set(const unsigned int & _v){ vtype=UINTV; val.uiv=_v; return *this;};
 clsCell & Set(const double &       _v){ vtype=DBLV;  val.dbv=_v; return *this;};
 clsCell & Set(const float &        _v){ vtype=FLOV;  val.flv=_v; return *this;};
 clsCell & Set(const time_t &       _v){ vtype=UTV;   memcpy((void*)&val.utv,&_v, sizeof(_v));return *this;};
#ifdef _WIN32
 clsCell & Set(const SYSTEMTIME &   _v){ vtype=WTV;   memcpy((void*)&val.wtv,&_v, sizeof(_v));return *this;};
#endif
 clsCell & Set(const std::string &  _v){ vtype=STRV;  strv=_v;    return *this;};
 clsCell & Set(const std::wstring & _v){ vtype=WSTRV; wstrv=_v;   return *this;};
 clsCell & Set(const std::nullptr_t _v){ vtype=EMPTYV;strv.clear();wstrv.clear(); memset(&val,0,sizeof(val));(void)_v;return *this;};

//----------------------------
   clsCell():vtype(EMPTYV),Style(0),stError(0),row(0),col(0) {memset(&val,0,sizeof(val));};

   virtual ~clsCell();

   clsCell(const clsCell& other){
               Copy(other);
            };

   clsCell& operator=(const clsCell& rhs){
                if (this == &rhs) return *this; // handle self assignment
                Copy(rhs);
                return *this;
            };

    bool operator==(const clsCell& rhs) const { return row==rhs.row&&col==rhs.col&&priority==rhs.priority;};
    bool operator<(const clsCell& rhs)  const { return (row!=rhs.row)?row<rhs.row:(col!=rhs.col?col<rhs.col:priority<rhs.priority);};
    inline bool check(const clsCell& sample) const { return !(row==sample.row&&col==sample.col&&priority<sample.priority); } //  to control overwite conditions

    protected:
      void  Copy(const clsCell& p);
};


///------------------------------------------------------

class clsScratchTable{
 std::vector<clsCell> Cells;
 size_t counter; // it counts operation on the Cells. Here possible might be overflow theoretically for extremely big tables in 32b environment.
 size_t FindRowEntry(const size_t row, const size_t from);
 size_t FindLastCol(const size_t row, const size_t from);
 size_t FindCell(const size_t row,const size_t col, const size_t from);
public:
  bool Add( clsCell & v);
  bool Del(const size_t row, const size_t col);
  void Sort(); // row-col-priority;  it also deletes ovewritten and marked-to-delete cells. Used automatically
  size_t InsertTo(SimpleXlsx::CWorksheet & sheet);
  void Clear() {Cells.clear(); counter=0;};
  void GetDim(size_t &rw, size_t & cl) const;
  clsScratchTable(){counter=0; };
  virtual ~clsScratchTable();
};

};



#endif // header guard

