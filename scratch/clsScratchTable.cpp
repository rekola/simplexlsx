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

#include <algorithm>
#include "clsScratchTable.h"


using namespace ScratchTable;
  clsCell::~clsCell(){};
//-------------------------------------------------
   void  clsCell::Copy(const clsCell& p){
        vtype=p.vtype;
        Style=p.Style;
        stError=p.stError;
        row=p.row;
        col=p.col;
        priority=p.priority;
        if(vtype<STRV){
          memcpy(&val,&p.val,sizeof(val));
          return;
        }
        if(vtype==STRV) {strv=p.strv; return;};
        if(vtype==WSTRV) {wstrv=p.wstrv;};
      };
//------------------------------
void clsCell::InsertTo(SimpleXlsx::CWorksheet & sheet)const{
 switch(vtype){
  case INTV:   sheet.AddCell(val.iv,Style);  break;
  case UINTV:  sheet.AddCell(val.uiv,Style); break;
  case DBLV:   sheet.AddCell(val.dbv,Style); break;
  case FLOV:   sheet.AddCell(val.flv,Style); break;
  case STRV:   sheet.AddCell(strv,Style);    break;
  case LINTV:  sheet.AddCell(val.liv,Style); break;
  case ULINTV: sheet.AddCell(val.uliv,Style);break;
  case WSTRV:  sheet.AddCell(wstrv,Style);   break;
#ifdef _WIN32
  case WTV:    sheet.AddCell(SimpleXlsx::CellDataTime(val.wtv,Style)); break;
#endif
  case UTV:    sheet.AddCell(SimpleXlsx::CellDataTime(val.utv,Style)); break;
  case EMPTYV: if(Style) {sheet.AddCell("",Style);} else sheet.AddEmptyCells(1); break;
 default:sheet.AddCell("?wrong cell type",stError); break;
 }
 }
 ///------------------------------------------------------
   clsScratchTable::~clsScratchTable(){};
//---------------------------------------------
  bool clsScratchTable::Add( clsCell & v){
   v.priority=-static_cast<int64_t>(counter); // set priority to control overwrite
   Cells.push_back(v);
   counter++; // Not a Cells.size()... !!!!
   return true;
  }
//-----------------------------------------------
 bool clsScratchTable::Del(const size_t row, const size_t col){
    auto rs=false;
    for(auto & cl:Cells){
      if(cl.row==row&&cl.col==col){
       cl.priority=1; // mark cell to delete
       rs=true;
       continue;
      }
     }
    counter++;
   return rs;
  }
//--------------------------------------------------
 void clsScratchTable::Sort(){
   std::sort(Cells.begin(),Cells.end());
   if(counter==Cells.size()) return; // cells were only appended, no overwrite or delete operations
  // remove overwitten or deleted cells
   std::vector<clsCell> cells;
   cells.push_back(Cells.front());
   for(auto cl=Cells.cbegin()+1;cl!=Cells.cend();cl++){
   if(cl->check(cells.back())&&cl->priority<1)
        cells.push_back(*cl);
    }
   Cells=cells;
   };
//---------------------------------------------------
 size_t clsScratchTable::FindRowEntry(const size_t row, const size_t from){
   size_t cmax=Cells.size(); // default value- not found/empty
   for(auto cl=Cells.cbegin()+from;cl!=Cells.cend();cl++){
    if(cl->row<row) continue;
    if(cl->row==row) {cmax=cl-Cells.cbegin(); break;};
    if(cl->row>row) break; // row is empty, so it was not found
   }
   return cmax;
   };
//--------------------------------------------------------------

size_t clsScratchTable::FindLastCol(const size_t row, const size_t from){
 for(auto cl=Cells.cbegin()+from;cl!=Cells.cend();cl++)
   if(cl->row!=row) return cl-Cells.cbegin()-1;
 return Cells.size()-1;
};
//-------------------------------------------------------------
size_t clsScratchTable::FindCell(const size_t row,const size_t col, const size_t from){
  for(auto cl=Cells.cbegin()+from;cl!=Cells.cend();cl++){
   if(cl->row!=row) break;
   if(cl->row==row&&cl->col==col) return cl-Cells.cbegin();
  }
 return Cells.size();
};
//-------------------------------------------------------------

 size_t clsScratchTable::InsertTo(SimpleXlsx::CWorksheet & sheet){
   if(Cells.size()==0) return -1; // empty storage;

   Sort();
   size_t minrow=Cells.front().row; // catch first row index
   size_t maxrow=Cells.back().row; // catch last row index;
   size_t clrow=0;// index of the last treated row in Cells
 if(minrow!=0) sheet.AddEmptyRows(minrow);
   size_t cnt_empty_rows=0; // counter for delayed insert
 for(size_t row=minrow;row<=maxrow;++row){
    size_t cl0row=FindRowEntry(row,clrow); // index of row to treat in Cells
    if(cl0row==Cells.size()){
     cnt_empty_rows++;  // row not exists, count it to insert later
     continue;
    }
   if(cnt_empty_rows) {
      sheet.AddEmptyRows(cnt_empty_rows);
      cnt_empty_rows=0;
      }

   sheet.BeginRow();
     size_t clmincol=cl0row;
     size_t clmaxcol=FindLastCol(row,cl0row);
     size_t mincol=Cells[clmincol].col;
     size_t maxcol=Cells[clmaxcol].col;
     size_t cnt_empty_cl=0; //  counter for delayed insert
     if(mincol!=0) sheet.AddEmptyCells(mincol);
     for(auto col=mincol;col<=maxcol;++col){
       size_t c0cl=FindCell(row,col,clmincol);
       if(c0cl==Cells.size()){
         cnt_empty_cl++; // row not exists, count it to insert later
         continue;
       }
       if(cnt_empty_cl) {
         sheet.AddEmptyCells(cnt_empty_cl);
         cnt_empty_cl=0;
        }
       Cells.at(c0cl).InsertTo(sheet);
       clmincol=c0cl;
     }
   sheet.EndRow();
   clrow=cl0row;
 };
   return  Cells.size();
   }

